<#
.SYNOPSIS
    Reports last N months of Azure cost per subscription, for a single
    tenant, broken down by pricing model (Reservation vs On-Demand vs
    Savings Plan).

.DESCRIPTION
    Single, parameterized replacement for the old per-tenant copies of
    this script. Run it once per tenant (e.g. once for "AlternativeSoft",
    once for "AlternativeSoft.cloud") - that's still the right call if
    the two tenants need separate interactive sign-ins, but now the
    fix/maintenance work only has to happen in one place. Use
    Merge-CostReports.ps1 afterwards to combine the outputs from both
    runs into a single master report.

.PARAMETER TenantName
    Name or domain fragment to match against Get-AzTenant (e.g.
    "AlternativeSoft" or "AlternativeSoft.cloud"). If omitted, the
    script lists the tenants visible to the signed-in account and
    prompts you to pick one.

.PARAMETER MonthsBack
    How many months back to report, inclusive of the current month.
    Default 12.

.PARAMETER OutputDir
    Where to write the CSVs. Default is the current directory.

.NOTES
    Requires: Az.Accounts.
    Prerequisites: Cost Management Reader (or Reader) role on each subscription.

    READ-ONLY GUARANTEE:
    Only calls the Cost Management "query" (POST) endpoint, which is a
    read-only reporting query. No resource, budget, or billing config is
    created, modified, or deleted. The only local write is the CSV export.

.EXAMPLE
    ./azure-cost-report.ps1 -TenantName "AlternativeSoft"
    ./azure-cost-report.ps1 -TenantName "AlternativeSoft.cloud"
#>

param(
    [string]$TenantName,

    [int]$MonthsBack = 12,

    [string]$OutputDir = "."
)

# ------------------------------------------------------------------
# 1. Connect
# ------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    Install-Module -Name Az.Accounts -Scope CurrentUser -Force -AllowClobber
}
if (-not (Get-AzContext)) {
    Connect-AzAccount | Out-Null
}

# ------------------------------------------------------------------
# 2. Resolve the tenant
# ------------------------------------------------------------------
if ($TenantName) {
    $tenant = Get-AzTenant | Where-Object { $_.Name -like "*$TenantName*" -or $_.DefaultDomain -like "*$TenantName*" } | Select-Object -First 1

    if (-not $tenant) {
        Write-Error "Tenant matching '$TenantName' not found. Run 'Get-AzTenant' to see exact names/domains."
        return
    }
} else {
    $allTenants = @(Get-AzTenant)

    if (-not $allTenants) {
        Write-Error "No tenants visible to the signed-in account. Run 'Get-AzTenant' to check."
        return
    }

    if ($allTenants.Count -eq 1) {
        $tenant = $allTenants[0]
        Write-Host "Only one tenant visible - using it automatically." -ForegroundColor Cyan
    } else {
        Write-Host "`nNo -TenantName supplied. Choose a tenant:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $allTenants.Count; $i++) {
            $t = $allTenants[$i]
            $label = if ($t.Name) { $t.Name } else { $t.DefaultDomain }
            Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $label, $t.Id)
        }

        do {
            $selection = Read-Host "Enter number (1-$($allTenants.Count))"
            $selectionOk = $selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $allTenants.Count
            if (-not $selectionOk) {
                Write-Host "Invalid selection, try again." -ForegroundColor Red
            }
        } until ($selectionOk)

        $tenant = $allTenants[[int]$selection - 1]
    }
}

Write-Host "Using tenant: $($tenant.Name) [$($tenant.Id)]" -ForegroundColor Cyan

try {
    Connect-AzAccount -Tenant $tenant.Id -ErrorAction Stop | Out-Null
} catch {
    Write-Error "Could not authenticate to tenant $($tenant.Name): $($_.Exception.Message)"
    return
}

# ------------------------------------------------------------------
# 3. Time period
# ------------------------------------------------------------------
$today = Get-Date
$fromDate = (Get-Date -Day 1).AddMonths(-($MonthsBack - 1))
$fromStr = $fromDate.ToString("yyyy-MM-ddT00:00:00Z")
$toStr   = $today.ToString("yyyy-MM-ddT23:59:59Z")

Write-Host "Querying cost from $fromStr to $toStr" -ForegroundColor Cyan

# ------------------------------------------------------------------
# 4. Helper: query cost for one subscription, with full error surfacing
# ------------------------------------------------------------------
function Get-SubscriptionMonthlyCost {
    param(
        [string]$SubscriptionId,
        [string]$SubscriptionName,
        [string]$FromStr,
        [string]$ToStr
    )

    $tokenObj = Get-AzAccessToken -ResourceUrl "https://management.azure.com/"
    if ($tokenObj.Token -is [System.Security.SecureString]) {
        # Newer Az.Accounts versions return the token as a SecureString by
        # default. Using it directly in a header sends the literal string
        # "System.Security.SecureString" instead of the token, causing
        # "InvalidAuthenticationToken" with no other clue. Decode it here.
        $token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
        )
    } else {
        $token = $tokenObj.Token
    }
    $headers = @{
        Authorization  = "Bearer $token"
        "Content-Type" = "application/json"
    }

    $uri = "https://management.azure.com/subscriptions/$SubscriptionId/providers/Microsoft.CostManagement/query?api-version=2023-11-01"

    $body = @{
        type      = "ActualCost"
        timeframe = "Custom"
        timePeriod = @{ from = $FromStr; to = $ToStr }
        dataset = @{
            granularity = "Monthly"
            aggregation = @{ totalCost = @{ name = "Cost"; function = "Sum" } }
            grouping    = @(@{ type = "Dimension"; name = "PricingModel" })
        }
    } | ConvertTo-Json -Depth 10

    try {
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ErrorAction Stop
        return $resp
    } catch {
        Write-Warning "Cost query failed for subscription '$SubscriptionName' ($SubscriptionId):"
        if ($_.ErrorDetails.Message) {
            Write-Warning $_.ErrorDetails.Message
        } else {
            Write-Warning $_.Exception.Message
        }
        return $null
    }
}

# ------------------------------------------------------------------
# 5. Loop subscriptions in this tenant
# ------------------------------------------------------------------
$subs = Get-AzSubscription -TenantId $tenant.Id -ErrorAction SilentlyContinue
if (-not $subs) {
    Write-Error "No subscriptions visible under tenant $($tenant.Name)."
    return
}

Write-Host "`nSubscriptions found:" -ForegroundColor Cyan
$subs | Format-Table Name, Id -AutoSize

$flatRows = @()

# Candidate names for the date/period column. The Cost Management query
# API does not always come back with a column literally named
# "UsageDate" for Monthly granularity - it can be "BillingMonth" or
# "Date" depending on API version / grouping. Trying several names (and
# failing LOUDLY, not silently) avoids the old bug where a -1 IndexOf
# wrapped around to $row[-1] (the last column, Currency) and quietly
# produced a garbage "month" like "0000-0G" for every row.
$dateColumnCandidates = @("UsageDate", "BillingMonth", "Date")

foreach ($sub in $subs) {
    Write-Host "`nSubscription: $($sub.Name) [$($sub.Id)]" -ForegroundColor Yellow

    try {
        Set-AzContext -Subscription $sub.Id -Tenant $tenant.Id -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not set context: $($_.Exception.Message)"
        continue
    }

    $result = Get-SubscriptionMonthlyCost -SubscriptionId $sub.Id -SubscriptionName $sub.Name -FromStr $fromStr -ToStr $toStr
    if (-not $result -or -not $result.properties.rows) {
        Write-Warning "No cost data returned for $($sub.Name)."
        continue
    }

    $columns = $result.properties.columns.name
    Write-Host "Columns returned: $($columns -join ', ')" -ForegroundColor DarkGray

    $costIdx         = [array]::IndexOf($columns, "Cost")
    $pricingModelIdx = [array]::IndexOf($columns, "PricingModel")
    $currencyIdx     = [array]::IndexOf($columns, "Currency")

    $dateIdx = -1
    $dateColumnName = $null
    foreach ($candidate in $dateColumnCandidates) {
        $idx = [array]::IndexOf($columns, $candidate)
        if ($idx -ge 0) {
            $dateIdx = $idx
            $dateColumnName = $candidate
            break
        }
    }

    if ($dateIdx -lt 0) {
        Write-Error "Could not find a date/period column (tried: $($dateColumnCandidates -join ', ')) for subscription '$($sub.Name)'. Actual columns: $($columns -join ', '). Skipping this subscription rather than silently mis-bucketing its costs."
        continue
    }
    if ($costIdx -lt 0) {
        Write-Error "Could not find a 'Cost' column for subscription '$($sub.Name)'. Actual columns: $($columns -join ', '). Skipping this subscription."
        continue
    }

    foreach ($row in $result.properties.rows) {
        $usageDateStr = $row[$dateIdx].ToString().PadLeft(8, '0')
        $monthKey = $usageDateStr.Substring(0,4) + "-" + $usageDateStr.Substring(4,2)

        $flatRows += [PSCustomObject]@{
            TenantName       = $tenant.Name
            SubscriptionName = $sub.Name
            SubscriptionId   = $sub.Id
            Month            = $monthKey
            PricingModel     = if ($pricingModelIdx -ge 0) { $row[$pricingModelIdx] } else { "Unknown" }
            Cost             = [double]$row[$costIdx]
            Currency         = if ($currencyIdx -ge 0) { $row[$currencyIdx] } else { "" }
        }
    }
}

if (-not $flatRows) {
    Write-Error "No cost data collected for any subscription in this tenant."
    return
}

# ------------------------------------------------------------------
# 6. Pivot to wide format: one row per subscription, columns per month
# ------------------------------------------------------------------
$allMonths = $flatRows | Select-Object -ExpandProperty Month -Unique | Sort-Object
$subKeys = $flatRows | Select-Object TenantName, SubscriptionName, SubscriptionId -Unique

$pivotedRows = @()
foreach ($key in $subKeys) {
    $rowObj = [ordered]@{
        TenantName       = $key.TenantName
        SubscriptionName = $key.SubscriptionName
        SubscriptionId   = $key.SubscriptionId
    }
    foreach ($month in $allMonths) {
        $monthRows = $flatRows | Where-Object {
            $_.SubscriptionId -eq $key.SubscriptionId -and $_.Month -eq $month
        }
        $reservationCost = ($monthRows | Where-Object { $_.PricingModel -eq "Reservation" } | Measure-Object -Property Cost -Sum).Sum
        $onDemandCost     = ($monthRows | Where-Object { $_.PricingModel -eq "OnDemand" } | Measure-Object -Property Cost -Sum).Sum
        $savingsPlanCost  = ($monthRows | Where-Object { $_.PricingModel -eq "SavingsPlan" } | Measure-Object -Property Cost -Sum).Sum
        $totalCost        = ($monthRows | Measure-Object -Property Cost -Sum).Sum

        $rowObj["${month}_Reservation"] = [math]::Round(($reservationCost | ForEach-Object { $_ }), 2)
        $rowObj["${month}_OnDemand"]    = [math]::Round(($onDemandCost | ForEach-Object { $_ }), 2)
        $rowObj["${month}_SavingsPlan"] = [math]::Round(($savingsPlanCost | ForEach-Object { $_ }), 2)
        $rowObj["${month}_Total"]       = [math]::Round(($totalCost | ForEach-Object { $_ }), 2)
    }
    $pivotedRows += [PSCustomObject]$rowObj
}

# ------------------------------------------------------------------
# 7. Export
# ------------------------------------------------------------------
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$safeTenant = ($tenant.Name -replace '[^\w\-]', '_')

$flatFile = Join-Path $OutputDir "cost_${safeTenant}_flat_$timestamp.csv"
$flatRows | Sort-Object SubscriptionName, Month, PricingModel | Export-Csv -Path $flatFile -NoTypeInformation
Write-Host "`nExported long-format CSV: $flatFile" -ForegroundColor Cyan

$pivotFile = Join-Path $OutputDir "cost_${safeTenant}_pivoted_$timestamp.csv"
$pivotedRows | Export-Csv -Path $pivotFile -NoTypeInformation
Write-Host "Exported wide/pivoted CSV: $pivotFile" -ForegroundColor Cyan

Write-Host "`nDone. $($subKeys.Count) subscriptions, $($allMonths.Count) months covered." -ForegroundColor Green
