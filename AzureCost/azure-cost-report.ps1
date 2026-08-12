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
    Name, domain fragment, or Tenant ID GUID to match against Get-AzTenant (e.g.
    "AlternativeSoft", "AlternativeSoft.cloud", or "842ae988-60f7-4e75-9477-32c43f90daf5").
    If omitted, the script lists the tenants visible to the signed-in account and
    prompts you to pick one.

.PARAMETER Subscriptions
    One or more subscription names, IDs, or wildcard patterns to limit the query to.
    If omitted, all subscriptions in the selected tenant are queried.

.PARAMETER MonthsBack
    How many months back to report, inclusive of the current month.
    Default 12.

.PARAMETER OutputDir
    Where to write the CSVs. Default is "Output".

.NOTES
    Requires: Az.Accounts.
    Prerequisites: Cost Management Reader (or Reader) role on each subscription.

    READ-ONLY GUARANTEE:
    Only calls the Cost Management "query" (POST) endpoint, which is a
    read-only reporting query. No resource, budget, or billing config is
    created, modified, or deleted. The only local write is the CSV export.

.EXAMPLE
    pwsh AzureCost/azure-cost-report.ps1
    pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft"
    pwsh AzureCost/azure-cost-report.ps1 -TenantId "842ae988-60f7-4e75-9477-32c43f90daf5" -Subscription "ASWeb"
    pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft.cloud" -Subscriptions "CIBC","Wilshire" -MonthsBack 6 -OutputDir "Output"
#>

param(
    [Alias("Tenant", "TenantId")]
    [string]$TenantName,

    [Alias("Subscription", "SubscriptionName", "SubscriptionId", "SubscriptionNames", "SubscriptionIds")]
    [string[]]$Subscriptions,

    [Alias("Months", "Period")]
    [int]$MonthsBack = 12,

    [Alias("Start", "From")]
    [string]$StartDate,

    [Alias("End", "To")]
    [string]$EndDate,

    [string]$OutputDir = "Output"
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
    $tenant = Get-AzTenant | Where-Object {
        $_.Id -eq $TenantName -or
        $_.Name -eq $TenantName -or
        $_.DefaultDomain -eq $TenantName -or
        $_.Name -like "*$TenantName*" -or
        $_.DefaultDomain -like "*$TenantName*"
    } | Select-Object -First 1

    if (-not $tenant) {
        Write-Error "Could not find a tenant matching '$TenantName'."
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
    Set-AzContext -TenantId $tenant.Id -ErrorAction Stop | Out-Null
} catch {
    try {
        Connect-AzAccount -TenantId $tenant.Id -ErrorAction Stop | Out-Null
    } catch {
        Write-Error "Could not authenticate to tenant $($tenant.Name): $($_.Exception.Message)"
        return
    }
}

# ------------------------------------------------------------------
# 3. Time period & 12-month Azure API chunking
# ------------------------------------------------------------------
# Azure Cost Management query API rejects requests whose time span exceeds 1 year (12 months).
# If the requested period is > 12 months (e.g. 24, 36 months), we automatically partition
# the request into <= 12-month chunks and merge the returned monthly rows.
function Get-TimeChunks($overallStart, $overallEnd) {
    $chunks = @()
    $curStart = (Get-Date -Year $overallStart.Year -Month $overallStart.Month -Day 1)
    while ($curStart -le $overallEnd) {
        # 12-month window: start of curStart to last day of 12th month
        $chunkEnd = (Get-Date -Year $curStart.Year -Month $curStart.Month -Day 1).AddMonths(12).AddDays(-1)
        if ($chunkEnd -gt $overallEnd) {
            $chunkEnd = $overallEnd
        }
        $chunks += [PSCustomObject]@{
            FromStr = $curStart.ToString("yyyy-MM-ddT00:00:00Z")
            ToStr   = (Get-Date -Year $chunkEnd.Year -Month $chunkEnd.Month -Day $chunkEnd.Day).ToString("yyyy-MM-ddT23:59:59Z")
        }
        $curStart = (Get-Date -Year $curStart.Year -Month $curStart.Month -Day 1).AddMonths(12)
    }
    return $chunks
}

$today = Get-Date
if ($EndDate) {
    $parsedEnd = [DateTime]::MinValue
    if ([DateTime]::TryParse($EndDate, [ref]$parsedEnd)) {
        $today = $parsedEnd
    }
}

if ($StartDate) {
    $parsedStart = [DateTime]::MinValue
    if ([DateTime]::TryParse($StartDate, [ref]$parsedStart)) {
        $fromDate = (Get-Date -Year $parsedStart.Year -Month $parsedStart.Month -Day 1)
    } else {
        $fromDate = (Get-Date -Year $today.Year -Month $today.Month -Day 1).AddMonths(-($MonthsBack - 1))
    }
} else {
    $fromDate = (Get-Date -Year $today.Year -Month $today.Month -Day 1).AddMonths(-($MonthsBack - 1))
}

$timeChunks = Get-TimeChunks $fromDate $today

Write-Host "Querying historical cost from $($fromDate.ToString('yyyy-MM-01')) to $($today.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
if ($timeChunks.Count -gt 1) {
    Write-Host "Auto-partitioned into $($timeChunks.Count) period chunk(s) (Azure Cost API 1-year per query limit):" -ForegroundColor DarkGray
    foreach ($c in $timeChunks) {
        Write-Host "  - $($c.FromStr.Substring(0,10)) to $($c.ToStr.Substring(0,10))" -ForegroundColor DarkGray
    }
}

# ------------------------------------------------------------------
# 4. Helper: query cost for one subscription, with rate-limiting retry & full error surfacing
# ------------------------------------------------------------------
function Get-SubscriptionMonthlyCost {
    param(
        [string]$SubscriptionId,
        [string]$SubscriptionName,
        [string]$TenantId,
        [string]$FromStr,
        [string]$ToStr,
        [int]$MaxRetries = 5
    )

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

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $tokenObj = if ($TenantId) {
            Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -TenantId $TenantId -ErrorAction SilentlyContinue
        } else {
            Get-AzAccessToken -ResourceUrl "https://management.azure.com/" -ErrorAction SilentlyContinue
        }

        if (-not $tokenObj) {
            $tokenObj = Get-AzAccessToken -ResourceUrl "https://management.azure.com/"
        }

        if ($tokenObj.Token -is [System.Security.SecureString]) {
            # Cross-platform SecureString decoding (works on macOS, Linux, and Windows)
            $token = [System.Net.NetworkCredential]::new("", $tokenObj.Token).Password
        } elseif ($tokenObj.Token -is [string]) {
            $token = $tokenObj.Token
        } else {
            $token = [string]$tokenObj
        }

        $headers = @{
            Authorization  = "Bearer $token"
            "Content-Type" = "application/json"
        }

        try {
            $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ErrorAction Stop
            return $resp
        } catch {
            $errMsg = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
            $isRateLimited = ($errMsg -like "*429*" -or $errMsg -like "*Too many requests*")

            if ($isRateLimited -and $attempt -lt $MaxRetries) {
                # Attempt to extract Azure Retry-After header
                $retryAfter = 0
                try {
                    if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                        $ra = $_.Exception.Response.Headers["Retry-After"]
                        if ($ra) { [int]::TryParse($ra.ToString(), [ref]$retryAfter) | Out-Null }
                    }
                } catch { }

                # Exponential backoff multiplying by 2 (5s, 10s, 20s, 40s, 80s)
                $backoff = [int](5 * [math]::Pow(2, $attempt - 1))
                $delaySeconds = [math]::Max($retryAfter, $backoff)

                Write-Host "Rate limit (429) hit for '$SubscriptionName'. Retrying in ${delaySeconds}s (attempt $attempt of $MaxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            Write-Warning "Cost query failed for subscription '$SubscriptionName' ($SubscriptionId):"
            Write-Warning $errMsg
            return $null
        }
    }

    return $null
}

function Normalize-MonthKey($rawDate) {
    if ($null -eq $rawDate) { return $null }
    $rawStr = $rawDate.ToString().Trim()

    # Case 1: dd/MM/yyyy or d/M/yyyy (e.g. 01/09/2025 or 01/09/2025 00:00:00)
    if ($rawStr -match "^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})") {
        $firstNum  = [int]$Matches[1]
        $secondNum = [int]$Matches[2]
        $year      = $Matches[3]

        if ($firstNum -eq 1 -and $secondNum -ge 1 -and $secondNum -le 12) {
            $month = $secondNum.ToString().PadLeft(2, "0")
            return "$year-$month-01"
        } elseif ($secondNum -gt 12 -and $firstNum -ge 1 -and $firstNum -le 12) {
            $month = $firstNum.ToString().PadLeft(2, "0")
            return "$year-$month-01"
        } elseif ($firstNum -gt 12 -and $secondNum -ge 1 -and $secondNum -le 12) {
            $month = $secondNum.ToString().PadLeft(2, "0")
            return "$year-$month-01"
        } else {
            $month = $secondNum.ToString().PadLeft(2, "0")
            return "$year-$month-01"
        }
    }

    # Case 2: yyyy-MM-dd or yyyy/MM/dd or yyyy-MM
    if ($rawStr -match "^(\d{4})[/\-](\d{1,2})") {
        $year = $Matches[1]
        $month = [int]$Matches[2]
        $monthStr = $month.ToString().PadLeft(2, "0")
        return "$year-$monthStr-01"
    }

    # Case 3: 8-digit integer like 20251201 or "20251201"
    if ($rawStr -match "^(\d{4})(\d{2})(\d{2})$") {
        return "$($Matches[1])-$($Matches[2])-01"
    }

    # Case 4: 6-digit integer like 202512 or "202512"
    if ($rawStr -match "^(\d{4})(\d{2})$") {
        return "$($Matches[1])-$($Matches[2])-01"
    }

    # Case 5: General DateTime parsing fallback
    $parsedDate = [DateTime]::MinValue
    if ([DateTime]::TryParse($rawStr, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
        return (Get-Date -Year $parsedDate.Year -Month $parsedDate.Month -Day 1).ToString("yyyy-MM-01")
    }

    return $rawStr
}

# ------------------------------------------------------------------
# 5. Loop subscriptions in this tenant
# ------------------------------------------------------------------
$allSubsInTenant = Get-AzSubscription -TenantId $tenant.Id -ErrorAction SilentlyContinue | Where-Object { $_.TenantId -eq $tenant.Id }
if (-not $allSubsInTenant) {
    Write-Error "No subscriptions visible under tenant $($tenant.Name)."
    return
}

if ($Subscriptions -and $Subscriptions.Count -gt 0) {
    $cleanSubPatterns = @()
    foreach ($s in $Subscriptions) {
        if ($s -match ",") {
            $cleanSubPatterns += ($s -split ",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        } else {
            $cleanSubPatterns += $s.Trim()
        }
    }

    $subs = $allSubsInTenant | Where-Object {
        $subObj = $_
        $matched = $false
        foreach ($pattern in $cleanSubPatterns) {
            if ($subObj.Id -eq $pattern -or $subObj.Name -eq $pattern -or $subObj.Name -like "*$pattern*") {
                $matched = $true
                break
            }
        }
        $matched
    }

    if (-not $subs) {
        Write-Error "None of the specified subscriptions ($($cleanSubPatterns -join ', ')) were found in tenant '$($tenant.Name)'."
        Write-Host "`nAvailable subscriptions in this tenant:" -ForegroundColor Cyan
        $allSubsInTenant | Format-Table Name, Id -AutoSize
        return
    }
} else {
    $subs = $allSubsInTenant
}

Write-Host "`nSubscriptions to query ($($subs.Count) of $(($allSubsInTenant).Count)):" -ForegroundColor Cyan
$subs | Format-Table Name, Id -AutoSize

$flatRows = @()
$failedSubscriptions = @()

# Candidate names for the date/period column. The Cost Management query
# API does not always come back with a column literally named
# "UsageDate" for Monthly granularity - it can be "BillingMonth" or
# "Date" depending on API version / grouping.
$dateColumnCandidates = @("UsageDate", "BillingMonth", "Date")

foreach ($sub in $subs) {
    Write-Host "`nSubscription: $($sub.Name) [$($sub.Id)]" -ForegroundColor Yellow

    $subTenantId = if ($sub.TenantId) { $sub.TenantId } else { $tenant.Id }
    try {
        Set-AzContext -SubscriptionId $sub.Id -TenantId $subTenantId -ErrorAction Stop | Out-Null
    } catch {
        try {
            Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "Could not set context for subscription $($sub.Name): $($_.Exception.Message)"
            $failedSubscriptions += [PSCustomObject]@{ SubscriptionName = $sub.Name; SubscriptionId = $sub.Id; Reason = "Could not set context: $($_.Exception.Message)" }
            continue
        }
    }

    $subSuccess = $true
    $subRows = @()

    foreach ($chunk in $timeChunks) {
        if ($timeChunks.Count -gt 1) {
            Write-Host "  Querying chunk: $($chunk.FromStr.Substring(0,10)) to $($chunk.ToStr.Substring(0,10))..." -ForegroundColor DarkGray
        }

        $result = Get-SubscriptionMonthlyCost -SubscriptionId $sub.Id -SubscriptionName $sub.Name -TenantId $subTenantId -FromStr $chunk.FromStr -ToStr $chunk.ToStr
        if (-not $result) {
            Write-Warning "No cost data returned for $($sub.Name) for period chunk $($chunk.FromStr.Substring(0,10)) (API query failed or rate limited)."
            $failedSubscriptions += [PSCustomObject]@{ SubscriptionName = $sub.Name; SubscriptionId = $sub.Id; Reason = "API query failed or rate limited" }
            $subSuccess = $false
            break
        }

        if (-not $result.properties.rows -or $result.properties.rows.Count -eq 0) {
            continue
        }

        $columns = $result.properties.columns.name

        $costIdx         = [array]::IndexOf($columns, "Cost")
        $pricingModelIdx = [array]::IndexOf($columns, "PricingModel")
        $currencyIdx     = [array]::IndexOf($columns, "Currency")

        $dateIdx = -1
        foreach ($candidate in $dateColumnCandidates) {
            $idx = [array]::IndexOf($columns, $candidate)
            if ($idx -ge 0) {
                $dateIdx = $idx
                break
            }
        }

        if ($dateIdx -lt 0) {
            Write-Error "Could not find a date/period column (tried: $($dateColumnCandidates -join ', ')) for subscription '$($sub.Name)'. Skipping this subscription."
            $failedSubscriptions += [PSCustomObject]@{ SubscriptionName = $sub.Name; SubscriptionId = $sub.Id; Reason = "Date column not found" }
            $subSuccess = $false
            break
        }
        if ($costIdx -lt 0) {
            Write-Error "Could not find a 'Cost' column for subscription '$($sub.Name)'. Skipping this subscription."
            $failedSubscriptions += [PSCustomObject]@{ SubscriptionName = $sub.Name; SubscriptionId = $sub.Id; Reason = "Cost column not found" }
            $subSuccess = $false
            break
        }

        foreach ($row in $result.properties.rows) {
            $monthKey = Normalize-MonthKey $row[$dateIdx]

            $subRows += [PSCustomObject]@{
                TenantName       = $tenant.Name
                SubscriptionName = $sub.Name
                SubscriptionId   = $sub.Id
                Month            = $monthKey
                PricingModel     = if ($pricingModelIdx -ge 0) { $row[$pricingModelIdx] } else { "Unknown" }
                Cost             = [math]::Round([double]$row[$costIdx], 2)
                Currency         = if ($currencyIdx -ge 0) { $row[$currencyIdx] } else { "" }
            }
        }

        if ($timeChunks.Count -gt 1) {
            Start-Sleep -Milliseconds 500
        }
    }

    if ($subSuccess -and $subRows.Count -gt 0) {
        $flatRows += $subRows
        Write-Host "Collected $($subRows.Count) cost record(s) for $($sub.Name)." -ForegroundColor DarkGray
    } elseif ($subSuccess -and $subRows.Count -eq 0) {
        Write-Host "No cost usage rows returned for $($sub.Name) (subscription may be empty/inactive)." -ForegroundColor DarkGray
    }

    # Polite pacing delay between subscriptions to prevent ARM burst rate limit triggers
    Start-Sleep -Milliseconds 800
}

function Show-RetrySuggestion {
    param($failedList)

    if ($failedList -and $failedList.Count -gt 0) {
        $namesArray = @()
        foreach ($item in $failedList) {
            if ($item.SubscriptionName) {
                $namesArray += "`"$($item.SubscriptionName)`""
            }
        }
        $failedNames = ($namesArray | Select-Object -Unique) -join ","
        $tenantParam = if ($TenantName) { $TenantName } else { $tenant.Name }

        Write-Host "`n==================================================================" -ForegroundColor Yellow
        Write-Host "WARNING: $($failedList.Count) subscription(s) failed during execution:" -ForegroundColor Yellow
        foreach ($f in $failedList) {
            Write-Host "  - $($f.SubscriptionName) [$($f.SubscriptionId)]: $($f.Reason)" -ForegroundColor Red
        }

        Write-Host "`nDue to error(s), please rerun all failed subscriptions in a single command:" -ForegroundColor Cyan
        $retryCmd = "pwsh AzureCost/azure-cost-report.ps1 -Tenant `"$tenantParam`" -Subscriptions $failedNames"
        if ($StartDate) {
            $retryCmd += " -StartDate `"$StartDate`""
        } elseif ($MonthsBack -ne 12) {
            $retryCmd += " -MonthsBack $MonthsBack"
        }
        if ($EndDate) {
            $retryCmd += " -EndDate `"$EndDate`""
        }
        if ($OutputDir -ne "Output") {
            $retryCmd += " -OutputDir `"$OutputDir`""
        }
        Write-Host "`n  $retryCmd`n" -ForegroundColor Green
        Write-Host "==================================================================" -ForegroundColor Yellow
    }
}

if (-not $flatRows) {
    Show-RetrySuggestion $failedSubscriptions
    Write-Error "No cost data collected for any subscription in this tenant."
    return
}

# ------------------------------------------------------------------
# 6. Export
# ------------------------------------------------------------------
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$safeTenant = ($tenant.Name -replace '[^\w\-]', '_')

$flatFile = Join-Path $OutputDir "cost_${safeTenant}_flat_$timestamp.csv"
$flatRows | Sort-Object Month, TenantName, SubscriptionName, PricingModel | Export-Csv -Path $flatFile -NoTypeInformation
Write-Host "`nExported cost CSV: $flatFile" -ForegroundColor Cyan

$allMonths = $flatRows | Select-Object -ExpandProperty Month -Unique
$subKeys = $flatRows | Select-Object TenantName, SubscriptionName, SubscriptionId -Unique

Write-Host "`nDone. $($subKeys.Count) subscriptions, $($allMonths.Count) months covered." -ForegroundColor Green

Show-RetrySuggestion $failedSubscriptions
