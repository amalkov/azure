<#
.SYNOPSIS
    Reports active AVD sessions, active users, closed sessions, and monthly unique users
    per host pool across all subscriptions for a single tenant.

.DESCRIPTION
    Scans all subscriptions in the selected Azure AD / Entra ID tenant to discover
    Azure Virtual Desktop (AVD) host pools and their diagnostic Log Analytics workspaces.
    Queries each Log Analytics workspace for:
      - Active sessions (State == 'Connected' within last 24h)
      - Active unique users (State == 'Connected' within last 24h)
      - Closed sessions (State == 'Completed' within last 24h)
      - Monthly unique users (any connection event within last 30 days)

.PARAMETER TenantName
    Name, domain fragment, or Tenant ID GUID to match against Get-AzTenant (e.g.
    "AlternativeSoft", "AlternativeSoft.cloud", or "842ae988-60f7-4e75-9477-32c43f90daf5").
    If omitted, the script presents an interactive menu to select a tenant.

.PARAMETER Subscriptions
    One or more subscription names, IDs, or wildcard patterns to limit the scan to.
    If omitted, all subscriptions in the selected tenant are scanned.

.PARAMETER OutputDir
    Directory path where CSV results are exported. Defaults to "Output".

.NOTES
    Requires modules: Az.Accounts, Az.Resources, Az.Monitor, Az.OperationalInsights
    Prerequisites: Reader or Monitoring Reader role on target subscriptions and workspaces.

    READ-ONLY GUARANTEE:
    This script strictly performs read-only queries (Get-*, Connect-*, Set-AzContext,
    and Invoke-AzOperationalInsightsQuery). No Azure resource or configuration is modified.

.EXAMPLE
    pwsh AvdStats/avd-stats.ps1
    pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft"
    pwsh AvdStats/avd-stats.ps1 -TenantId "842ae988-60f7-4e75-9477-32c43f90daf5" -Subscription "ASWeb"
    pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft.cloud" -Subscriptions "CIBC","Wilshire" -OutputDir "Output"
#>

param(
    [Alias("Tenant", "TenantId")]
    [string]$TenantName,

    [Alias("Subscription", "SubscriptionName", "SubscriptionId", "SubscriptionNames", "SubscriptionIds")]
    [string[]]$Subscriptions,

    [string]$OutputDir = "Output"
)

# ------------------------------------------------------------------
# 1. Ensure required modules are available
# ------------------------------------------------------------------
$requiredModules = @('Az.Accounts', 'Az.Resources', 'Az.Monitor', 'Az.OperationalInsights')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing missing module: $m" -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
}

# ------------------------------------------------------------------
# 2. Connect & resolve the tenant
# ------------------------------------------------------------------
if (-not (Get-AzContext)) {
    Connect-AzAccount | Out-Null
}

if ($TenantName) {
    $tenant = Get-AzTenant | Where-Object {
        $_.Id -eq $TenantName -or
        $_.TenantId -eq $TenantName -or
        $_.Name -like "*$TenantName*" -or
        $_.DefaultDomain -like "*$TenantName*"
    } | Select-Object -First 1

    if (-not $tenant) {
        Write-Error "Tenant matching '$TenantName' not found. Run 'Get-AzTenant' to see exact names/domains/IDs."
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

Write-Host "`nUsing tenant: $($tenant.Name) [$($tenant.Id)]" -ForegroundColor Cyan

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
# 3. Discover subscriptions & host pools in this tenant
# ------------------------------------------------------------------
$allSubsInTenant = Get-AzSubscription -TenantId $tenant.Id -ErrorAction SilentlyContinue | Where-Object { $_.TenantId -eq $tenant.Id }

if (-not $allSubsInTenant) {
    Write-Error "No subscriptions found under tenant '$($tenant.Name)'."
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

Write-Host "`nSubscriptions to scan ($($subs.Count) of $(($allSubsInTenant).Count)):" -ForegroundColor Cyan
$subs | Format-Table Name, Id -AutoSize

$hostPoolInventory = @()
$failedSubscriptions = @()

foreach ($sub in $subs) {
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

    $hostPools = Get-AzResource -ResourceType "Microsoft.DesktopVirtualization/hostpools" -ErrorAction SilentlyContinue

    foreach ($hp in $hostPools) {
        $diag = Get-AzDiagnosticSetting -ResourceId $hp.ResourceId -ErrorAction SilentlyContinue
        $wsId = if ($diag -and $diag.Count -gt 0) { $diag[0].WorkspaceId } else { $null }

        $hostPoolInventory += [PSCustomObject]@{
            Subscription   = $sub.Name
            SubscriptionId = $sub.Id
            HostPoolName   = $hp.Name
            ResourceId     = $hp.ResourceId
            DiagEnabled    = ($diag.Count -gt 0)
            WorkspaceId    = $wsId
        }
    }
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
        $retryCmd = "pwsh AvdStats/avd-stats.ps1 -Tenant `"$tenantParam`" -Subscriptions $failedNames"
        if ($OutputDir -ne "Output") {
            $retryCmd += " -OutputDir `"$OutputDir`""
        }
        Write-Host "`n  $retryCmd`n" -ForegroundColor Green
        Write-Host "==================================================================" -ForegroundColor Yellow
    }
}

if (-not $hostPoolInventory) {
    Write-Host "`nNo AVD host pools found in tenant $($tenant.Name)." -ForegroundColor Yellow
    Show-RetrySuggestion $failedSubscriptions
    return
}

Write-Host "`nHost pools discovered across tenant:" -ForegroundColor Cyan
$hostPoolInventory | Format-Table Subscription, HostPoolName, DiagEnabled -AutoSize

$withWorkspace = $hostPoolInventory | Where-Object { $_.WorkspaceId }
if (-not $withWorkspace) {
    Write-Error "No host pools in this tenant have a Log Analytics workspace configured in diagnostic settings."
    Show-RetrySuggestion $failedSubscriptions
    return
}

# ------------------------------------------------------------------
# 4. Query each distinct Log Analytics workspace
# ------------------------------------------------------------------
$workspaceGroups = $withWorkspace | Group-Object WorkspaceId
$allActiveResults = @()
$allCoverageResults = @()

foreach ($group in $workspaceGroups) {
    $wsResourceId = $group.Name
    $parts   = $wsResourceId -split '/'
    $wsSubId = $parts[2]
    $wsRg    = $parts[4]
    $wsName  = $parts[-1]

    Write-Host "`n=== Querying Log Analytics Workspace: $wsName (RG: $wsRg) ===" -ForegroundColor Yellow

    try {
        Set-AzContext -Subscription $wsSubId -Tenant $tenant.Id -ErrorAction Stop | Out-Null
        $workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $wsRg -Name $wsName -ErrorAction Stop
    } catch {
        Write-Warning "Could not resolve workspace '$wsName': $($_.Exception.Message)"
        continue
    }

    # Query details:
    # - HostPoolName is parsed from _ResourceId (.../hostpools/<name>)
    # - arg_max(TimeGenerated, State) keeps the latest connection State
    # - Active: State == 'Connected' within last 24h
    # - Closed: State == 'Completed' within last 24h
    # - MonthlyUsers: distinct users over the past 30 days
    $activeSessionsQuery = @"
let sessionStatus =
    WVDConnections
    | where TimeGenerated > ago(24h)
    | extend HostPoolName = tostring(split(_ResourceId, '/')[-1])
    | summarize arg_max(TimeGenerated, State) by HostPoolName, UserName, SessionHostName
    | summarize ActiveSessions = countif(State == 'Connected'),
                ActiveUsers = dcountif(UserName, State == 'Connected'),
                ClosedSessions = countif(State == 'Completed')
      by HostPoolName;
let monthlyUsers =
    WVDConnections
    | where TimeGenerated > ago(30d)
    | extend HostPoolName = tostring(split(_ResourceId, '/')[-1])
    | summarize UsersLastMonth = dcount(UserName) by HostPoolName;
sessionStatus
| join kind=fullouter monthlyUsers on HostPoolName
| project HostPoolName = iff(isempty(HostPoolName), HostPoolName1, HostPoolName),
          ActiveSessions = coalesce(ActiveSessions, 0),
          ActiveUsers = coalesce(ActiveUsers, 0),
          ClosedSessions = coalesce(ClosedSessions, 0),
          UsersLastMonth = coalesce(UsersLastMonth, 0)
| sort by HostPoolName asc
"@

    try {
        $activeResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspace.CustomerId -Query $activeSessionsQuery -ErrorAction Stop
        if ($activeResults.Results) {
            $activeResults.Results | Format-Table -AutoSize
            $allActiveResults += $activeResults.Results
        }
    } catch {
        Write-Warning "Active sessions query failed on workspace '$wsName': $($_.Exception.Message)"
        if ($_.ErrorDetails.Message) { Write-Warning $_.ErrorDetails.Message }
        continue
    }

    $coverageQuery = @"
WVDConnections
| where TimeGenerated > ago(24h)
| extend HostPoolName = tostring(split(_ResourceId, '/')[-1])
| summarize LastSeen = max(TimeGenerated) by HostPoolName
| sort by HostPoolName asc
"@
    try {
        $coverageResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspace.CustomerId -Query $coverageQuery -ErrorAction Stop
        if ($coverageResults.Results) {
            $allCoverageResults += $coverageResults.Results
        }
    } catch {
        Write-Warning "Coverage query failed on workspace '$wsName': $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------
# 5. Diagnostic coverage check
# ------------------------------------------------------------------
$reportingNames = $allCoverageResults.HostPoolName
$missing = $hostPoolInventory | Where-Object { $_.DiagEnabled -and ($_.HostPoolName -notin $reportingNames) }
if ($missing) {
    Write-Host "`n=== WARNING: Host pools with diagnostics enabled but no telemetry in last 24h ===" -ForegroundColor Yellow
    $missing | Format-Table Subscription, HostPoolName -AutoSize
}

# ------------------------------------------------------------------
# 6. Export to CSV
# ------------------------------------------------------------------
if ($allActiveResults.Count -gt 0) {
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $safeTenantName = ($tenant.Name -replace '[^\w\-]', '_')
    $outFile = Join-Path $OutputDir "avd_active_sessions_${safeTenantName}_$timestamp.csv"
    $allActiveResults | Export-Csv -Path $outFile -NoTypeInformation
    Write-Host "`nExported active session results to: $outFile" -ForegroundColor Green
} else {
    Write-Host "`nNo active session records found to export." -ForegroundColor Yellow
}

Show-RetrySuggestion $failedSubscriptions
