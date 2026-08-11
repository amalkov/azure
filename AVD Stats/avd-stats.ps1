<#
.SYNOPSIS
    Reports active AVD sessions/users per host pool, across all subscriptions,
    by auto-discovering the Log Analytics workspace from host pool diagnostic settings.

.NOTES
    Requires modules: Az.Accounts, Az.Resources, Az.Monitor, Az.OperationalInsights
    Install with:
        Install-Module -Name Az -Scope CurrentUser -Force
#>

# ------------------------------------------------------------------
# 0. Ensure required modules are available
# ------------------------------------------------------------------
$requiredModules = @('Az.Accounts', 'Az.Resources', 'Az.Monitor', 'Az.OperationalInsights')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing missing module: $m" -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
}

# ------------------------------------------------------------------
# 1. Connect
# ------------------------------------------------------------------
if (-not (Get-AzContext)) {
    Connect-AzAccount | Out-Null
}

Write-Host "Scanning all subscriptions for AVD host pools..." -ForegroundColor Cyan
$subs = Get-AzSubscription

$hostPoolInventory = @()
$workspaceResourceId = $null

foreach ($sub in $subs) {
    try {
        Set-AzContext -Subscription $sub.Id -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not set context for subscription $($sub.Name): $_"
        continue
    }

    $hostPools = Get-AzResource -ResourceType "Microsoft.DesktopVirtualization/hostpools" -ErrorAction SilentlyContinue

    foreach ($hp in $hostPools) {
        $diag = Get-AzDiagnosticSetting -ResourceId $hp.ResourceId -ErrorAction SilentlyContinue

        $hostPoolInventory += [PSCustomObject]@{
            Subscription   = $sub.Name
            SubscriptionId = $sub.Id
            HostPoolName   = $hp.Name
            ResourceId     = $hp.ResourceId
            DiagEnabled    = ($diag.Count -gt 0)
            WorkspaceId    = if ($diag) { ($diag | Select-Object -First 1).WorkspaceId } else { $null }
        }

        # Capture the first workspace we find (assumes one central workspace)
        if (-not $workspaceResourceId -and $diag -and $diag[0].WorkspaceId) {
            $workspaceResourceId = $diag[0].WorkspaceId
        }
    }
}

Write-Host "`nHost pool inventory:" -ForegroundColor Cyan
$hostPoolInventory | Format-Table Subscription, HostPoolName, DiagEnabled -AutoSize

if (-not $workspaceResourceId) {
    Write-Error "No Log Analytics workspace found via diagnostic settings. Check that diagnostics are enabled on at least one host pool."
    return
}

# ------------------------------------------------------------------
# 2. Resolve workspace details from the discovered resource ID
# ------------------------------------------------------------------
# WorkspaceId looks like:
# /subscriptions/<subId>/resourceGroups/<rg>/providers/Microsoft.OperationalInsights/workspaces/<name>
$parts = $workspaceResourceId -split '/'
$wsSubId = $parts[2]
$wsRg    = $parts[4]
$wsName  = $parts[-1]

Write-Host "`nUsing Log Analytics workspace: $wsName (RG: $wsRg, Sub: $wsSubId)" -ForegroundColor Cyan

Set-AzContext -Subscription $wsSubId | Out-Null
$workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $wsRg -Name $wsName

# ------------------------------------------------------------------
# 3. Query active sessions per host pool
# ------------------------------------------------------------------
$activeSessionsQuery = @"
WVDConnections
| where TimeGenerated > ago(1h)
| summarize LastState = arg_max(TimeGenerated, State) by HostPoolName, UserName, SessionHostName
| where LastState == 'Connected'
| summarize ActiveSessions = count(), ActiveUsers = dcount(UserName) by HostPoolName
| sort by HostPoolName asc
"@

Write-Host "`nQuerying active sessions..." -ForegroundColor Cyan
$activeResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspace.CustomerId -Query $activeSessionsQuery

Write-Host "`n=== ACTIVE SESSIONS (last 1h) ===" -ForegroundColor Green
$activeResults.Results | Format-Table -AutoSize

# ------------------------------------------------------------------
# 4. Coverage check — which host pools have reported in last 24h
# ------------------------------------------------------------------
$coverageQuery = @"
WVDConnections
| where TimeGenerated > ago(24h)
| summarize LastSeen = max(TimeGenerated) by HostPoolName
| sort by HostPoolName asc
"@

$coverageResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspace.CustomerId -Query $coverageQuery

Write-Host "`n=== COVERAGE (host pools reporting in last 24h) ===" -ForegroundColor Green
$coverageResults.Results | Format-Table -AutoSize

# Flag host pools with diagnostics enabled but no data in the workspace
$reportingNames = $coverageResults.Results.HostPoolName
$missing = $hostPoolInventory | Where-Object { $_.DiagEnabled -and ($_.HostPoolName -notin $reportingNames) }
if ($missing) {
    Write-Host "`n=== WARNING: host pools with diagnostics enabled but no recent data ===" -ForegroundColor Yellow
    $missing | Format-Table Subscription, HostPoolName -AutoSize
}

# ------------------------------------------------------------------
# 5. Export to CSV
# ------------------------------------------------------------------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$outFile = "./avd_active_sessions_$timestamp.csv"
$activeResults.Results | Export-Csv -Path $outFile -NoTypeInformation
Write-Host "`nExported active session results to: $outFile" -ForegroundColor Cyan
