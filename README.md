# Azure Management & Reporting Toolset

A collection of PowerShell automation tools for querying, analyzing, and reporting on Azure infrastructure across multiple tenants and subscriptions.

## Features

- **Azure Virtual Desktop (AVD) Statistics (`avd-stats.ps1`)**: Discovers host pools across subscriptions, auto-resolves linked Log Analytics workspaces, and aggregates active sessions, active users, closed sessions (24h), and monthly unique users (30d).
- **Azure Cost Reporting (`azure-cost-report.ps1`)**: Queries Azure Cost Management REST APIs per tenant and subscription, categorizing monthly spend by pricing model (`Reservation`, `OnDemand`, `SavingsPlan`).
- **Multi-Tenant Cost Aggregation (`Merge-CostReports.ps1`)**: Merges multiple single-tenant cost reports into a consolidated master report with dynamic chronological pivoting and tenant subtotals.

---

## Directory Structure

```text
.
├── AvdStats/
│   └── avd-stats.ps1             # AVD active sessions & user statistics
├── AzureCost/
│   ├── azure-cost-report.ps1     # Monthly subscription cost query by pricing model
│   └── Merge-CostReports.ps1     # Multi-tenant cost consolidation and merge tool
├── Output/                       # Default destination for generated CSV reports
└── README.md
```

---

## Prerequisites & Installation

### PowerShell Requirements

The scripts run on **PowerShell 7+ (`pwsh`)** or **Windows PowerShell 5.1**. Required Azure PowerShell modules are automatically detected and installed in the user scope if missing:

- `Az.Accounts`
- `Az.Resources`
- `Az.Monitor`
- `Az.OperationalInsights`

To manually install the Azure PowerShell module suite:

```powershell
pwsh -Command "Install-Module -Name Az -Scope CurrentUser -Force -AllowClobber"
```

### Azure RBAC Permissions (Read-Only)

The scripts only perform read queries. Ensure your Azure identity has the following roles:

- **AVD Stats**: `Reader` or `Desktop Virtualization Reader` on subscriptions/host pools, and `Monitoring Reader` or `Log Analytics Reader` on target Log Analytics workspaces.
- **Cost Reports**: `Cost Management Reader` (or `Reader` / `Billing Reader`) on target subscriptions.

> [!NOTE]
> **READ-ONLY GUARANTEE**: All scripts strictly query data (`Get-*`, `Connect-*`, `Set-AzContext`, `Invoke-AzOperationalInsightsQuery`, and `Invoke-RestMethod` to `Microsoft.CostManagement/query`). No cloud resources, budgets, subscriptions, or settings are ever created, modified, or deleted.

---

## Usage Guide

### 1. Azure Virtual Desktop (AVD) Stats

Collects active sessions and user counts across all host pools in a tenant:

```bash
# Interactive tenant selection (writes to Output/ by default)
pwsh AvdStats/avd-stats.ps1

# Target specific tenant by Name
pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft"

# Target specific tenant by Tenant ID GUID and filter to specific subscription(s)
pwsh AvdStats/avd-stats.ps1 -TenantId "842ae988-60f7-4e75-9477-32c43f90daf5" -Subscription "ASWeb"

# Target specific tenant with multiple subscriptions
pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft.cloud" -Subscriptions "CIBC","Wilshire" -OutputDir "Output"
```

#### Output Schema (`avd_active_sessions_<Tenant>_<Timestamp>.csv`)

| Column           | Description                                                            |
| :--------------- | :--------------------------------------------------------------------- |
| `HostPoolName`   | Name of the Azure Virtual Desktop host pool                            |
| `ActiveSessions` | Number of currently open sessions (`State == 'Connected'` in last 24h) |
| `ActiveUsers`    | Distinct users with active sessions                                    |
| `ClosedSessions` | Sessions that completed in the last 24 hours                           |
| `UsersLastMonth` | Unique active users with any connection in the last 30 days            |

---

### 2. Azure Cost Reports

Generates monthly actual cost breakdowns by pricing model (`Reservation`, `OnDemand`, `SavingsPlan`) per subscription for a given tenant:

```bash
# Interactive tenant selection (defaults to last 12 months, outputs to Output/)
pwsh AzureCost/azure-cost-report.ps1

# Target specific tenant by Name
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft" -MonthsBack 24

# Target specific tenant by Tenant ID GUID (useful when multiple tenants share the same name)
pwsh AzureCost/azure-cost-report.ps1 -TenantId "842ae988-60f7-4e75-9477-32c43f90daf5"

# Target specific subscription by name or ID (or wildcard pattern)
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft" -Subscription "ASWeb"

# Target multiple subscriptions for custom historical period (e.g. 24 months, 36 months)
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft" -MonthsBack 24

# Target specific date range
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft" -StartDate "2024-01-01" -EndDate "2026-08-01"
```

> [!TIP]
> **Historical Queries (> 12 Months)**: Azure Cost Management API limits single queries to a maximum of 1 year (12 months). When requesting historical periods greater than 12 months (e.g., `-MonthsBack 24` or `-MonthsBack 36`), the script automatically partitions the query into 12-month chunks and seamlessly merges the results.

#### Generated Output (`cost_<Tenant>_flat_<Timestamp>.csv`):

- Columns: `TenantName`, `SubscriptionName`, `SubscriptionId`, `Month` (formatted as `YYYY-MM-01`, representing the first day of the month), `PricingModel`, `Cost` (rounded to 2 decimal places), `Currency`

---

### 3. Merging Multi-Tenant Cost Reports

Combines individual tenant cost CSV files into a unified master report:

```bash
# Scans Output/ directory, auto-detects latest tenant runs, and creates master file
pwsh AzureCost/Merge-CostReports.ps1

# Explicit input/output directories
pwsh AzureCost/Merge-CostReports.ps1 -InputDir "Output" -OutputDir "Output"

# Explicit list of specific CSV files
pwsh AzureCost/Merge-CostReports.ps1 -Files "Output/cost_AlternativeSoft_flat_20260805_1746.csv","Output/cost_AlternativeSoft_cloud_flat_20260805_1746.csv"
```

#### Master Output:

- **`cost_MASTER_flat_<Timestamp>.csv`**: Consolidated cost report across all tenants and subscriptions with an on-screen summary table.

---

## Typical Multi-Tenant Workflow

```bash
# Step 1: Run cost report for Tenant 1
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft" -OutputDir "Output"

# Step 2: Run cost report for Tenant 2
pwsh AzureCost/azure-cost-report.ps1 -Tenant "AlternativeSoft.cloud" -OutputDir "Output"

# Step 3: Merge into consolidated master report
pwsh AzureCost/Merge-CostReports.ps1 -InputDir "Output" -OutputDir "Output"

# Step 4: Run AVD usage stats for both tenants
pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft" -OutputDir "Output"
pwsh AvdStats/avd-stats.ps1 -Tenant "AlternativeSoft.cloud" -OutputDir "Output"
```
