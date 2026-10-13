# 02 · SharePoint List Automation

The firm tracks client work in SharePoint lists on an internal hub site: a client database, a service catalogue, workflow templates (*DATABASE WORKFLOWS*), active jobs (*MONITOR KLIEN*) and their step-by-step progress (*MONITOR TAKA*, up to 70 steps per service). These scripts build and maintain those lists.

## `provisioning/`

| Script | What it does |
|---|---|
| `New-MasterLists.ps1` | Creates master lists (services, steps, service-step map, log) with fields + QuickLaunch header |
| `New-WorkflowDatabaseList.ps1` | Creates the workflow list with a required lookup + 20 step lookups via `Add-PnPFieldFromXml` |
| `Add-TakaLookupColumns.ps1` | Adds 20 lookup columns to an existing list in one CSOM batch |
| `Get-ListInternalNames.ps1` | Exports display name ↔ internal name ↔ type of every column (handy before writing updates) |

## `data-fix/`

Bulk corrections with safety rails.

| Script | What it does | Safety |
|---|---|---|
| `Update-ClientNameFormat.ps1` | Normalises client names (`PT.` → `PT `) | CSV backup for rollback, preview, typed `YA` confirmation, `SystemUpdate` option |
| `Update-MonitorKlienTED.ps1` | Derives target-days (TED) per service from the workflow list and writes it to active jobs | `$dryRun`, `$testID`, conflict detection with `HashSet`, `SystemUpdate` |
| `Add-MissingKbliCompliance.ps1` | Adds the missing compliance row for every client that lacks it, with next zero-padded ID | `$dryRun`, duplicate-ID guard |
| `Set-LkpmDeadline.ps1` | Moves deadline of all *LKPM REPORT* items to the 15th of the month | one-off |
| `Set-TaxManagerFromAccountManager.ps1` | Fills a Person column based on rules (`$Config.TaxManagerRules`), with text-column fallback | rule-based, skip unmatched |

## `monitor-taka/`

| Script | What it does |
|---|---|
| `Test-TakaSequence.ps1` | **Read-only validator**: rebuilds the expected step sequence for a service and compares it with a client's actual rows — count, duplicate/invalid numbers, order, missing & extra steps → `SYNCHRONIZED / NOT SYNCHRONIZED` |
| `Get-TakaIdCompleteness.ps1` | Data-quality check of the text ID vs lookup ID columns → CSV |
| `Remove-TakaOrphanItems.ps1` | Deletes rows with neither ID after a Y/N prompt (to Recycle Bin) |

## `reporting/`

| Script | What it does |
|---|---|
| `Export-MonitorKlienComments.ps1` | Exports active jobs with **all item comments** (REST `/Comments` with paging, @mention resolution, UTC→WITA) to CSV; retries if the file is locked |

## Requirements

`PnP.PowerShell`, certificate app-only auth (see [`config`](../config)). Some provisioning scripts use interactive login.
