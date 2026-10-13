# 01 · SharePoint Client Portal Governance

Each client of the firm gets its own SharePoint team site ("client portal") with a standard set of lists (e.g. *IMMIGRATION DASHBOARD*, *REGULATORY COMPLIANCE*) and a *PORTAL* document library shared with the client's guest users. These scripts keep dozens of portals consistent from one place.

All mass scripts discover portals dynamically with `Get-PnPTenantSite` and filter out internal / system sites, so new portals are picked up automatically.

| Script | What it does | Auth |
|---|---|---|
| `Get-ClientPortalSites.ps1` | Lists all client portal sites (reusable discovery/filter block) | Interactive |
| `Get-PortalAdminsAndOwners.ps1` | Exports site collection admins + Owners group members of every portal to CSV | Certificate |
| `Get-PortalNavigationReport.ps1` | Read-only inventory of QuickLaunch & TopNav nodes (incl. children) across portals | Certificate |
| `Repair-PortalNavigationLink.ps1` | Idempotently fixes/creates the *PORTAL* QuickLaunch link → `Updated / Already Correct / Created` report | Certificate |
| `Set-PortalGuestReadOnly.ps1` | Breaks inheritance on sensitive lists and gives each guest (`#ext#`) member **Read** only | Certificate |
| `Sync-PortalListSchema-IMMD.ps1` | Rolls out a list schema from a template site: removes extra columns, adds missing ones (normal before calculated), syncs calculated formula, JSON formatting and views | Certificate |
| `Sync-PortalListFormatting-RC.ps1` | Copies column/view JSON formatting of *REGULATORY COMPLIANCE* to all portals and reorders view fields | Certificate |
| `Test-PortalListExists.ps1` | Audits whether a list exists on every portal → CSV | Interactive |
| `Copy-ListToAllPortals.ps1` | Clones a list from the template portal to every portal (`Copy-PnPList`) | Interactive |
| `Remove-ListFromAllPortals.ps1` | Rollback for the clone above ⚠️ destructive, no prompt | Interactive |
| `Copy-PortalPage.ps1` | Copies a modern page with all web parts between sites (`Export-PnPPage` → `Invoke-PnPSiteTemplate`) | Interactive |
| `Start-PortalOneDriveSync.ps1` | Finds portals via Search and launches `odopen://sync` so each *PORTAL* library syncs to the user's OneDrive | Interactive |

## Techniques

- Tenant-wide discovery + exclusion filters
- Per-site `try/catch` so one broken site doesn't stop the run; results to timestamped CSV
- CSOM: `BreakRoleInheritance`, `RoleAssignments`, `AddFieldAsXml`, `CustomFormatter`
- `-ReturnConnection` to work with several sites in one session

## Requirements

`PnP.PowerShell` 2.x, PowerShell 7, an Entra app (see [`config`](../config)).
