# 03 · Tenant Reporting & Audit

| Script | What it does | Module / Auth |
|---|---|---|
| `Get-TenantLargeFileSummary.ps1` | Uses **SharePoint Search (KQL)** to find files ≥ 3 MB on selected sites; counts by site, extension and size band (3–5, 5–8, > 8 MB) → CSV. Paged 500 rows at a time — far faster than crawling libraries | PnP, certificate |
| `Get-TenantLargeFileSummary-LastMonth.ps1` | Same, limited to files modified in the last month (`LastModifiedTime>=`) | PnP, certificate |
| `Get-MailboxTimeZoneAudit.ps1` | Audits every enabled user's mailbox regional settings (time zone + additional time zones) via Graph → CSV. Used to troubleshoot meeting-time issues | Microsoft.Graph, certificate |
| `Export-SharedMailboxAuditLog.ps1` | Pulls the Unified Audit Log for a shared mailbox in a time window, parses `AuditData` JSON, flags restore/recover operations, converts UTC → GMT+8 → CSV | ExchangeOnlineManagement, interactive |
| `Compare-RoomCalendarConfig.ps1` | Compares working hours/days/time zone of a user vs a room mailbox to explain "room unavailable" booking errors | ExchangeOnlineManagement, interactive |

Edit the date window at the top of `Export-SharedMailboxAuditLog.ps1` before running. Mailbox and room names come from `$Config`.
