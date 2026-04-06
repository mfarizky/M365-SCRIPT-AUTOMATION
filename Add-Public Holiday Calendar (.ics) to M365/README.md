# Import-ICSToM365

Import holidays from an `.ics` file into a Microsoft 365 calendar via the Graph API.

---

## What it does

1. Connects to Microsoft Graph using app credentials (client secret)
2. Downloads an `.ics` file from a URL
3. Parses all `VEVENT` blocks
4. Creates all-day events tagged `[HOLIDAY]` in the target mailbox calendar

---

## Requirements

| Requirement | Version |
|---|---|
| PowerShell | 7.0+ |
| Microsoft.Graph.Authentication | 2.0.0+ |
| Microsoft.Graph.Calendar | 2.0.0+ |

> **Module auto-install:** the script checks and installs missing or outdated modules automatically on first run (`Scope: CurrentUser`). No manual installation needed.

---

## Setup

### 1. Azure App Registration

In [Azure Portal](https://portal.azure.com) → App registrations → New registration:

- **API Permissions** (Application, not Delegated):
  - `Calendars.ReadWrite`
- Grant admin consent
- Create a **Client Secret** and copy the value immediately

### 2. Configure secrets

```powershell
Copy-Item .env.example .env
# then edit .env with your values
```

Never commit `.env` to git — it's already in `.gitignore`.

### 3. Run

```powershell
# Default (reads .env in current folder)
.\Import-ICSToM365.ps1

# Custom .env location
.\Import-ICSToM365.ps1 -EnvFile "C:\secrets\prod.env"
```

---

## Folder structure

```
ics-to-m365/
├── Import-ICSToM365.ps1   ← main script
├── dependencies.psd1      ← pinned module versions
├── .env.example           ← template — copy to .env
├── .env                   ← your secrets (gitignored)
├── .gitignore
├── README.md
└── scratch/               ← local experiments, gitignored
```

---

## Environment variables

| Variable | Description | Example |
|---|---|---|
| `TENANT_ID` | Azure tenant ID | `xxxxxxxx-xxxx-...` |
| `CLIENT_ID` | App registration client ID | `xxxxxxxx-xxxx-...` |
| `CLIENT_SECRET` | App client secret value | `abc123~...` |
| `TARGET_MAILBOX` | UPN of target mailbox | `user@domain.com` |
| `ICS_URL` | Public URL to `.ics` file | `https://...` |
| `TIME_ZONE` | Windows timezone string | `SE Asia Standard Time` |

---

## Output example

```
==============================
 DEPENDENCY CHECK
==============================
[   OK    ] Microsoft.Graph.Authentication v2.19.0
[   OK    ] Microsoft.Graph.Calendar v2.19.0

==============================
 IMPORTING TO M365 CALENDAR
==============================
  ADDED : New Year's Day (2025-01-01)
  ADDED : Eid al-Fitr (2025-03-31)
  SKIP  : (incomplete event — missing SUMMARY)

==============================
 IMPORT FINISHED
==============================
Added          : 17
Skipped        : 1
Errors         : 0
```

---

## Troubleshooting

**`Insufficient privileges`** — make sure `Calendars.ReadWrite` is Application permission (not Delegated) and admin consent has been granted.

**`No events parsed`** — the ICS URL might require authentication, or the format uses non-standard line endings. Check the raw content by running `Invoke-WebRequest -Uri $url` manually.

**`ParseExact` date error** — some ICS files include time components in DTSTART (e.g. `20250101T000000Z`). The script trims to 8 characters (`yyyyMMdd`) so this is handled, but log the raw value if issues persist.

---

## Notes

- Events are created as **all-day**, **Out of Office (`oof`)**, prefixed with `[HOLIDAY]`
- The script does **not** deduplicate — running twice will create duplicate events
- Tested with ICS files from: Google Calendar, Outlook, TimeAndDate.com
