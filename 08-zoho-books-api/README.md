# 08 · Zoho Books API

Clients for the Zoho Books v3 REST API, used to pull accounting data (bills, expenses, reporting tags, TDS payables) for reporting and payroll work across many client organisations.

## `python/`

| Script | What it does |
|---|---|
| `auth_helper.py` | Loads credentials from `Auth/.env` and exchanges the **refresh token** for an access token (`get_new_access_token()`) |
| `get_tagged_bills_multi_org.py` | Loops over many organisation IDs, pages through all bills of the year (`page_context.has_more_page`), keeps only bills that have reporting tags, writes one JSON per organisation |

```bash
pip install requests python-dotenv
cp Auth/.env.example Auth/.env     # fill client id/secret, refresh token, org IDs
python get_tagged_bills_multi_org.py
```

## `powershell/`

| Script | What it does |
|---|---|
| `Get-ZohoReportingTags.ps1` | Refreshes the OAuth token and lists reporting tags |
| `Export-ZohoExpensesAndBills.ps1` | Dumps all expenses and bills **with per-item detail calls** (throttled), converting 64-bit IDs to strings before `ConvertFrom-Json` to avoid precision loss |
| `Get-ZohoTdsPayables.ps1` | Exports the TDS payables report to CSV |

Credentials come from `$Config.Zoho` (see [`config`](../config)).

## Getting a refresh token

1. Zoho API Console → *Self Client* → generate a grant code with scope `ZohoBooks.fullaccess.all`.
2. `POST https://accounts.zoho.com/oauth/v2/token` with `grant_type=authorization_code` to receive the refresh token.
3. Store it only in `Auth/.env` / `config.local.ps1`.
