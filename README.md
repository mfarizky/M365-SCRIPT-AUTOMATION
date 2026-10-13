# Microsoft 365 Automation Portfolio

A collection of production scripts I wrote while administering a Microsoft 365 tenant for a consulting firm (~50+ client portals, SharePoint-based operational lists, OneDrive storage clean-up, and finance integrations with Zoho Books).

Every script here was used against a real tenant. Credentials, tenant identifiers, client names and personal paths have been moved out of the code into a local config file that is **not** part of this repository.

## Contents

| Folder | What it shows |
|---|---|
| [`01-sharepoint-client-portal`](01-sharepoint-client-portal) | Tenant-wide governance of client portal sites: inventory, permission hardening for guests, navigation repair, list schema & formatting roll-out |
| [`02-sharepoint-list-automation`](02-sharepoint-list-automation) | Provisioning master lists, bulk data fixes with dry-run / backup / rollback, workflow sequence validation, comment export |
| [`03-tenant-reporting-audit`](03-tenant-reporting-audit) | Storage reports via SharePoint Search (KQL), mailbox time-zone audit via Graph, shared mailbox audit log, room calendar diagnostics |
| [`04-onedrive-file-compression`](04-onedrive-file-compression) | Scan → compress → track pipeline for large images/PDFs (ImageMagick, Ghostscript, Adobe PDF Services API) that preserves OneDrive version history, plus an incident-recovery toolkit |
| [`05-list-version-history-analytics`](05-list-version-history-analytics) | Backfill + incremental export of SharePoint list version history for analytics (Dataverse / Power BI) |
| [`06-power-apps`](06-power-apps) | Power Fx OnSave logic for a parent/child form with concurrency check and bulk Patch |
| [`07-python-kbli-extractor`](07-python-kbli-extractor) | Python OCR pipeline that extracts business classification codes (KBLI) from Indonesian deeds of establishment |
| [`08-zoho-books-api`](08-zoho-books-api) | Zoho Books REST API clients (OAuth refresh flow, pagination, multi-organization export) in Python and PowerShell |
| [`09-intune-endpoint`](09-intune-endpoint) | Intune remediation script |
| [`setup`](setup) | One-time workstation setup |

## Tech stack

PowerShell 7 · PnP.PowerShell · Microsoft Graph PowerShell · Exchange Online Management · SharePoint CSOM/REST · Power Fx · Python (PyMuPDF, pytesseract, OpenCV, requests) · Ghostscript · ImageMagick · Adobe PDF Services API · Zoho Books API

## Getting started

```powershell
# 1. Install prerequisites
.\setup\Install-Prerequisites.ps1

# 2. Create your local config (git-ignored)
Copy-Item .\config\config.example.ps1 .\config\config.local.ps1
notepad .\config\config.local.ps1

# 3. (Recommended) keep the certificate password out of files
$env:PNP_PFX_PASSWORD = "<pfx password>"

# 4. Run any script
.\03-tenant-reporting-audit\Get-TenantLargeFileSummary.ps1
```

Every PowerShell script dot-sources `config/Load-Config.ps1`, which loads `config/config.local.ps1` into a `$Config` hashtable. See [`config/README.md`](config/README.md).

Python folders use a `.env` file instead (each has a `.env.example`).

## Authentication patterns used

- **App-only (certificate)** – Entra ID app registration + PFX certificate, used for unattended/bulk jobs (`Connect-PnPOnline -ClientId -Tenant -CertificatePath`).
- **Delegated (interactive)** – for ad-hoc admin tasks that should run as the signed-in admin.
- **OAuth 2.0 refresh-token flow** – Zoho Books and Adobe PDF Services.

## Safety conventions

Scripts that change data follow a few rules: dry-run switches, typed confirmation before destructive actions, CSV backup of previous values, `SystemUpdate` to avoid touching Modified/Editor, Recycle Bin instead of hard delete, and per-item try/catch with a CSV log.

> Comments and console messages are mostly in Bahasa Indonesia.
