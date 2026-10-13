# Configuration

No script in this repository contains credentials. Environment-specific values live in `config.local.ps1`, which is listed in `.gitignore`.

| File | Committed | Purpose |
|---|---|---|
| `config.example.ps1` | ✅ | Template with placeholder values |
| `config.local.ps1` | ❌ | Your real values (create from the example) |
| `Load-Config.ps1` | ✅ | Loader dot-sourced by every script; builds `$Config` |

## Setup

```powershell
Copy-Item .\config\config.example.ps1 .\config\config.local.ps1
```

Fill in tenant, app registration and paths. For the certificate password, prefer an environment variable over the file:

```powershell
$env:PNP_PFX_PASSWORD = "<pfx password>"
```

`Load-Config.ps1` converts it to a `SecureString` (`$Config.PfxPassword`) and removes the plain-text copy from memory.

## App registration (app-only)

1. Register an app in Entra ID and upload the public certificate (`.cer`).
2. Grant application permissions, e.g. SharePoint `Sites.FullControl.All`, Graph `User.Read.All`, `MailboxSettings.Read`.
3. Keep the `.pfx` **outside** the repository and point `PfxPath` to it.
