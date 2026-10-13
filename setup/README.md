# Setup

`Install-Prerequisites.ps1` – sets the execution policy to `RemoteSigned` for the current user and installs `PnP.PowerShell` if missing.

Other modules used in this repo:

```powershell
Install-Module Microsoft.Graph            -Scope CurrentUser
Install-Module ExchangeOnlineManagement   -Scope CurrentUser
Install-Module ImportExcel                -Scope CurrentUser
```
