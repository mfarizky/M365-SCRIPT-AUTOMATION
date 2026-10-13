# 09 · Intune Endpoint

| Script | What it does |
|---|---|
| `Fix-WSC-Antivirus-Compliance.ps1` | Remediation for devices reported as non-compliant because Windows Security Center doesn't report antivirus status: restarts `wscsvc`, resets the Windows Security app package (SecHealthUI) and restarts the Intune Management Extension to force a fresh compliance check |

Deploy as an Intune **Remediation** (or platform script) running as SYSTEM.
