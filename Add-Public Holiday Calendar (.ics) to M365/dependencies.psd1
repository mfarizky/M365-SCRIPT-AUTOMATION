# ============================================================
# dependencies.psd1
# Pinned PowerShell module versions for this tool.
# Import-ICSToM365.ps1 reads this automatically.
# To update: change MinVersion, then re-run the script.
# ============================================================
@{
    RequiredModules = @(
        @{
            ModuleName   = "Microsoft.Graph.Authentication"
            MinVersion   = "2.0.0"
            # Latest stable as of 2025-04: 2.19.0
            # https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication
        }
        @{
            ModuleName   = "Microsoft.Graph.Calendar"
            MinVersion   = "2.0.0"
            # Latest stable as of 2025-04: 2.19.0
            # https://www.powershellgallery.com/packages/Microsoft.Graph.Calendar
        }
    )
}
