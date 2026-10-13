# Restart Windows Security Center Service
Get-Service -Name wscsvc | Restart-Service -Force

# Reset Security Health UI App Package
Get-AppxPackage -Name Microsoft.SecHealthUI -AllUsers | Reset-AppxPackage

# Force Intune Management Extension Device Sync
Get-Service -Name IntuneManagementExtension | Restart-Service -Force