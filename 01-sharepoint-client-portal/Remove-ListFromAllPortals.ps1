# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==============================
# CONFIGURATION
# ==============================

$clientId = $Config.InteractiveClientId

$baseUrl = "$($Config.SharePointRoot)/sites/"
$validTemplates = @("STS#3","SITEPAGEPUBLISHING#0")

$sourceSiteURL = $Config.TemplatePortalUrl
$listNameToDelete = "REGULATORY COMPLIANCE"

# SITE YANG DIKECUALIKAN
$excludedSites = @($Config.TemplatePortalAlias, $Config.InternalSiteAlias)

# ==============================
# CONNECT KE SOURCE
# ==============================

Connect-PnPOnline -Url $sourceSiteURL -ClientId $clientId -Interactive
Write-Host "Connected ke Source Site" -ForegroundColor Green

# ==============================
# AMBIL DESTINATION SITES
# ==============================

$destinationSites = Get-PnPTenantSite -IncludeOneDriveSites:$false |
Where-Object {

    $siteName = $_.Url.Split("/sites/")[1]

    $_.Url.StartsWith($baseUrl) -and
    $siteName.Length -le 4 -and
    -not ($excludedSites -contains $siteName) -and
    $validTemplates -contains $_.Template

} |
Select-Object -ExpandProperty Url

Write-Host "Total Sites ditemukan: $($destinationSites.Count)" -ForegroundColor Cyan

# ==============================
# LOOP DELETE LIST
# ==============================

foreach ($destinationWebURL in $destinationSites) {

    Write-Host "Processing $destinationWebURL" -ForegroundColor Yellow

    try {

        Connect-PnPOnline -Url $destinationWebURL -ClientId $clientId -Interactive

        $list = Get-PnPList -Identity $listNameToDelete -ErrorAction SilentlyContinue

        if ($list) {
            Remove-PnPList -Identity $listNameToDelete -Force
            Write-Host "SUCCESS delete di $destinationWebURL" -ForegroundColor Green
        }
        else {
            Write-Host "List tidak ditemukan di $destinationWebURL" -ForegroundColor DarkGray
        }

    }
    catch {
        Write-Host "FAILED di $destinationWebURL" -ForegroundColor Red
    }
}

Write-Host "=== SELESAI ===" -ForegroundColor Cyan