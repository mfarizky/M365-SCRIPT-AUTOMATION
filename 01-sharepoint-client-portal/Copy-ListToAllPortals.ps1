# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

$clientId = $Config.InteractiveClientId

$baseUrl = "$($Config.SharePointRoot)/sites/"
$validTemplates = @("STS#3","SITEPAGEPUBLISHING#0")

$sourceSiteURL = $Config.TemplatePortalUrl
$SourceListUrl = "$($Config.TemplatePortalUrl)/Lists/REGULATORY%20COMPLIANCE"

# Connect sekali ke source (karena ini sudah terbukti berhasil)
Connect-PnPOnline -Url $sourceSiteURL -ClientId $clientId -Interactive
Write-Host "Connected ke Source"

# Ambil semua destination site sesuai filter
$destinationSites = Get-PnPTenantSite -IncludeOneDriveSites:$false |
Where-Object {
    $siteName = $_.Url.Split("/sites/")[1]

    $_.Url.StartsWith($baseUrl) -and
    $siteName.Length -le 4 -and
    $siteName -ne $Config.TemplatePortalAlias -and
    $validTemplates -contains $_.Template
} |
Select-Object -ExpandProperty Url

# Loop copy
foreach ($destinationWebURL in $destinationSites) {

    Write-Host "Processing $destinationWebURL" -ForegroundColor Cyan

    try {
        Copy-PnPList `
            -SourceListUrl $SourceListUrl `
            -DestinationWebUrl $destinationWebURL `
            -Title "REGULATORY COMPLIANCE"

        Write-Host "SUCCESS copy ke $destinationWebURL" -ForegroundColor Green
    }
    catch {
        Write-Host "FAILED di $destinationWebURL" -ForegroundColor Red
    }
}