# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==============================
# CONFIGURATION
# ==============================

$clientId = $Config.InteractiveClientId

$baseUrl = "$($Config.SharePointRoot)/sites/"
$validTemplates = @("STS#3","SITEPAGEPUBLISHING#0")

$sourceSiteURL = $Config.TemplatePortalUrl
$listNameToCheck = "REGULATORY COMPLIANCE"

# SITE YANG DIKECUALIKAN
$excludedSites = @("ncp","gki")

# RESULT STORAGE
$results = @()

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
# LOOP CHECK LIST
# ==============================

foreach ($destinationWebURL in $destinationSites) {

    Write-Host "Checking $destinationWebURL" -ForegroundColor Yellow

    try {

        Connect-PnPOnline -Url $destinationWebURL -ClientId $clientId -Interactive

        $list = Get-PnPList -Identity $listNameToCheck -ErrorAction SilentlyContinue

        if ($list) {

            Write-Host "LIST EXISTS di $destinationWebURL" -ForegroundColor Green

            $results += [PSCustomObject]@{
                SiteURL = $destinationWebURL
                ListStatus = "Exists"
            }

        }
        else {

            Write-Host "LIST TIDAK ADA di $destinationWebURL" -ForegroundColor DarkYellow

            $results += [PSCustomObject]@{
                SiteURL = $destinationWebURL
                ListStatus = "Not Found"
            }

        }

    }
    catch {

        Write-Host "FAILED cek di $destinationWebURL" -ForegroundColor Red

        $results += [PSCustomObject]@{
            SiteURL = $destinationWebURL
            ListStatus = "Error"
        }

    }

}

# ==============================
# EXPORT RESULT
# ==============================

$results | Export-Csv -Path "ListAudit.csv" -NoTypeInformation

Write-Host "=== SELESAI CEK ===" -ForegroundColor Cyan
Write-Host "Hasil tersimpan di ListAudit.csv"