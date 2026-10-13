# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

$ClientId = $Config.InteractiveClientId

# URL source & target
$sourceUrl = $Config.TemplatePortalUrl
$targetUrl = "$($Config.SharePointRoot)/sites/clienttest"

# Koneksi ke source site
Connect-PnPOnline -Url $sourceUrl -ClientId $ClientId
$sourceCon = Get-PnPConnection

# Koneksi ke target site
Connect-PnPOnline -Url $targetUrl -ClientId $ClientId
$targetCon = Get-PnPConnection

# Nama halaman yang mau diexport
$pageName = "SiteTemplate.aspx"

# 🔍 Cek webpart / komponen di halaman
Write-Host "🔍Komponen di halaman $pageName dari $sourceUrl"
Get-PnPPageComponent -Connection $sourceCon -Page $pageName | ForEach-Object {
    Write-Host ("- " + $_.Title + " (" + $_.Type + ")")
}

# File sementara
$tempFile = [System.IO.Path]::GetTempFileName()

# 📤 Export halaman dengan semua komponen
Write-Host "📤 Exporting $pageName..."
Export-PnPPage -Connection $sourceCon -Identity $pageName -Out $tempFile -Force 

# 📥 Apply ke target site
Write-Host "📥 Applying $pageName ke $targetUrl..."
Invoke-PnPSiteTemplate -Connection $targetCon -Path $tempFile

# Bersihkan file sementara
Remove-Item $tempFile -Force

Write-Host "✅ Page $pageName berhasil di-copy dan semua komponen ikut terbawa."
