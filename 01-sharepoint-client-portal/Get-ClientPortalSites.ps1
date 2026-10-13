# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# --- KONFIGURASI ADMIN ---
$adminSiteUrl = $Config.AdminUrl
$clientId = $Config.InteractiveClientId

# Hubungkan ke SharePoint Admin
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId

Write-Host "Mengambil seluruh data site..." -ForegroundColor Cyan

# 1. Tarik semua site collection
$allSites = Get-PnPTenantSite

# 2. Proses Filtering
$filteredSites = $allSites | Where-Object {
    $url = $_.Url
    $title = $_.Title

    # Harus mengandung '/sites/' DAN tidak mengandung kata-kata blacklist (AND Logic)
    (
        $url -like "*/sites/*" -and
        $url -notlike "*content*" -and
        $url -notlike "*app*" -and
        $url -notlike "*VivaHome*" -and
        $url -notlike "*gki*" -and
        $url -notlike "*ganawiki*" -and
        $url -notlike "*portal*" -and
        $title -notlike "*GKI*" -and
        $title -notlike "*PORTAL*" -and
        $title -notlike "*FILL GAPS SITE*" -and
        $title -notlike "*Communication*" -and
        $title -notlike "*$($Config.InternalTitleKeyword)*" -and
        $title -notlike "*RedirectSite*" # <-- Filter Baru Tambahan
    )
}

# --- SUMMARY TOTAL SITE YANG BERHASIL DITARIK ---
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "         SHAREPOINT SITE SUMMARY          " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site Sebelum Filter : $($allSites.Count)" -ForegroundColor Yellow
Write-Host " Total Site Lolos Filter   : $($filteredSites.Count)" -ForegroundColor Green
Write-Host " Total Site Ter-exclude    : $($allSites.Count - $filteredSites.Count)" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan

# Menampilkan list ringkas hasil akhir dalam bentuk tabel
Write-Host "`nDaftar Site Target:" -ForegroundColor Green
$filteredSites | Select-Object @{Name="Display Name"; Expression={$_.Title}}, @{Name="Web URL"; Expression={$_.Url}} | Format-Table -AutoSize

# Memasukkan hasil filter ke variable array untuk kebutuhan script utama nanti
$targetSites = $filteredSites.Url