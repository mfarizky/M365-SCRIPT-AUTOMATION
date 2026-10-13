# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# KONFIGURASI TARGET SITE & LIST
# ==========================================
$siteUrl     = $Config.HubSiteUrl
$listName    = "MONITOR TAKA"
$tenant      = $Config.TenantDomain
$clientId    = $Config.AppClientId
$pfxPath     = $Config.PfxPath
$pfxPassword = $Config.PfxPassword

# 1. Koneksi ke SharePoint
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Menganalisis data dari '$listName'..." -ForegroundColor Cyan
$allItems = Get-PnPListItem -List $listName -PageSize 2000 -ErrorAction SilentlyContinue

$targetItems = @()

# 2. Filter Item yang KEDUANYA KOSONG
foreach ($item in $allItems) {
    $valIDMK0 = [string]$item["IDMK0"]
    $isIDMK0Empty = [string]::IsNullOrWhiteSpace($valIDMK0)

    $rawIDMK = $item["ID_x0020_MK"]
    if ($null -eq $rawIDMK) { $rawIDMK = $item["ID MK"] }
    
    $valIDMK = ""
    if ($rawIDMK -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
        $valIDMK = [string]$rawIDMK.LookupValue
    } elseif ($rawIDMK) {
        $valIDMK = [string]$rawIDMK
    }
    $isIDMKEmpty = [string]::IsNullOrWhiteSpace($valIDMK)

    if ($isIDMK0Empty -and $isIDMKEmpty) {
        $targetItems += $item
    }
}

$totalTarget = $targetItems.Count

# 3. Tampilkan Hasil Analisis
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "          HASIL PEMERIKSAAN DATA          " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Jumlah item yang IDMK0 & ID MK KEDUANYA KOSONG: $totalTarget row" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan

# Jika tidak ada data yang kosong, hentikan proses
if ($totalTarget -eq 0) {
    Write-Host "Tidak ada data kosong yang ditemukan. Proses dibatalkan secara otomatis." -ForegroundColor Green
    Disconnect-PnPOnline
    exit
}

# 4. Minta Konfirmasi Pengguna
Write-Host ""
$confirmation = Read-Host "Apakah Anda yakin ingin memindahkan $totalTarget item ini ke Recycle Bin? (Ketik 'Y' untuk lanjut / tombol lain untuk batal)"

if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
    Write-Host "`nProses pemindahan ke Recycle Bin dimulai..." -ForegroundColor Cyan
    $deletedCount = 0

    foreach ($item in $targetItems) {
        Write-Host "Memindahkan ke Recycle Bin - Item ID: $($item.Id)..." -ForegroundColor Yellow
        Remove-PnPListItem -List $listName -Identity $item.Id -Recycle -Force
        $deletedCount++
    }

    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "PROSES SELESAI!" -ForegroundColor Green
    Write-Host "Berhasil memindahkan $deletedCount item ke Recycle Bin SharePoint." -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Cyan
} else {
    Write-Host "`nProses DIBATALKAN oleh pengguna. Tidak ada data yang dihapus." -ForegroundColor Magenta
}

Disconnect-PnPOnline