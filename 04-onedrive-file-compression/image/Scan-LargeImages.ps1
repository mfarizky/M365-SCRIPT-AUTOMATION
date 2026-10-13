# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI PATH & THRESHOLD
# ==========================================
$rootPath        = $Config.Compression.ImageScanRoot
$resultDir       = Join-Path $PSScriptRoot "REPORT"
$thresholdLowMB  = 5
$thresholdHighMB = 8

# Ekstensi gambar yang mau di-scan
$imageExtensions = @("*.png", "*.jpg", "*.jpeg")

# Konfigurasi Domain & Site SharePoint
$spBaseUrl    = $Config.Compression.ImageSpBaseUrl

# Pastikan folder RESULT ada
if (!(Test-Path -LiteralPath $resultDir)) {
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
}

# NAMA CSV TETAP (tidak pakai timestamp lagi, supaya bisa di-append terus)
$summaryCsv = Join-Path $resultDir "IMG_Summary_Report.csv"

# ==========================================
# 2. SCANNING ALL IMAGES (>= 5MB)
# ==========================================
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "         IMAGE SUMMARY & SHAREPOINT MAPPER               " -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "Scanning target: $rootPath" -ForegroundColor DarkGray
Write-Host "Ekstensi       : $($imageExtensions -join ', ')" -ForegroundColor DarkGray
Write-Host "Kategori 1     : $thresholdLowMB MB - $thresholdHighMB MB" -ForegroundColor DarkGray
Write-Host "Kategori 2     : > $thresholdHighMB MB`n" -ForegroundColor DarkGray

Write-Host "[1/2] Mencari file gambar..." -ForegroundColor Cyan

$allFiles = Get-ChildItem -LiteralPath $rootPath -Include $imageExtensions -Recurse -ErrorAction SilentlyContinue | Where-Object { 
    ($_.Length / 1MB) -ge $thresholdLowMB
}

$totalFiles = $allFiles.Count

if ($totalFiles -eq 0) {
    Write-Host "Tidak ditemukan file gambar yang ukurannya di atas $thresholdLowMB MB." -ForegroundColor Green
    return
}

Write-Host "Ditemukan $totalFiles file gambar (hasil scan kali ini).`n" -ForegroundColor Green

# ==========================================
# 3. KUMPULKAN DATA HASIL SCAN & CONVERT TO SHAREPOINT URL
# ==========================================
Write-Host "[2/2] Memetakan Path Lokal ke URL SharePoint..." -ForegroundColor Cyan

$scannedList = [System.Collections.Generic.List[PSObject]]::new()

foreach ($file in $allFiles) {
    $sizeMB = [math]::Round($file.Length / 1MB, 2)
    $localPath = $file.FullName

    # Tentukan kategori berdasarkan ukuran
    if ($sizeMB -ge $thresholdHighMB) {
        $kategori = "> $thresholdHighMB MB"
    } else {
        $kategori = "$thresholdLowMB MB - $thresholdHighMB MB"
    }

    $relativePath = $localPath.Substring($rootPath.Length).Replace("\", "/")
    $encodedRelative = [System.Uri]::EscapeUriString($relativePath)
    $sharePointUrl   = "$spBaseUrl$encodedRelative"

    $item = [PSCustomObject]@{
        "Nama File"                      = $file.Name
        "Ekstensi"                        = $file.Extension
        "Ukuran (MB)"                      = $sizeMB
        "Status"                           = ""
        "Ukuran Setelah Compressed (MB)"  = ""
        "Kategori"                         = $kategori
        "Local Path"                       = $localPath
        "SharePoint URL"                   = $sharePointUrl
    }

    $scannedList.Add($item)
}

# ==========================================
# 4. GABUNGKAN DENGAN CSV LAMA (KALAU ADA) - APPEND, TANPA DUPLIKAT
# ==========================================
$finalList = [System.Collections.Generic.List[PSObject]]::new()

if (Test-Path -LiteralPath $summaryCsv) {
    Write-Host "`nCSV lama ditemukan, menggabungkan (append) tanpa duplikat..." -ForegroundColor Cyan
    $existingRows = Import-Csv -LiteralPath $summaryCsv -Delimiter ";"

    # Buat lookup set dari Local Path yang SUDAH ada di CSV lama
    $existingPaths = New-Object System.Collections.Generic.HashSet[string]
    foreach ($row in $existingRows) {
        [void]$existingPaths.Add($row."Local Path")
        $finalList.Add($row)   # baris lama tetap dipertahankan APA ADANYA (termasuk Status-nya)
    }

    # Hanya tambahkan file hasil scan baru yang BELUM ada di CSV lama
    $newCount = 0
    foreach ($item in $scannedList) {
        if (-not $existingPaths.Contains($item."Local Path")) {
            $finalList.Add($item)
            $newCount++
        }
    }

    Write-Host "Baris lama dipertahankan : $($existingRows.Count)" -ForegroundColor DarkGray
    Write-Host "Baris baru ditambahkan   : $newCount" -ForegroundColor Green
}
else {
    Write-Host "`nCSV belum ada, membuat baru..." -ForegroundColor Cyan
    foreach ($item in $scannedList) {
        $finalList.Add($item)
    }
}

# ==========================================
# 5. TAMPILKAN TABEL & EKSPOR CSV
# ==========================================
$finalList = $finalList | Sort-Object "Local Path"

$finalList | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"

$finalList | Format-Table "Nama File", "Ekstensi", "Ukuran (MB)", "Status", "Kategori", "Local Path" -AutoSize

# Hitung jumlah per kategori (dari keseluruhan CSV setelah digabung)
$countLow  = ($finalList | Where-Object { $_.Kategori -eq "$thresholdLowMB MB - $thresholdHighMB MB" }).Count
$countHigh = ($finalList | Where-Object { $_.Kategori -eq "> $thresholdHighMB MB" }).Count

# Hitung jumlah berdasarkan Status compress
$countCompressed = ($finalList | Where-Object { $_.Status -eq "Compressed" }).Count
$countPending     = ($finalList | Where-Object { $_.Status -ne "Compressed" }).Count

Write-Host "`n=========================================================" -ForegroundColor Yellow
Write-Host " Ringkasan Selesai!" -ForegroundColor Green
Write-Host " Total file hasil scan kali ini : $totalFiles file" -ForegroundColor White
Write-Host " Total file di CSV (keseluruhan): $($finalList.Count) file" -ForegroundColor White
Write-Host " Kategori $thresholdLowMB-$thresholdHighMB MB              : $countLow file" -ForegroundColor White
Write-Host " Kategori > $thresholdHighMB MB               : $countHigh file" -ForegroundColor White
Write-Host " Sudah Compressed                : $countCompressed file" -ForegroundColor Green
Write-Host " Belum Compressed (Pending)      : $countPending file" -ForegroundColor DarkYellow
Write-Host " Report CSV                      : $summaryCsv" -ForegroundColor White
Write-Host "=========================================================" -ForegroundColor Yellow