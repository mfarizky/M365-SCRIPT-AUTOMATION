# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI PATH & THRESHOLD
# ==========================================
$rootPath        = $Config.Compression.PdfScanRoot
$resultDir       = Join-Path $PSScriptRoot "REPORT"
$thresholdLowMB  = 5
$thresholdHighMB = 8
$minTextChars    = 50

$spBaseUrl    = $Config.Compression.PdfSpBaseUrl

$gsPath = Get-Command gswin64c -ErrorAction SilentlyContinue
if (-not $gsPath) {
    Write-Host "Ghostscript (gswin64c) belum terinstall / tidak ada di PATH." -ForegroundColor Red
    return
}

if (!(Test-Path -LiteralPath $resultDir)) {
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
}

$summaryCsv = Join-Path $resultDir "PDF_Summary_Report.csv"

# ==========================================
# FUNCTION: DETEKSI JENIS PDF
# ==========================================
function Get-PdfJenis {
    param([string]$pdfPath)

    $tempTxt = Join-Path $env:TEMP "pdfcheck_$(Get-Random).txt"

    try {
        & gswin64c -q -dNOPAUSE -dBATCH -sDEVICE=txtwrite -o $tempTxt $pdfPath 2>&1 | Out-Null

        if (Test-Path -LiteralPath $tempTxt) {
            $content = Get-Content -LiteralPath $tempTxt -Raw -ErrorAction SilentlyContinue
            $charCount = if ($content) { ($content -replace '\s', '').Length } else { 0 }
            Remove-Item -LiteralPath $tempTxt -ErrorAction SilentlyContinue

            if ($charCount -lt $minTextChars) { return "Scanned (Image Only)" }
            else { return "Ada Teks (Native/OCR)" }
        }
        else { return "Unknown (gagal dibaca)" }
    }
    catch {
        Remove-Item -LiteralPath $tempTxt -ErrorAction SilentlyContinue
        return "Unknown (error)"
    }
}

# ==========================================
# FUNCTION: APPEND 1 BARIS KE CSV LANGSUNG
# ==========================================
function Add-RowToCsv {
    param(
        [PSCustomObject]$Row,
        [string]$Path
    )

    $fileExists = Test-Path -LiteralPath $Path

    if (-not $fileExists) {
        # Belum ada file -> buat baru sekalian dengan header
        $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }
    else {
        # Sudah ada -> append TANPA header
        $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Append
    }
}

# ==========================================
# 2. SCANNING ALL PDF (>= 5MB)
# ==========================================
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "          PDF SUMMARY & SHAREPOINT MAPPER                " -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "Scanning target: $rootPath" -ForegroundColor DarkGray
Write-Host "Kategori 1     : $thresholdLowMB MB - $thresholdHighMB MB" -ForegroundColor DarkGray
Write-Host "Kategori 2     : > $thresholdHighMB MB`n" -ForegroundColor DarkGray

Write-Host "[1/2] Mencari file PDF..." -ForegroundColor Cyan

$allFiles = Get-ChildItem -LiteralPath $rootPath -Filter *.pdf -Recurse -ErrorAction SilentlyContinue | Where-Object { 
    ($_.Length / 1MB) -ge $thresholdLowMB
}

$totalFiles = $allFiles.Count

if ($totalFiles -eq 0) {
    Write-Host "Tidak ditemukan file PDF yang ukurannya di atas $thresholdLowMB MB." -ForegroundColor Green
    return
}

Write-Host "Ditemukan $totalFiles file PDF (di folder, saat scan kali ini).`n" -ForegroundColor Green

# ==========================================
# 3. CEK CSV LAMA -> ambil daftar Local Path yang SUDAH tercatat
# ==========================================
$existingPaths = New-Object System.Collections.Generic.HashSet[string]
$existingCount = 0

if (Test-Path -LiteralPath $summaryCsv) {
    $existingRows = Import-Csv -LiteralPath $summaryCsv -Delimiter ";"
    foreach ($row in $existingRows) {
        [void]$existingPaths.Add($row."Local Path")
    }
    $existingCount = $existingRows.Count
    Write-Host "CSV lama ditemukan. Baris yang sudah tercatat: $existingCount`n" -ForegroundColor DarkGray
}

# ==========================================
# 4. PROSES FILE BARU - LANGSUNG APPEND PER FILE
# ==========================================
Write-Host "[2/2] Memproses file baru (deteksi jenis PDF + append CSV per file)..." -ForegroundColor Cyan

$newCount = 0
$fileCounter = 0
$filesToCheck = $allFiles | Where-Object { -not $existingPaths.Contains($_.FullName) }
$totalToCheck = $filesToCheck.Count

if ($totalToCheck -eq 0) {
    Write-Host "Semua file di folder sudah tercatat di CSV. Tidak ada yang baru." -ForegroundColor Green
    return
}

foreach ($file in $filesToCheck) {
    $fileCounter++
    $sizeMB = [math]::Round($file.Length / 1MB, 2)
    $localPath = $file.FullName

    Write-Host "  [$fileCounter/$totalToCheck] Deteksi jenis: $($file.Name)" -ForegroundColor DarkCyan
    $jenisPdf = Get-PdfJenis -pdfPath $localPath

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
        "Jenis PDF"                        = $jenisPdf
        "Status"                           = ""
        "Ukuran Setelah Compressed (MB)"  = ""
        "Kategori"                         = $kategori
        "Local Path"                       = $localPath
        "SharePoint URL"                   = $sharePointUrl
    }

    # LANGSUNG APPEND KE CSV -- baris ini yang bikin riwayat tetap ada walau berhenti di tengah
    Add-RowToCsv -Row $item -Path $summaryCsv

    Write-Host "    -> [$jenisPdf] tersimpan ke CSV" -ForegroundColor DarkGray

    $newCount++
}

# ==========================================
# 5. RINGKASAN AKHIR (baca ulang CSV final utk hitung total)
# ==========================================
$finalList = Import-Csv -LiteralPath $summaryCsv -Delimiter ";"

$countLow        = ($finalList | Where-Object { $_.Kategori -eq "$thresholdLowMB MB - $thresholdHighMB MB" }).Count
$countHigh       = ($finalList | Where-Object { $_.Kategori -eq "> $thresholdHighMB MB" }).Count
$countScanned    = ($finalList | Where-Object { $_."Jenis PDF" -eq "Scanned (Image Only)" }).Count
$countNative     = ($finalList | Where-Object { $_."Jenis PDF" -eq "Ada Teks (Native/OCR)" }).Count
$countCompressed = ($finalList | Where-Object { $_.Status -eq "Compressed" }).Count
$countPending    = ($finalList | Where-Object { $_.Status -ne "Compressed" }).Count

Write-Host "`n=========================================================" -ForegroundColor Yellow
Write-Host " Ringkasan Selesai!" -ForegroundColor Green
Write-Host " File PDF ditemukan di folder (scan kali ini) : $totalFiles file" -ForegroundColor White
Write-Host " Baris baru ditambahkan ke CSV                : $newCount file" -ForegroundColor White
Write-Host " Total file di CSV (keseluruhan)               : $($finalList.Count) file" -ForegroundColor White
Write-Host " Kategori $thresholdLowMB-$thresholdHighMB MB                             : $countLow file" -ForegroundColor White
Write-Host " Kategori > $thresholdHighMB MB                              : $countHigh file" -ForegroundColor White
Write-Host " Jenis Scanned (Image Only)                    : $countScanned file" -ForegroundColor White
Write-Host " Jenis Ada Teks (Native/OCR)                   : $countNative file" -ForegroundColor White
Write-Host " Sudah Compressed                              : $countCompressed file" -ForegroundColor Green
Write-Host " Belum Compressed (Pending)                    : $countPending file" -ForegroundColor DarkYellow
Write-Host " Report CSV                                    : $summaryCsv" -ForegroundColor White
Write-Host "=========================================================" -ForegroundColor Yellow