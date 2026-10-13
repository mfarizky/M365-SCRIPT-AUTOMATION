# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
# ==========================================
# KOMPRES IMAGE - OVERWRITE IN-PLACE (AMAN UNTUK ONEDRIVE VERSION HISTORY)
# ==========================================
$csvPath     = Join-Path $PSScriptRoot "REPORT\IMG_Summary_Report.csv"   # output dari Scan-LargeImages.ps1
$jpgQuality  = 82
$batchSize   = 250

# ==========================================
# 1. VALIDASI AWAL
# ==========================================
$magickPath = Get-Command magick -ErrorAction SilentlyContinue
if (-not $magickPath) {
    Write-Host "ImageMagick belum terinstall / tidak ada di PATH. Install dulu: winget install x.ImageMagick" -ForegroundColor Red
    return
}

if (!(Test-Path -LiteralPath $csvPath)) {
    Write-Host "File CSV hasil scan tidak ditemukan: $csvPath" -ForegroundColor Red
    return
}

# ==========================================
# 2. BACA CSV & SUSUN ULANG KOLOM
# ==========================================
$rawRows = Import-Csv -LiteralPath $csvPath -Delimiter ";"

# Urutan kolom: Status & Ukuran Setelah ditaruh SEBELUM Kategori
$allRows = foreach ($row in $rawRows) {
    [PSCustomObject]@{
        "Nama File"                       = $row."Nama File"
        "Ekstensi"                         = $row."Ekstensi"
        "Ukuran (MB)"                       = $row."Ukuran (MB)"
        "Status"                            = if ($row.PSObject.Properties.Name -contains "Status") { $row.Status } else { "" }
        "Ukuran Setelah Compressed (MB)"   = if ($row.PSObject.Properties.Name -contains "Ukuran Setelah Compressed (MB)") { $row."Ukuran Setelah Compressed (MB)" } else { "" }
        "Kategori"                          = $row.Kategori
        "Local Path"                        = $row."Local Path"
        "SharePoint URL"                    = $row."SharePoint URL"
    }
}

# ==========================================
# 3. FILTER: HANYA AMBIL YANG BELUM "Compressed"
# ==========================================
$pendingRows = $allRows | Where-Object { $_.Status -ne "Compressed" }

Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "         KOMPRES IMAGE - OVERWRITE IN-PLACE               " -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "Total file di CSV   : $($allRows.Count)" -ForegroundColor DarkGray
Write-Host "Sudah Compressed    : $($allRows.Count - $pendingRows.Count)" -ForegroundColor DarkGray
Write-Host "Belum diproses       : $($pendingRows.Count)" -ForegroundColor DarkGray
Write-Host "Batch kali ini       : $batchSize file`n" -ForegroundColor Yellow

$filesToProcess = $pendingRows | Select-Object -First $batchSize
$total = $filesToProcess.Count

if ($total -eq 0) {
    Write-Host "Tidak ada file yang perlu diproses (semua sudah Compressed / habis)." -ForegroundColor Green
    return
}

$counter = 0
$results = [System.Collections.Generic.List[PSObject]]::new()

# ==========================================
# 4. PROSES TIAP FILE
# ==========================================
foreach ($row in $filesToProcess) {
    $counter++
    $localPath = $row."Local Path"

    # --- Cek file ada ---
    if (!(Test-Path -LiteralPath $localPath)) {
        Write-Host "[$counter/$total] SKIP (file tidak ditemukan): $localPath" -ForegroundColor DarkYellow
        $row.Status = "Not Found"
        continue
    }

    $file = Get-Item -LiteralPath $localPath

    # --- Cek file masih cloud-only (belum di-download / placeholder OneDrive) ---
    if ($file.Attributes -match "Offline" -or $file.Attributes -match "RecallOnDataAccess") {
        Write-Host "[$counter/$total] SKIP (file masih cloud-only, belum di-download): $($file.Name)" -ForegroundColor DarkYellow
        $row.Status = "Skipped (cloud-only)"
        continue
    }

    Write-Host "[$counter/$total] Processing: $($file.Name)" -ForegroundColor Cyan

    # --- Cek ekstensi didukung ---
    if ($file.Extension -notmatch "jpe?g|png") {
        Write-Host "  Format tidak dikenali, dilewati." -ForegroundColor DarkYellow
        $row.Status = "Skipped (unknown format)"
        continue
    }

    $sizeBefore = [math]::Round($file.Length / 1MB, 2)
    $tempPath   = Join-Path $file.DirectoryName "$($file.BaseName)_temp_$(Get-Random)$($file.Extension)"

    # --- 4a. COMPRESS ke TEMP FILE dulu ---
    try {
        if ($file.Extension -match "jpe?g") {
            & magick $file.FullName -strip -quality $jpgQuality $tempPath 2>&1 | Out-Null
        }
        elseif ($file.Extension -eq ".png") {
            & magick $file.FullName -strip -define png:compression-level=9 -define png:compression-filter=5 $tempPath 2>&1 | Out-Null
        }
    }
    catch {
        Write-Host "  GAGAL compress: $($file.Name) — $($_.Exception.Message)" -ForegroundColor Red
        $row.Status = "Failed (compress error)"
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
        continue
    }

    # --- 4b. VALIDASI hasil compress ---
    $isValid = $false
    if (Test-Path -LiteralPath $tempPath) {
        $tempSize = (Get-Item -LiteralPath $tempPath).Length
        if ($tempSize -gt 0) {
            & magick identify $tempPath 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) {
                $isValid = $true
            }
        }
    }

    if (-not $isValid) {
        Write-Host "  GAGAL/CORRUPT hasil compress — file asli TIDAK disentuh: $($file.Name)" -ForegroundColor Red
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
        $row.Status = "Failed (corrupt output)"
        continue
    }

    # --- 4c. Kalau hasil compress malah LEBIH BESAR, jangan timpa (percuma) ---
    if ($tempSize -ge $file.Length) {
        Write-Host "  Hasil compress tidak lebih kecil, file asli DIPERTAHANKAN: $($file.Name)" -ForegroundColor DarkYellow
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
        $row.Status = "Skipped (no size benefit)"
        continue
    }

    # --- 4d. OVERWRITE IN-PLACE (path & nama TETAP SAMA -> aman utk version history) ---
    try {
        Copy-Item -LiteralPath $tempPath -Destination $file.FullName -Force -ErrorAction Stop
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue

        $sizeAfter = [math]::Round((Get-Item -LiteralPath $file.FullName).Length / 1MB, 2)
        $savedPct  = [math]::Round((1 - ($sizeAfter / $sizeBefore)) * 100, 1)

        $row.Status = "Compressed"
        $row."Ukuran Setelah Compressed (MB)" = $sizeAfter

        Write-Host "  Berhasil overwrite in-place ($sizeBefore MB -> $sizeAfter MB, hemat $savedPct%)" -ForegroundColor Green

        $results.Add([PSCustomObject]@{
            "Nama File"    = $file.Name
            "Sebelum (MB)" = $sizeBefore
            "Setelah (MB)" = $sizeAfter
            "Hemat (%)"    = $savedPct
        })
    }
    catch {
        Write-Host "  GAGAL overwrite (kemungkinan file locked/in-use): $($file.Name) — $($_.Exception.Message)" -ForegroundColor Red
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
        $row.Status = "Failed (locked/in-use)"
    }
}

# ==========================================
# 5. SIMPAN ULANG CSV (STATUS & UKURAN TERBARU)
# ==========================================
$allRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

# ==========================================
# 6. RINGKASAN
# ==========================================
Write-Host "`n=========================================================" -ForegroundColor Yellow
if ($results.Count -gt 0) {
    $results | Format-Table "Nama File", "Sebelum (MB)", "Setelah (MB)", "Hemat (%)" -AutoSize

    $totalBefore = [math]::Round(($results."Sebelum (MB)" | Measure-Object -Sum).Sum, 2)
    $totalAfter  = [math]::Round(($results."Setelah (MB)" | Measure-Object -Sum).Sum, 2)

    Write-Host "Total sebelum : $totalBefore MB" -ForegroundColor White
    Write-Host "Total setelah : $totalAfter MB" -ForegroundColor Green
    Write-Host "Hemat total   : $([math]::Round((1-($totalAfter/$totalBefore))*100,1))%" -ForegroundColor Green
} else {
    Write-Host "Tidak ada file yang berhasil dikompres di batch ini." -ForegroundColor DarkYellow
}
Write-Host " CSV sudah diupdate  : $csvPath" -ForegroundColor White
Write-Host "=========================================================" -ForegroundColor Yellow

$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed
Write-Host "Runtime            : $($elapsed.Hours)j $($elapsed.Minutes)m $($elapsed.Seconds)d $($elapsed.Milliseconds)ms" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Yellow