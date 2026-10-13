# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI PATH & PARAMETER
# ==========================================
$gsPath          = $Config.Compression.GhostscriptPath

# Path Induk Corporate
$corporateRoot   = $Config.Compression.PdfCorporateRoot
$backupCorporate = $Config.Compression.PdfBackupRoot

$resultDir       = Join-Path $PSScriptRoot "RESULT"
$thresholdMB    = 5
$limitFiles      = 190  # Batasan batch per run

# Pastikan folder RESULT dan BACKUP CORPORATE ada
if (!(Test-Path -LiteralPath $resultDir)) { New-Item -ItemType Directory -Path $resultDir -Force | Out-Null }
if (!(Test-Path -LiteralPath $backupCorporate)) { New-Item -ItemType Directory -Path $backupCorporate -Force | Out-Null }

# File CSV Tetap (Statis)
$pendingLogFile   = Join-Path $resultDir "Compress_Pending_Queue.csv"
$executionLogFile = Join-Path $resultDir "Compress_Execution_Report.csv"

# Inisialisasi Header CSV jika file belum pernah ada
if (!(Test-Path -LiteralPath $pendingLogFile)) {
    "Waktu_Didaftarkan;Nama_File;Hierarchy_Folder;Ukuran_Awal_MB;Local_Path" | Out-File -LiteralPath $pendingLogFile -Encoding UTF8
}

if (!(Test-Path -LiteralPath $executionLogFile)) {
    "Waktu_Eksekusi;Status;Nama_File;Hierarchy_Folder;Ukuran_Awal_MB;Ukuran_Akhir_MB;Backup_Path;Local_Path" | Out-File -LiteralPath $executionLogFile -Encoding UTF8
}

# PERBAIKAN: -ForegroundColor diubah dari Header ke Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "     BATCH PDF COMPRESSION (ROW-BY-ROW LOGGING)          " -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "Target Induk Scan : $corporateRoot" -ForegroundColor DarkGray
Write-Host "Target Induk Backup: $backupCorporate" -ForegroundColor DarkGray
Write-Host "Batas Ukuran File : > $thresholdMB MB" -ForegroundColor DarkGray
Write-Host "CSV Antrean       : $pendingLogFile" -ForegroundColor DarkGray
Write-Host "CSV Eksekusi      : $executionLogFile`n" -ForegroundColor DarkGray

# ==========================================
# 2. SCANNING & FILTER FILE UNPROCESSED ONLY
# ==========================================
Write-Host "[1/3] Memproses pencarian file PDF > $thresholdMB MB yang BELUM pernah dikompresi..." -ForegroundColor Cyan

$rawFiles = Get-ChildItem -LiteralPath $corporateRoot -Filter *.pdf -Recurse -ErrorAction SilentlyContinue | Where-Object { 
    ($_.Length / 1MB) -gt $thresholdMB
}

$unprocessedFiles = [System.Collections.Generic.List[PSObject]]::new()

# Saring file yang BELUM ADA di folder backup
foreach ($file in $rawFiles) {
    $relativeFromCorp = $file.FullName.Substring($corporateRoot.Length).TrimStart('\')
    $ptFolderName     = $relativeFromCorp.Split('\')[0]
    $targetBackupFile = Join-Path $backupCorporate (Join-Path $ptFolderName $file.Name)

    if (!(Test-Path -LiteralPath $targetBackupFile)) {
        $unprocessedFiles.Add($file)
    }
}

$allFiles = $unprocessedFiles | Select-Object -First $limitFiles
$totalFiles = $allFiles.Count

if ($totalFiles -eq 0) {
    Write-Host "Tidak ditemukan file PDF baru/belum dikompresi di atas $thresholdMB MB. Semua sudah selesai!" -ForegroundColor Green
    return
}

Write-Host "Ditemukan $totalFiles file PDF BARU untuk diproses pada batch ini.`n" -ForegroundColor Green

# ==========================================
# 3. PROSES SATU PER SATU (ADD PENDING -> COMPRESS -> ADD EXECUTION)
# ==========================================
Write-Host "[2/3] Memulai proses bertahap per file...`n" -ForegroundColor Cyan

$currentIndex = 0

foreach ($file in $allFiles) {
    $currentIndex++
    $fileSizeMB = [math]::Round($file.Length / 1MB, 2)
    $sourcePath = $file.FullName
    $startTime  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Format Hierarchy Folder
    $folderOnly = $file.DirectoryName
    $marker     = $Config.Compression.OneDriveMarker
    if ($folderOnly.Contains($marker)) {
        $cleanPath = $folderOnly.Substring($folderOnly.IndexOf($marker) + $marker.Length)
    } else {
        $cleanPath = $folderOnly
    }
    $hierarchyDir = $cleanPath.Replace("\", " > ")

    # --- TAHAP A: DAFTARKAN FILE INI SAJA KE CSV PENDING ---
    "$startTime;$($file.Name);`"$hierarchyDir`";$fileSizeMB;`"$sourcePath`"" | Out-File -LiteralPath $pendingLogFile -Append -Encoding UTF8

    # Indikator Progress
    $percent = [math]::Round(($currentIndex / $totalFiles) * 100, 0)
    Write-Progress -Activity "Batch PDF Compression" -Status "[$currentIndex/$totalFiles] $($file.Name)" -PercentComplete $percent

    Write-Host "[$currentIndex/$totalFiles] " -NoNewline -ForegroundColor White
    Write-Host "Memproses: " -NoNewline -ForegroundColor Gray
    Write-Host "$($file.Name) " -NoNewline -ForegroundColor Magenta
    Write-Host "($fileSizeMB MB)" -ForegroundColor Yellow
    Write-Host "   -> Logged to Pending CSV" -ForegroundColor DarkGray
    
    # Deteksi PT dan Path Backup Flat
    $relativeFromCorp = $sourcePath.Substring($corporateRoot.Length).TrimStart('\')
    $ptFolderName     = $relativeFromCorp.Split('\')[0]
    $ptBackupFolder   = Join-Path $backupCorporate $ptFolderName
    
    if (!(Test-Path -LiteralPath $ptBackupFolder)) {
        New-Item -ItemType Directory -Path $ptBackupFolder -Force | Out-Null
    }

    $backupPath = Join-Path $ptBackupFolder $file.Name

    # --- TAHAP B: PROSES COPY BACKUP KE FOLDER PT ---
    try {
        Copy-Item -LiteralPath $sourcePath -Destination $backupPath -Force
        Write-Host "   [1/2] Backup berhasil -> $backupPath" -ForegroundColor DarkGray
    } catch {
        Write-Host "   !! GAGAL BACKUP: ($($_.Exception.Message)) - Melewati Kompresi!" -ForegroundColor Red
        "$startTime;GAGAL_BACKUP;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        continue
    }

    # --- TAHAP C: PROSES KOMPRESI DENGAN FILE TEMP SEBELUM OVERWRITE ---
    $tempPath = Join-Path $file.DirectoryName ("temp_gs_compress_" + $file.Name)

    $gsArgs = "-sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook " +
              "-dEmbedAllFonts=true -dSubsetFonts=true -dNOPAUSE -dQUIET -dBATCH " +
              "-sOutputFile=`"$tempPath`" `"$sourcePath`""

    Start-Process -FilePath $gsPath -ArgumentList $gsArgs -Wait -NoNewWindow

    # --- TAHAP D: VERIFIKASI, OVERWRITE & CATAT HASIL EKSEKUSI ---
    if ((Test-Path -LiteralPath $tempPath) -and ((Get-Item -LiteralPath $tempPath).Length -gt 0)) {
        $newSizeMB = [math]::Round((Get-Item -LiteralPath $tempPath).Length / 1MB, 2)

        try {
            Remove-Item -LiteralPath $sourcePath -Force -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $sourcePath -Force -ErrorAction Stop
            
            Write-Host "   [2/2] SUKSES! File Ditimpa ($fileSizeMB MB -> $newSizeMB MB)" -ForegroundColor Green
            
            # Catat Hasil ke CSV Eksekusi (SUKSES)
            "$startTime;SUKSES;$($file.Name);`"$hierarchyDir`";$fileSizeMB;$newSizeMB;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        } catch {
            Write-Host "   !! GAGAL OVERWRITE: File terkunci OneDrive/System ($($_.Exception.Message))" -ForegroundColor Red
            "$startTime;GAGAL_OVERWRITE;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        }
    } else {
        Write-Host "   [2/2] GAGAL KOMPRESI!" -ForegroundColor Red
        "$startTime;GAGAL_COMPRESS;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
    }

    # Cleanup Temp
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }

    Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
}

Write-Host "`n[SELESAI] Seluruh batch diproses." -ForegroundColor Yellow
Write-Host "Log Antrean PENDING : $pendingLogFile" -ForegroundColor White
Write-Host "Log Hasil EKSEKUSI  : $executionLogFile" -ForegroundColor White