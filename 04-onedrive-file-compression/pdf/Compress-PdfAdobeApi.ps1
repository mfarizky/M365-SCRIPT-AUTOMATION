# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI KREDENSIAL, PATH & PARAMETER
# ==========================================
# Client ID dan Client Secret Adobe PDF Services
$clientId     = $Config.Adobe.ClientId
$clientSecret = $Config.Adobe.ClientSecret

# Path Induk Corporate & Backup
$corporateRoot   = $Config.Compression.PdfCorporateRoot
$backupCorporate = $Config.Compression.PdfBackupRootAdobe

$resultDir       = Join-Path $PSScriptRoot "RESULT"
$thresholdMB    = 5
$limitFiles      = 100  # Batasan batch per run

# Pastikan folder RESULT dan BACKUP CORPORATE ada
if (!(Test-Path -LiteralPath $resultDir)) { New-Item -ItemType Directory -Path $resultDir -Force | Out-Null }
if (!(Test-Path -LiteralPath $backupCorporate)) { New-Item -ItemType Directory -Path $backupCorporate -Force | Out-Null }

# File CSV Log Statis
$pendingLogFile   = Join-Path $resultDir "Compress_Pending_Queue.csv"
$executionLogFile = Join-Path $resultDir "Compress_Execution_Report.csv"

# Inisialisasi Header CSV jika file belum pernah ada
if (!(Test-Path -LiteralPath $pendingLogFile)) {
    "Waktu_Didaftarkan;Nama_File;Hierarchy_Folder;Ukuran_Awal_MB;Local_Path" | Out-File -LiteralPath $pendingLogFile -Encoding UTF8
}

if (!(Test-Path -LiteralPath $executionLogFile)) {
    "Waktu_Eksekusi;Status;Nama_File;Hierarchy_Folder;Ukuran_Awal_MB;Ukuran_Akhir_MB;Backup_Path;Local_Path" | Out-File -LiteralPath $executionLogFile -Encoding UTF8
}

Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "  BATCH PDF COMPRESSION (ADOBE PDF SERVICES - MEDIUM)    " -ForegroundColor Yellow
Write-Host "=========================================================" -ForegroundColor Yellow
Write-Host "Target Induk Scan : $corporateRoot" -ForegroundColor DarkGray
Write-Host "Target Induk Backup: $backupCorporate" -ForegroundColor DarkGray
Write-Host "Batas Ukuran File : > $thresholdMB MB" -ForegroundColor DarkGray
Write-Host "CSV Antrean       : $pendingLogFile" -ForegroundColor DarkGray
Write-Host "CSV Eksekusi      : $executionLogFile`n" -ForegroundColor DarkGray

# ==========================================
# 2. OTENTIKASI ADOBE PDF SERVICES (OAUTH)
# ==========================================
Write-Host "[1/4] Mendapatkan Access Token dari Adobe..." -ForegroundColor Cyan
$authUri = "https://pdf-services.adobe.io/token"
$authBody = @{
    client_id     = $clientId
    client_secret = $clientSecret
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $authUri -Method Post -Body $authBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    $accessToken = $tokenResponse.access_token
    Write-Host "-> Otentikasi Adobe Berhasil!`n" -ForegroundColor Green
} catch {
    Write-Host "-> ERR: Gagal otentikasi Adobe. Cek kembali Client ID & Secret Anda." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

$adobeHeaders = @{
    "Authorization" = "Bearer $accessToken"
    "x-api-key"     = $clientId
}

# ==========================================
# 3. SCANNING & FILTER FILE UNPROCESSED ONLY
# ==========================================
Write-Host "[2/4] Memproses pencarian file PDF > $thresholdMB MB yang BELUM pernah dikompresi..." -ForegroundColor Cyan

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
# 4. FUNGSI UTAMA ADOBE COMPRESS JOB
# ==========================================
function Invoke-AdobeCompressSingle {
    param (
        [string]$SourcePath,
        [string]$TempPath,
        [hashtable]$Headers
    )

    try {
        # A. Minta Upload Asset URI
        $assetUri  = "https://pdf-services.adobe.io/assets"
        $assetBody = @{ mediaType = "application/pdf" } | ConvertTo-Json
        $assetResp = Invoke-RestMethod -Uri $assetUri -Method Post -Headers $Headers -Body $assetBody -ContentType "application/json" -ErrorAction Stop

        $uploadUrl = $assetResp.uploadUri
        $assetID   = $assetResp.assetID

        # B. Upload Binary File
        $fileBytes = [System.IO.File]::ReadAllBytes($SourcePath)
        Invoke-RestMethod -Uri $uploadUrl -Method Put -Body $fileBytes -ContentType "application/pdf" -ErrorAction Stop | Out-Null

        # C. Submit Compression Job (MEDIUM)
        $jobUri  = "https://pdf-services-ue1.adobe.io/operation/compresspdf"
        $jobBody = @{
            assetID          = $assetID
            compressionLevel = "MEDIUM"
        } | ConvertTo-Json

        $jobResp = Invoke-WebRequest -Uri $jobUri -Method Post -Headers $Headers -Body $jobBody -ContentType "application/json" -ErrorAction Stop
        
        $locationHeader = $jobResp.Headers["location"]
        $statusUrl = if ($locationHeader -is [array]) { $locationHeader[0] } else { $locationHeader }

        # D. Polling Status Job
        do {
            Start-Sleep -Seconds 2
            $statusResp = Invoke-RestMethod -Uri $statusUrl -Method Get -Headers $Headers -ErrorAction Stop
        } while ($statusResp.status -eq "in progress")

        # E. Download Output ke Temp Path
        if ($statusResp.status -eq "done") {
            $downloadUrl = $statusResp.asset.downloadUri
            Invoke-WebRequest -Uri $downloadUrl -OutFile $TempPath -ErrorAction Stop
            return $true
        } else {
            return $false
        }
    } catch {
        Write-Host "      !! API Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ==========================================
# 5. PROSES BERTAHAP PER FILE
# ==========================================
Write-Host "[3/4] Memulai proses kompresi via Adobe PDF Services...`n" -ForegroundColor Cyan

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

    # --- TAHAP A: DAFTARKAN KE CSV PENDING ---
    "$startTime;$($file.Name);`"$hierarchyDir`";$fileSizeMB;`"$sourcePath`"" | Out-File -LiteralPath $pendingLogFile -Append -Encoding UTF8

    # Indikator Progress Console
    $percent = [math]::Round(($currentIndex / $totalFiles) * 100, 0)
    Write-Progress -Activity "Adobe PDF Compression" -Status "[$currentIndex/$totalFiles] $($file.Name)" -PercentComplete $percent

    Write-Host "[$currentIndex/$totalFiles] " -NoNewline -ForegroundColor White
    Write-Host "Memproses: " -NoNewline -ForegroundColor Gray
    Write-Host "$($file.Name) " -NoNewline -ForegroundColor Magenta
    Write-Host "($fileSizeMB MB)" -ForegroundColor Yellow
    Write-Host "   -> Logged to Pending CSV" -ForegroundColor DarkGray
    
    # Deteksi Folder PT dan Path Backup Flat
    $relativeFromCorp = $sourcePath.Substring($corporateRoot.Length).TrimStart('\')
    $ptFolderName     = $relativeFromCorp.Split('\')[0]
    $ptBackupFolder   = Join-Path $backupCorporate $ptFolderName
    
    if (!(Test-Path -LiteralPath $ptBackupFolder)) {
        New-Item -ItemType Directory -Path $ptBackupFolder -Force | Out-Null
    }

    $backupPath = Join-Path $ptBackupFolder $file.Name

    # --- TAHAP B: PROSES COPY BACKUP KE FOLDER PT ---
    try {
        Copy-Item -LiteralPath $sourcePath -Destination $backupPath -Force -ErrorAction Stop
        Write-Host "   [1/2] Backup berhasil -> $backupPath" -ForegroundColor DarkGray
    } catch {
        Write-Host "   !! GAGAL BACKUP: ($($_.Exception.Message)) - Melewati Kompresi!" -ForegroundColor Red
        "$startTime;GAGAL_BACKUP;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        continue
    }

    # --- TAHAP C: PROSES KOMPRESI VIA ADOBE API KE FILE TEMP ---
    $tempPath = Join-Path $file.DirectoryName ("temp_adobe_compress_" + $file.Name)

    Write-Host "   [2/2] Mengirim job kompresi ke Adobe Cloud (MEDIUM)..." -ForegroundColor DarkGray
    $isSuccess = Invoke-AdobeCompressSingle -SourcePath $sourcePath -TempPath $tempPath -Headers $adobeHeaders

    # --- TAHAP D: VERIFIKASI, OVERWRITE & CATAT LOG ---
    if ($isSuccess -and (Test-Path -LiteralPath $tempPath) -and ((Get-Item -LiteralPath $tempPath).Length -gt 0)) {
        $newSizeMB = [math]::Round((Get-Item -LiteralPath $tempPath).Length / 1MB, 2)

        try {
            Remove-Item -LiteralPath $sourcePath -Force -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $sourcePath -Force -ErrorAction Stop
            
            Write-Host "   -> SUKSES Adobe Compress! Ditimpa ($fileSizeMB MB -> $newSizeMB MB)" -ForegroundColor Green
            
            # Log Sukses ke CSV Eksekusi
            "$startTime;SUKSES;$($file.Name);`"$hierarchyDir`";$fileSizeMB;$newSizeMB;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        } catch {
            Write-Host "   !! GAGAL OVERWRITE: File terkunci OneDrive/System ($($_.Exception.Message))" -ForegroundColor Red
            "$startTime;GAGAL_OVERWRITE;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
        }
    } else {
        Write-Host "   !! GAGAL KOMPRESI ADOBE!" -ForegroundColor Red
        "$startTime;GAGAL_COMPRESS;$($file.Name);`"$hierarchyDir`";$fileSizeMB;0;`"$backupPath`";`"$sourcePath`"" | Out-File -LiteralPath $executionLogFile -Append -Encoding UTF8
    }

    # Cleanup File Temp jika masih tersisa
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }

    Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
}

Write-Host "`n[4/4] [ SELESAI ] Seluruh batch Adobe Compression diproses." -ForegroundColor Yellow
Write-Host "Log Antrean PENDING : $pendingLogFile" -ForegroundColor White
Write-Host "Log Hasil EKSEKUSI  : $executionLogFile" -ForegroundColor White