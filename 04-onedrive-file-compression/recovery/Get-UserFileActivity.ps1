# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI KREDENSIAL & PARAMETER
# ==========================================
$siteUrl        = $Config.SharePointRoot
$tenant         = $Config.TenantDomain
$clientId       = $Config.AppClientId
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Target User untuk di-filter
$targetUserEmail = $Config.ActivityUserUpn
$targetUserName  = $Config.ActivityUserDisplayName

$outputPath     = (Join-Path $Config.ReportFolder "Tenant_Activity_$(Get-Date -Format yyyyMMdd_HHmm).csv")

# ==========================================
# 2. KONEKSI KE SHAREPOINT MENGGUNAKAN SERTIFIKAT
# ==========================================
Write-Host "Menghubungkan ke SharePoint Tenant..." -ForegroundColor Cyan

Connect-PnPOnline -Url $siteUrl `
                  -ClientId $clientId `
                  -Tenant $tenant `
                  -CertificatePath $pfxPath `
                  -CertificatePassword $pfxPassword

Write-Host "Berhasil terhubung!" -ForegroundColor Green

# ==========================================
# 3. EXTRACTION TENANT-WIDE ACTIVITY (16:00 - 22:00 GMT+8)
# ==========================================
# 16:00 - 22:00 WITA (GMT+8) -> 08:00 - 14:00 UTC
$startUtcStr = "2026-08-06T08:00:00Z"
$endUtcStr   = "2026-08-06T14:00:00Z"

$activityResults = @()

Write-Host "Mengambil data Activity khusus '$targetUserEmail' (06 Aug 2026 jam 16:00 - 22:00 WITA)..." -ForegroundColor Yellow

# Query KQL dengan filter user spesifik dan rentang waktu 16:00 - 22:00 WITA
$searchQuery = "LastModifiedTime>=$startUtcStr AND LastModifiedTime<=$endUtcStr AND IsContainer=false AND (EditorOWSUSER:$targetUserEmail OR Author:$targetUserEmail)"

try {
    # Mengambil hasil dari Search API
    $searchResults = Submit-PnPSearchQuery -Query $searchQuery `
                                           -SelectProperties "Title", "Path", "Author", "EditorOWSUSER", "LastModifiedTime", "FileType", "SPSiteUrl" `
                                           -SortList @{ "LastModifiedTime" = "Descending" } `
                                           -MaxResults 500

    foreach ($row in $searchResults.ResultRows) {
        
        # Format nama pengedit dari properti SharePoint Search
        $editorRaw = $row["EditorOWSUSER"]
        $editorName = if ($editorRaw) { ($editorRaw -split '\|')[-1].Trim() } else { "Unknown" }

        # Verifikasi agar hanya memasukkan data milik user target
        if ($editorRaw -like "*$targetUserEmail*" -or $editorName -like "*$targetUserName*" -or $row["Author"] -like "*$targetUserEmail*") {
            
            # Konversi tampilan waktu kembali ke WITA (GMT+8)
            $utcTime = [DateTime]$row["LastModifiedTime"]
            $localTime = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($utcTime, "Singapore Standard Time")

            $activityResults += [PSCustomObject]@{
                "Nama File / Dokumen" = $row["Title"]
                "Jenis File"          = $row["FileType"]
                "Diubah Oleh"         = $editorName
                "Waktu Modifikasi"    = $localTime.ToString("dd/MM/yyyy HH:mm:ss")
                "Site SharePoint"     = $row["SPSiteUrl"]
                "Link File / Path"    = $row["Path"]
            }
        }
    }

    # ==========================================
    # 4. SIMPAN KE CSV
    # ==========================================
    if ($activityResults.Count -gt 0) {
        # Buat folder jika belum ada
        $folderPath = Split-Path $outputPath
        if (-not (Test-Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
        }

        $activityResults | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n[SUKSES] Berhasil mengambil $($activityResults.Count) item Activity milik $targetUserEmail antara jam 16:00 - 22:00 WITA!" -ForegroundColor Green
        Write-Host "File disimpan di: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "`n[INFO] Tidak ada data aktivitas file ditemukan untuk $targetUserEmail pada jam 16:00 - 22:00 WITA." -ForegroundColor Yellow
    }

} catch {
    Write-Error "Terjadi kesalahan saat mengambil Activity Tenant: $_"
}

Disconnect-PnPOnline