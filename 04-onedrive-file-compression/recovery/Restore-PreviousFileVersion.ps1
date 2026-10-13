# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. CREDENTIALS & PATH CONFIGURATION
# ==========================================
$ClientId    = $Config.AppClientId
$Tenant      = $Config.TenantDomain
$PfxPath     = $Config.PfxPath
$PfxPassword = $Config.PfxPassword

$CsvPath     = Join-Path $Config.ReportFolder "Tenant_Activity.csv"   # output dari Get-UserFileActivity.ps1
$LogPath     = Join-Path $Config.ReportFolder "Log_Restore_Results.csv"

# ==========================================
# 2. BACA FILE CSV (DELIMITER TITIK KOMA ';')
# ==========================================
if (-not (Test-Path $CsvPath)) {
    Write-Host "[ERROR] File CSV tidak ditemukan di $CsvPath" -ForegroundColor Red
    exit
}

Write-Host "Membaca file CSV ($CsvPath)..." -ForegroundColor Cyan
$CsvData = Import-Csv -Path $CsvPath -Delimiter ';'
Write-Host "Ditemukan $($CsvData.Count) baris data.`n" -ForegroundColor Yellow

$LogResults = @()
$CurrentConnectedSite = ""

# ==========================================
# 3. PROSES AUTO-RESTORE MULTI-SITE
# ==========================================
foreach ($Row in $CsvData) {
    $Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $OriginalUrl = $Row.'Link File / Path'
    $NamaDokumen = $Row.'Nama File / Dokumen'
    $DiubahOleh  = $Row.'Diubah Oleh'
    
    $Status = ""
    $Message = ""
    $RestoredVersionLabel = ""
    $RelativeUrl = ""
    $TargetSiteUrl = ""

    if ([string]::IsNullOrWhiteSpace($OriginalUrl)) { continue }

    try {
        # 1. Parsing URI & URL Decoding (menangani spasi %20, dll)
        $Uri = [System.Uri]$OriginalUrl
        $RelativeUrl = [System.Web.HttpUtility]::UrlDecode($Uri.AbsolutePath)

        # 2. Ekstrak Site URL secara dinamis (contoh: $($Config.SharePointRoot)/sites/services)
        $PathSegments = $Uri.Segments
        if ($PathSegments.Count -ge 3 -and $PathSegments[1].Trim('/') -eq 'sites') {
            $TargetSiteUrl = "$($Uri.Scheme)://$($Uri.Host)/sites/$($PathSegments[2].Trim('/'))"
        } else {
            $TargetSiteUrl = "$($Uri.Scheme)://$($Uri.Host)"
        }

        # 3. Connect to SharePoint Online Site (Ganti koneksi HANYA jika site berubah)
        if ($CurrentConnectedSite -ne $TargetSiteUrl) {
            Write-Host "`n[KONEKSI SITE] Connect-PnPOnline ke: $TargetSiteUrl" -ForegroundColor Cyan
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            
            Connect-PnPOnline -Url $TargetSiteUrl `
                              -ClientId $ClientId `
                              -Tenant $Tenant `
                              -CertificatePath $PfxPath `
                              -CertificatePassword $PfxPassword `
                              -ErrorAction Stop

            $CurrentConnectedSite = $TargetSiteUrl
            Write-Host "Koneksi Berhasil!`n" -ForegroundColor Green
        }

        Write-Host "Memproses File: $NamaDokumen" -ForegroundColor Yellow

        # 4. Ambil daftar versi untuk mendapatkan Version Label N-1
        $Versions = Get-PnPFileVersion -Url $RelativeUrl -ErrorAction Stop

        if ($Versions -and $Versions.Count -gt 0) {
            # Ambil String VersionLabel N-1 (misal: "1.0", "2.0")
            $PreviousVersionObj = $Versions[$Versions.Count - 1]
            $RestoredVersionLabel = [string]$PreviousVersionObj.VersionLabel

            # 5. Restore previous version with version Label (Sesuai contoh Anda)
            Restore-PnPFileVersion -Url $RelativeUrl -Identity $RestoredVersionLabel -Force -ErrorAction Stop
            
            $Status = "BERHASIL"
            $Message = "Berhasil restore versi $RestoredVersionLabel ke SharePoint"
            Write-Host "  [BERHASIL] File '$NamaDokumen' di-restore ke versi $RestoredVersionLabel" -ForegroundColor Green

        } else {
            $Status = "SKIPPED"
            $Message = "File hanya memiliki 1 versi (tidak ada versi sebelumnya)"
            Write-Host "  [SKIP] $Message" -ForegroundColor DarkYellow
        }

    } catch {
        $Status = "GAGAL"
        $Message = $_.Exception.Message
        Write-Host "  [ERROR] Gagal memproses: $Message" -ForegroundColor Red
    }

    # Catat log
    $LogResults += [PSCustomObject]@{
        Timestamp          = $Timestamp
        NamaDokumen        = $NamaDokumen
        DiubahOleh         = $DiubahOleh
        SiteUrl            = $TargetSiteUrl
        OriginalUrl        = $OriginalUrl
        ParsedRelativePath = $RelativeUrl
        Status             = $Status
        RestoredVersion    = $RestoredVersionLabel
        DetailMessage      = $Message
    }
}

# Disconnect koneksi terakhir
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ==========================================
# 4. EXPORT LOG HASIL KE CSV
# ==========================================
$LogResults | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "PROSES AUTO-RESTORE SELESAI!" -ForegroundColor Cyan
Write-Host "Total Diproses : $($LogResults.Count)" -ForegroundColor White
Write-Host "Berhasil       : $(($LogResults | Where-Object Status -eq 'BERHASIL').Count)" -ForegroundColor Green
Write-Host "Gagal          : $(($LogResults | Where-Object Status -eq 'GAGAL').Count)" -ForegroundColor Red
Write-Host "Skipped        : $(($LogResults | Where-Object Status -eq 'SKIPPED').Count)" -ForegroundColor Yellow
Write-Host "File Log CSV   : $LogPath" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan