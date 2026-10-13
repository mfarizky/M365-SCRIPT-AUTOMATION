# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL & ADMIN (MENGGUNAKAN CERTIFICATE)
# ==========================================
$adminSiteUrl   = $Config.AdminUrl

# Konfigurasi App Registration & Sertifikat Spesifik
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword
$ExportPath     = Join-Path $Config.ReportFolder "Owner and Admin_Klien Portal.csv"

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES
# ==========================================
Write-Host "Konek ke Admin Tenant menggunakan App Registration & Certificate..." -ForegroundColor Magenta
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$allSites = Get-PnPTenantSite

# 3. Proses Filtering Otomatis
$filteredSites = $allSites | Where-Object {
    $url = $_.Url
    $title = $_.Title
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
        $title -notlike "*RedirectSite*"
    )
}
$targetSites = $filteredSites.Url

# Tampilkan Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "         SHAREPOINT SITE SUMMARY          " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site Sebelum Filter : $($allSites.Count)" -ForegroundColor Yellow
Write-Host " Total Site Lolos Filter   : $($filteredSites.Count)" -ForegroundColor Green
Write-Host " Total Site Ter-exclude    : $($allSites.Count - $filteredSites.Count)" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "Dari $($allSites.Count) site, ditemukan $($targetSites.Count) site yang cocok dengan kriteria filter." -ForegroundColor Green

# 4. Mulai Audit Data Admin & Owner
$Report = @()
$counter = 1

foreach ($SiteUrl in $targetSites) {
    $SiteInfo = $filteredSites | Where-Object { $_.Url -eq $SiteUrl }
    Write-Host "[$counter/$($targetSites.Count)] Mengaudit: $SiteUrl" -ForegroundColor Gray
    
    try {
        # === PERBAIKAN DI SINI: Menyesuaikan koneksi menggunakan Sertifikat ===
        $SiteConnection = Connect-PnPOnline -Url $SiteUrl `
                                            -ClientId $clientId `
                                            -Tenant $tenant `
                                            -CertificatePath $pfxPath `
                                            -CertificatePassword $pfxPassword `
                                            -ReturnConnection
        
        # A. Mengambil Site Collection Administrators (Primary & Secondary Admins)
        $SiteAdmins = Get-PnPSiteCollectionAdmin -Connection $SiteConnection | Select-Object -ExpandProperty Email
        $AdminString = $SiteAdmins -join "; "

        # B. Mengambil User dari Grup Default 'Owners'
        $OwnersGroup = Get-PnPGroup -Connection $SiteConnection | Where-Object { $_.Title -like "*Owners*" }
        $GroupOwners = @()
        if ($OwnersGroup) {
            $GroupOwners = Get-PnPGroupMember -Identity $OwnersGroup.Title -Connection $SiteConnection | Select-Object -ExpandProperty Email
        }
        $OwnerString = $GroupOwners -join "; "

        # C. Masukkan ke dalam Report
        $Report += [PSCustomObject]@{
            "Site Title"      = $SiteInfo.Title
            "Site URL"        = $SiteUrl
            "Site Admins"     = $AdminString
            "Group Owners"    = $OwnerString
            "Template"        = $SiteInfo.Template
        }
    } catch {
        Write-Host "Gagal mengambil data untuk $SiteUrl. Error: $_" -ForegroundColor Red
        $Report += [PSCustomObject]@{
            "Site Title"      = $SiteInfo.Title
            "Site URL"        = $SiteUrl
            "Site Admins"     = "ERROR ACCESSING SITE"
            "Group Owners"    = "ERROR ACCESSING SITE"
            "Template"        = $SiteInfo.Template
        }
    }
    $counter++
}

# 5. Ekspor Hasil ke CSV
if ($Report.Count -gt 0) {
    $Folder = Split-Path $ExportPath
    if (!(Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder | Out-Null }
    
    $Report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n[SUKSES] Hasil audit telah disimpan di: $ExportPath" -ForegroundColor Green
} else {
    Write-Host "`n[PERINGATAN] Tidak ada data yang ditemukan atau diekspor." -ForegroundColor Yellow
}