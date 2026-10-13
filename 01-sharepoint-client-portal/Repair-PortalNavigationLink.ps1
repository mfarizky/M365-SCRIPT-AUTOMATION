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

# Path Output CSV Report Navigation Massal
$csvReportFolder = $Config.ReportFolder
$csvReportPath   = "$csvReportFolder/Fix_Portal_Only_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Inisialisasi Master Array untuk Laporan CSV
$navReportList = @()

if (-not (Test-Path $csvReportFolder)) { New-Item -ItemType Directory -Path $csvReportFolder | Out-Null }

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES
# ==========================================
Write-Host "Konek ke Admin Tenant untuk Scan Daftar Site..." -ForegroundColor Magenta
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$allSites = Get-PnPTenantSite
$filteredSites = $allSites | Where-Object {
    $url = $_.Url; $title = $_.Title
    $cleanUrl = $url.TrimEnd('/')
    
    (
        $url -like "*/sites/*" -and 
        $url -notlike "*content*" -and 
        $url -notlike "*app*" -and 
        $url -notlike "*VivaHome*" -and 
        $url -notlike "*$($Config.InternalSiteAlias)*" -and 
        $url -notlike "*$($Config.InternalWikiAlias)*" -and 
        $url -notlike "*portal*" -and 
        $cleanUrl -ne $Config.ExcludedSiteUrl -and 
        $title -notlike "*$($Config.InternalSiteAlias)*" -and 
        $title -notlike "*PORTAL*" -and 
        $title -notlike "*FILL GAPS SITE*" -and 
        $title -notlike "*Communication*" -and 
        $title -notlike "*$($Config.InternalTitleKeyword)*" -and 
        $title -notlike "*RedirectSite*"
    )
}
$targetSites = $filteredSites.Url
Disconnect-PnPOnline

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "         MASS RUN: FIX PORTAL LINK ONLY    " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site Target Eksekusi: $($targetSites.Count)" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# ==========================================
# 3. LOOPING MULTI-SITE & UPDATE MENU PORTAL SAJA
# ==========================================
foreach ($targetUrl in $targetSites) {
    Write-Host "`nMemeriksa Menu PORTAL di: $targetUrl" -ForegroundColor Cyan
    
    $siteLog = [PSCustomObject]@{
        SiteUrl      = $targetUrl
        PortalStatus = "No Change"
        Errors       = ""
    }

    try {
        Connect-PnPOnline -Url $targetUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword
        
        # Cari node navigasi yang namanya "PORTAL" di QuickLaunch
        $portalNode = Get-PnPNavigationNode -Location QuickLaunch | Where-Object { $_.Title -eq "PORTAL" }

        $targetPortalUrl = "$targetUrl/PORTAL/Forms/AllItems.aspx"

        if ($portalNode) {
            # Jika ketemu dan URL-nya masih salah (mengarah ke Shared Documents), langsung timpa URL-nya
            if ($portalNode.Url -notlike "*/PORTAL/Forms/*") {
                $portalNode.Url = $targetPortalUrl
                $portalNode.Update()
                Invoke-PnPQuery
                Write-Host "   -> [OK] Link PORTAL berhasil diperbarui ke library /PORTAL/." -ForegroundColor Green
                $siteLog.PortalStatus = "Updated"
            } else {
                Write-Host "   -> [SKIP] Link PORTAL sudah benar." -ForegroundColor Yellow
                $siteLog.PortalStatus = "Already Correct"
            }
        } else {
            # Jika menu PORTAL ternyata tidak ada sama sekali, kita buatkan baru supaya aman
            $null = Add-PnPNavigationNode -Title "PORTAL" -Url $targetPortalUrl -Location QuickLaunch -ErrorAction Stop
            Write-Host "   -> [OK] Menu PORTAL tidak ditemukan, berhasil ditambahkan baru." -ForegroundColor Green
            $siteLog.PortalStatus = "Created New"
        }

        Disconnect-PnPOnline
    }
    catch {
        Write-Host "   -> [ERROR] Gagal memproses menu PORTAL di site ini: $_" -ForegroundColor Red
        $siteLog.PortalStatus = "Failed"
        $siteLog.Errors = $_.Exception.Message
        try { Disconnect-PnPOnline } catch {}
    }

    $navReportList += $siteLog
}

# ==========================================
# 4. EXPORT DATA KE CSV REPORT
# ==========================================
Write-Host "`nMengekspor laporan perbaikan ke CSV..." -ForegroundColor Magenta
$navReportList | Export-Csv -Path $csvReportPath -NoTypeInformation -Delimiter "," -Encoding UTF8

Write-Host "==========================================" -ForegroundColor Green
Write-Host "        PERBAIKAN MENU PORTAL SELESAI!     " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " File Report CSV tersimpan di:" -ForegroundColor Yellow
Write-Host " $csvReportPath" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Green