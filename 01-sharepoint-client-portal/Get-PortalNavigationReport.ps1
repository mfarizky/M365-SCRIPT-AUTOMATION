# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL & ADMIN (MENGGUNAKAN CERTIFICATE)
# ==========================================
$adminSiteUrl   = $Config.AdminUrl
$tenant         = $Config.TenantDomain
$clientId       = $Config.AppClientId
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Path Output CSV Report Navigation
$csvReportFolder = $Config.ReportFolder
$csvReportPath   = "$csvReportFolder/Navigation_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Inisialisasi Master Array untuk CSV
$navigationMasterList = @()

if (-not (Test-Path $csvReportFolder)) { New-Item -ItemType Directory -Path $csvReportFolder | Out-Null }

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES
# ==========================================
Write-Host "Konek ke Admin Tenant untuk Scan Daftar Site..." -ForegroundColor Magenta
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$allSites = Get-PnPTenantSite
$filteredSites = $allSites | Where-Object {
    $url = $_.Url; $title = $_.Title
    ($url -like "*/sites/*" -and $url -notlike "*content*" -and $url -notlike "*app*" -and $url -notlike "*VivaHome*" -and $url -notlike "*$($Config.InternalSiteAlias)*" -and $url -notlike "*$($Config.InternalWikiAlias)*" -and $url -notlike "*portal*" -and $title -notlike "*$($Config.InternalSiteAlias)*" -and $title -notlike "*PORTAL*" -and $title -notlike "*FILL GAPS SITE*" -and $title -notlike "*Communication*" -and $title -notlike "*$($Config.InternalTitleKeyword)*" -and $title -notlike "*RedirectSite*")
}
$targetSites = $filteredSites.Url
Disconnect-PnPOnline

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     RUNNING: NAVIGATION LINK AUDITOR     " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site yang Dihitung: $($targetSites.Count)" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# ==========================================
# 3. LOOPING SITES & EKSTRAKSI NAVIGASI (FIXED)
# ==========================================
foreach ($targetUrl in $targetSites) {
    Write-Host "`nMemeriksa Navigasi di: $targetUrl" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $targetUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword
        
        # Ambil data Menu Kiri (Quick Launch) dan Menu Atas (Top Navigation)
        $quickLaunchNodes = Get-PnPNavigationNode -Location QuickLaunch -ErrorAction SilentlyContinue
        $topNavNodes      = Get-PnPNavigationNode -Location TopNavigationBar -ErrorAction SilentlyContinue
        
        $siteItemCount = 0

        # --- A. PROSES MENU KIRI (QUICK LAUNCH) ---
        if ($quickLaunchNodes) {
            foreach ($node in $quickLaunchNodes) {
                $siteItemCount++
                $navigationMasterList += [PSCustomObject]@{
                    SiteUrl       = $targetUrl
                    MenuLocation  = "QuickLaunch (Menu Kiri)"
                    MenuTitle     = $node.Title
                    MenuUrl       = $node.Url
                    Hierarchy     = "Main Header"
                    ParentMenu    = "-"
                }

                # FIX: Inisialisasi koleksi Children secara eksplisit sebelum dibaca loop
                $null = Get-PnPProperty -ClientObject $node -Property Children -ErrorAction SilentlyContinue

                # Cek jika ada Sub-Menu di bawahnya (Child Nodes)
                if ($node.Children) {
                    foreach ($child in $node.Children) {
                        $siteItemCount++
                        $navigationMasterList += [PSCustomObject]@{
                            SiteUrl       = $targetUrl
                            MenuLocation  = "QuickLaunch (Menu Kiri)"
                            MenuTitle     = $child.Title
                            MenuUrl       = $child.Url
                            Hierarchy     = "Sub Menu (Child)"
                            ParentMenu    = $node.Title
                        }
                    }
                }
            }
        }

        # --- B. PROSES MENU ATAS (TOP NAVIGATION BAR) ---
        if ($topNavNodes) {
            foreach ($node in $topNavNodes) {
                $siteItemCount++
                $navigationMasterList += [PSCustomObject]@{
                    SiteUrl       = $targetUrl
                    MenuLocation  = "TopNavigationBar (Menu Atas)"
                    MenuTitle     = $node.Title
                    MenuUrl       = $node.Url
                    Hierarchy     = "Main Header"
                    ParentMenu    = "-"
                }

                # FIX: Inisialisasi koleksi Children secara eksplisit sebelum dibaca loop
                $null = Get-PnPProperty -ClientObject $node -Property Children -ErrorAction SilentlyContinue

                if ($node.Children) {
                    foreach ($child in $node.Children) {
                        $siteItemCount++
                        $navigationMasterList += [PSCustomObject]@{
                            SiteUrl       = $targetUrl
                            MenuLocation  = "TopNavigationBar (Menu Atas)"
                            MenuTitle     = $child.Title
                            MenuUrl       = $child.Url
                            Hierarchy     = "Sub Menu (Child)"
                            ParentMenu    = $node.Title
                        }
                    }
                }
            }
        }

        Write-Host "   -> Sukses menarik $siteItemCount item navigasi." -ForegroundColor Green
        Disconnect-PnPOnline
    }
    catch {
        Write-Host "   -> [ERROR] Gagal membaca navigasi di site ini: $_" -ForegroundColor Red
        try { Disconnect-PnPOnline } catch {}
    }
}

# ==========================================
# 4. EXPORT DATA KE CSV REPORT
# ==========================================
Write-Host "`nMengekspor data inventory navigasi ke CSV..." -ForegroundColor Magenta
$navigationMasterList | Export-Csv -Path $csvReportPath -NoTypeInformation -Delimiter "," -Encoding UTF8

Write-Host "==========================================" -ForegroundColor Green
Write-Host "       AUDIT NAVIGASI MASSAL SELESAI!      " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " Total records berhasil didata : $($navigationMasterList.Count) item" -ForegroundColor Yellow
Write-Host " File Laporan CSV tersimpan di :" -ForegroundColor White
Write-Host " $csvReportPath" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Green