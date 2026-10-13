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

# Path Output CSV Report Lockdown Massal
$csvReportFolder = $Config.ReportFolder
$csvReportPath   = "$csvReportFolder/Mass_GuestLockdown_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Inisialisasi Master Array untuk Laporan CSV
$lockdownReportList = @()

if (-not (Test-Path $csvReportFolder)) { New-Item -ItemType Directory -Path $csvReportFolder | Out-Null }

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES (MASSAL + FIX EQUAL FILTER)
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
        $cleanUrl -ne $Config.ExcludedSiteUrl -and # -> FIX: Menggunakan EQUAL murni agar cc lain aman
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
Write-Host "   MASS RUN: GUEST-ONLY LOCKDOWN ON LISTS " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site Target Eksekusi: $($targetSites.Count)" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# ==========================================
# 3. LOOPING MULTI-SITE & LOCKDOWN USER GUEST
# ==========================================
$listsToLock = @("IMMIGRATION DASHBOARD", "REGULATORY COMPLIANCE")

foreach ($targetUrl in $targetSites) {
    Write-Host "`nMengonfigurasi Hak Akses List di: $targetUrl" -ForegroundColor Cyan
    
    $siteLog = [PSCustomObject]@{
        SiteUrl           = $targetUrl
        LockdownStatus    = "Success"
        TotalGuestsLocked = 0
        Errors            = ""
    }

    try {
        Connect-PnPOnline -Url $targetUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword
        
        $ctx = Get-PnPContext
        $web = $ctx.Web
        
        # Ambil grup Members bawaan site beserta seluruh User di dalamnya
        $siteGroupMembers = $web.AssociatedMemberGroup
        $readRoleDef = $web.RoleDefinitions.GetByName("Read")
        
        $ctx.Load($siteGroupMembers)
        $ctx.Load($siteGroupMembers.Users)
        $ctx.Load($readRoleDef)
        $ctx.ExecuteQuery()

        foreach ($listName in $listsToLock) {
            try {
                $targetList = $web.Lists.GetByTitle($listName)
                $ctx.Load($targetList)
                $ctx.ExecuteQuery()
                
                if ($targetList) {
                    Write-Host "   -> Memproses List '$listName'..." -ForegroundColor Yellow
                    
                    # 1. Putus warisan izin list (Salin izin induk agar internal tetep bisa edit via grup)
                    $targetList.BreakRoleInheritance($true, $false)
                    $ctx.ExecuteQuery()
                    
                    # 2. Lakukan perulangan memeriksa setiap user di dalam grup Members
                    foreach ($user in $siteGroupMembers.Users) {
                        $isGuest = $false
                        if ($user.LoginName -like "*#ext#*" -or $user.UserId.NameId -like "*guest*") {
                            $isGuest = $true
                        }

                        if ($isGuest) {
                            Write-Host "      [*] Mendeteksi Guest: $($user.Title) ($($user.Email))" -ForegroundColor Magenta
                            
                            # 3. Hapus hak akses bawaan lama si Guest ini (jika ada individual assignment sebelumnya)
                            try {
                                $targetList.RoleAssignments.GetByPrincipal($user).DeleteObject()
                                $ctx.ExecuteQuery()
                            } catch {}
                            
                            # 4. Berikan hak akses Read murni khusus untuk user Guest ini secara individu di list tersebut
                            $roleBindings = [Microsoft.SharePoint.Client.RoleDefinitionBindingCollection]::new($ctx)
                            $roleBindings.Add($readRoleDef)
                            
                            $targetList.RoleAssignments.Add($user, $roleBindings)
                            $targetList.Update()
                            $ctx.ExecuteQuery()
                            
                            $siteLog.TotalGuestsLocked++
                            Write-Host "      [OK] Guest '$($user.Title)' berhasil dikunci menjadi Read-Only." -ForegroundColor Green
                        }
                    }
                }
            }
            catch {
                # Jika sub-proses satu list gagal/tidak ketemu, biarkan lanjut ke list berikutnya
                if ($_.Exception.Message -notlike "*Cannot find list*") {
                    Write-Host "      [WARN] Gagal memproses list $listName : $_" -ForegroundColor Red
                }
            }
        }
        
        Disconnect-PnPOnline
    }
    catch {
        Write-Host "   -> [ERROR] Gagal memproses site ini: $_" -ForegroundColor Red
        $siteLog.LockdownStatus = "Failed"
        $siteLog.Errors = $_.Exception.Message
        try { Disconnect-PnPOnline } catch {}
    }

    $lockdownReportList += $siteLog
}

# ==========================================
# 4. EXPORT DATA KE CSV REPORT
# ==========================================
Write-Host "`nMengekspor laporan akhir Lockdown ke CSV..." -ForegroundColor Magenta
$lockdownReportList | Export-Csv -Path $csvReportPath -NoTypeInformation -Delimiter "," -Encoding UTF8

Write-Host "==========================================" -ForegroundColor Green
Write-Host "    PROSES MASS LOCKDOWN GUEST SELESAI TOTAL! " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " File Report CSV tersimpan di:" -ForegroundColor Yellow
Write-Host " $csvReportPath" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Green