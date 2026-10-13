# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL & ADMIN (MENGGUNAKAN CERTIFICATE)
# ==========================================
$adminSiteUrl   = $Config.AdminUrl
$sourceSiteURL  = $Config.TemplatePortalUrl
$sourceListName = "REGULATORY COMPLIANCE"
$targetListName = "REGULATORY COMPLIANCE"

# Konfigurasi App Registration & Sertifikat Spesifik
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES
# ==========================================
Write-Host "Konek ke Admin Tenant menggunakan App Registration & Certificate..." -ForegroundColor Magenta
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$allSites = Get-PnPTenantSite

# Proses Filtering Otomatis
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

Disconnect-PnPOnline

# ==========================================
# 3. PROSES EKSTRAKSI FORMAT DARI SUMBER
# ==========================================
Write-Host "`nMenghubungkan ke Site Sumber..." -ForegroundColor Cyan
Connect-PnPOnline -Url $sourceSiteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$sourceFields = Get-PnPField -List $sourceListName | Where-Object { $_.CustomFormatter -ne $null -and $_.CustomFormatter -ne "" }
$sourceViews  = Get-PnPView -List $sourceListName | Where-Object { $_.CustomFormatter -ne $null -and $_.CustomFormatter -ne "" }

Disconnect-PnPOnline

# ==========================================
# 4. LOOPING TARGET SITES (OTOMATIS VIA CERTIFICATE)
# ==========================================
Write-Host "`nMemulai proses injeksi format ke $($targetSites.Count) target site..." -ForegroundColor Magenta

foreach ($targetUrl in $targetSites) {
    Write-Host "`nMengupdate format di: $targetUrl" -ForegroundColor Cyan
    
    # Hubungkan ke target site menggunakan Certificate
    Connect-PnPOnline -Url $targetUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword
    
    # Proteksi: Cek apakah list target ada di site ini
    $targetList = Get-PnPList -Identity $targetListName -ErrorAction SilentlyContinue
    if (-not $targetList) {
        Write-Host "   -> Skip: List '$targetListName' tidak ditemukan di site ini." -ForegroundColor Red
        Disconnect-PnPOnline
        continue
    }

    # -- A. Update Column Formatting --
    foreach ($field in $sourceFields) {
        $targetField = Get-PnPField -List $targetListName | Where-Object { $_.InternalName -eq $field.InternalName }
        if ($targetField) {
            Set-PnPField -List $targetListName -Identity $field.InternalName -Values @{CustomFormatter = $field.CustomFormatter}
            Write-Host "   -> Kolom '$($field.InternalName)' berhasil di-format." -ForegroundColor Green
        }
    }
    
    # -- B. Update View Formatting --
    foreach ($view in $sourceViews) {
        $targetView = Get-PnPView -List $targetListName | Where-Object { $_.Title -eq $view.Title }
        
        if ($targetView) {
            # --- VIEW SUDAH ADA ---
            $filteredFields = $targetView.ViewFields | Where-Object { $_ -ne "Created" -and $_ -ne "Attachments" }
            
            $newTargetFields = @()
            foreach ($f in $filteredFields) {
                $newTargetFields += $f
                if ($f -eq "STATUS") { 
                    $newTargetFields += "Attachments" 
                }
            }
            if ($newTargetFields -notcontains "Attachments") { $newTargetFields += "Attachments" }
            
            Set-PnPView -List $targetListName -Identity $view.Title -Fields $newTargetFields -Values @{CustomFormatter = $view.CustomFormatter}
            Write-Host "   -> View '$($view.Title)' diperbarui. Posisi 'Attachments' setelah 'STATUS'." -ForegroundColor Green
            
        } else {
            # --- VIEW BELUM ADA (BUAT BARU) ---
            $filteredFields = $view.ViewFields | Where-Object { $_ -ne "Created" -and $_ -ne "Attachments" }
            
            $newSourceFields = @()
            foreach ($f in $filteredFields) {
                $newSourceFields += $f
                if ($f -eq "STATUS") { 
                    $newSourceFields += "Attachments" 
                }
            }
            if ($newSourceFields -notcontains "Attachments") { $newSourceFields += "Attachments" }

            Add-PnPView -List $targetListName -Title $view.Title -Fields $newSourceFields -Query $view.ViewQuery -Values @{CustomFormatter = $view.CustomFormatter}
            Write-Host "   -> View Baru '$($view.Title)' berhasil dibuat dengan posisi Attachments disesuaikan." -ForegroundColor Yellow
        }
    }
    
    # Putuskan koneksi sebelum bergeser ke site berikutnya
    Disconnect-PnPOnline
}

Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "  PROSES DUPLIKASI MASSAL SELESAI LUAR BIASA!  " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green