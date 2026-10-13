# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL & ADMIN (MENGGUNAKAN CERTIFICATE)
# ==========================================
$adminSiteUrl   = $Config.AdminUrl
$sourceSiteURL  = $Config.TemplatePortalUrl
$sourceListName = "IMMIGRATION DASHBOARD"
$targetListName = "IMMIGRATION DASHBOARD"

# Konfigurasi App Registration & Sertifikat Spesifik
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Konfigurasi Path Output CSV Report
$csvReportFolder = $Config.ReportFolder
$csvReportPath   = "$csvReportFolder/SyncReport_Immigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Inisialisasi Array untuk Laporan CSV
$syncReport = @()

# Ensure folder report ada, jika belum buat otomatis
if (-not (Test-Path $csvReportFolder)) { New-Item -ItemType Directory -Path $csvReportFolder | Out-Null }

# ==========================================
# 2. PROSES SEARCH & FILTERING TARGET SITES (MASSAL)
# ==========================================
Write-Host "Konek ke Admin Tenant menggunakan App Registration & Certificate..." -ForegroundColor Magenta
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$allSites = Get-PnPTenantSite

# Proses Filtering Otomatis sesuai rules parameter awal
$filteredSites = $allSites | Where-Object {
    $url = $_.Url
    $title = $_.Title
    (
        $url -like "*/sites/*" -and
        $url -notlike "*content*" -and
        $url -notlike "*app*" -and
        $url -notlike "*VivaHome*" -and
        $url -notlike "*$($Config.InternalSiteAlias)*" -and
        $url -notlike "*$($Config.InternalWikiAlias)*" -and
        $url -notlike "*portal*" -and
        $title -notlike "*$($Config.InternalSiteAlias)*" -and
        $title -notlike "*PORTAL*" -and
        $title -notlike "*FILL GAPS SITE*" -and
        $title -notlike "*Communication*" -and
        $title -notlike "*$($Config.InternalTitleKeyword)*" -and
        $title -notlike "*RedirectSite*"
    )
}

# Target site ditentukan penuh dari hasil filter tenant secara dinamis
$targetSites = $filteredSites.Url

# Tampilkan Summary Scan Tenant
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "         SHAREPOINT SITE SUMMARY          " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Total Site Sebelum Filter : $($allSites.Count)" -ForegroundColor Yellow
Write-Host " Total Site Lolos Filter   : $($filteredSites.Count)" -ForegroundColor Green
Write-Host " Total Site Ter-exclude    : $($allSites.Count - $filteredSites.Count)" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Cyan

Disconnect-PnPOnline

# ==========================================
# 3. PROSES EKSTRAKSI FORMAT & SCHEMA DARI SUMBER
# ==========================================
Write-Host "`nMenghubungkan ke Site Sumber..." -ForegroundColor Cyan
Connect-PnPOnline -Url $sourceSiteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

$sourceAllFields      = Get-PnPField -List $sourceListName
$sourceFormattedFields = $sourceAllFields | Where-Object { $_.CustomFormatter -ne $null -and $_.CustomFormatter -ne "" }
$sourceViews           = Get-PnPView -List $sourceListName

Disconnect-PnPOnline

# ==========================================
# 4. LOOPING MASS DEPLOYMENT & LOGGING REPORT
# ==========================================
Write-Host "`nMemulai eksekusi massal kloning identik ke $($targetSites.Count) site..." -ForegroundColor Magenta

foreach ($targetUrl in $targetSites) {
    Write-Host "`nMengonfigurasi list di: $targetUrl" -ForegroundColor Cyan
    
    # Template objek record untuk report CSV per site
    $siteLog = [PSCustomObject]@{
        SiteUrl          = $targetUrl
        Status           = "Success"
        Details          = "Kloning skema, rumus STATUS terbaru, view default All Items, dan pembersihan kolom asing berhasil."
        DeletedColumns   = ""
        CreatedColumns   = ""
        Errors           = ""
    }

    try {
        Connect-PnPOnline -Url $targetUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword
        
        # Proteksi awal: Cek apakah target list eksis
        $targetList = Get-PnPList -Identity $targetListName -ErrorAction SilentlyContinue
        if (-not $targetList) {
            Write-Host "   -> Skip: List '$targetListName' tidak ditemukan di site ini." -ForegroundColor Red
            $siteLog.Status = "Skipped"
            $siteLog.Details = "List '$targetListName' tidak ditemukan di site target ini."
            Disconnect-PnPOnline
            $syncReport += $siteLog
            continue
        }

        # -- A. Bersihkan Kolom Berbeda (Hapus Kolom Asing) --
        $targetAllFields = Get-PnPField -List $targetListName
        $systemFields = @("ID", "Title", "Created", "Author", "Modified", "Editor", "Attachments", "ContentType", "DocIcon", "FolderChildCount", "ItemChildCount", "ComplianceAssetId")
        $deletedList = @()

        foreach ($targetField in $targetAllFields) {
            if ($targetField.InternalName -notin $sourceAllFields.InternalName -and $targetField.InternalName -notin $systemFields -and -not $targetField.ReadOnlyField -and -not $targetField.Hidden) {
                Write-Host "   -> Kolom asing ditemukan: '$($targetField.InternalName)'. Menghapus..." -ForegroundColor Red
                Remove-PnPField -List $targetListName -Identity $targetField.InternalName -Force -ErrorAction SilentlyContinue | Out-Null
                $deletedList += $targetField.InternalName
            }
        }
        if ($deletedList.Count -gt 0) { $siteLog.DeletedColumns = ($deletedList -join "; ") }

        # -- B. Sinkronisasi Struktur Kolom --
        $targetAllFieldsLive = Get-PnPField -List $targetListName
        $targetInternalNames = @($targetAllFieldsLive | ForEach-Object { $_.InternalName })
        $createdList = @()

        $normalFields = $sourceAllFields | Where-Object { $_.TypeAsString -ne "Calculated" }
        $calculatedFields = $sourceAllFields | Where-Object { $_.TypeAsString -eq "Calculated" }
        $sortedSourceFields = $normalFields + $calculatedFields

        foreach ($sourceField in $sortedSourceFields) {
            
            # Jalur Khusus Kolom STATUS (Sinkronisasi Formula Dinamis)
            if ($sourceField.InternalName -eq "STATUS") {
                $sourceCalcField = [Microsoft.SharePoint.Client.FieldCalculated]$sourceField
                $sourceFormula   = $sourceCalcField.Formula

                if ("STATUS" -notin $targetInternalNames) {
                    Write-Host "   -> Kolom STATUS belum ada. Membuat baru dari schema sumber..." -ForegroundColor Yellow
                    $targetList.Fields.AddFieldAsXml($sourceField.SchemaXml, $false, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
                    $targetList.Context.ExecuteQuery()
                    $createdList += "STATUS"
                } 
                else {
                    $targetField     = $targetAllFieldsLive | Where-Object { $_.InternalName -eq "STATUS" }
                    $targetCalcField = [Microsoft.SharePoint.Client.FieldCalculated]$targetField

                    # Update formula jika terdeteksi ada perbedaan nilai di source
                    if ($targetCalcField.Formula -ne $sourceFormula) {
                        Write-Host "   -> [UPDATE RUMUS] Mengubah formula STATUS agar sinkron dengan sumber..." -ForegroundColor Yellow
                        $targetCalcField.Formula = $sourceFormula
                        $targetCalcField.Update()
                        $targetList.Context.ExecuteQuery()
                    }
                }
                
                Set-PnPField -List $targetListName -Identity "STATUS" -Values @{ Hidden = $sourceField.Hidden; Required = $sourceField.Required; Title = $sourceField.Title } -ErrorAction SilentlyContinue | Out-Null
                continue
            }

            # Jalur Generik Kolom Lainnya
            if ($sourceField.InternalName -in $targetInternalNames) {
                $matchedTarget = $targetAllFieldsLive | Where-Object { $_.InternalName -eq $sourceField.InternalName }
                $propUpdates = @{}
                if ($matchedTarget.Hidden -ne $sourceField.Hidden) { $propUpdates["Hidden"] = $sourceField.Hidden }
                if ($matchedTarget.Required -ne $sourceField.Required) { $propUpdates["Required"] = $sourceField.Required }
                if ($matchedTarget.Title -ne $sourceField.Title) { $propUpdates["Title"] = $sourceField.Title }

                if ($propUpdates.Count -gt 0) {
                    Set-PnPField -List $targetListName -Identity $sourceField.InternalName -Values $propUpdates -ErrorAction SilentlyContinue | Out-Null
                }
            }
            else {
                $targetList.Fields.AddFieldAsXml($sourceField.SchemaXml, $false, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
                $targetList.Context.ExecuteQuery()
                
                Set-PnPField -List $targetListName -Identity $sourceField.InternalName -Values @{ Hidden = $sourceField.Hidden; Required = $sourceField.Required; Title = $sourceField.Title } -ErrorAction SilentlyContinue | Out-Null
                $createdList += $sourceField.InternalName
            }
        }
        if ($createdList.Count -gt 0) { $siteLog.CreatedColumns = ($createdList -join "; ") }

        # -- C. Sinkronisasi JSON Column Formatting --
        foreach ($field in $sourceFormattedFields) {
            Set-PnPField -List $targetListName -Identity $field.InternalName -Values @{CustomFormatter = $field.CustomFormatter} -ErrorAction SilentlyContinue | Out-Null
        }
        
        # -- D. Sinkronisasi Layout View & Pembersihan "OPEN" --
        $checkOpenView = Get-PnPView -List $targetListName | Where-Object { $_.Title -eq "OPEN" }
        if ($checkOpenView) {
            Remove-PnPView -List $targetListName -Identity "OPEN" -Force -ErrorAction SilentlyContinue | Out-Null
        }

        foreach ($view in $sourceViews) {
            $targetView = Get-PnPView -List $targetListName | Where-Object { $_.Title -eq $view.Title }
            [string[]]$exactViewFields = @($view.ViewFields | ForEach-Object { $_ })

            $viewValues = @{ CustomFormatter = $view.CustomFormatter }
            if ($view.Title -eq "All Items") { $viewValues["DefaultView"] = $true }

            if ($targetView) {
                Set-PnPView -List $targetListName -Identity $view.Title -Fields $exactViewFields -Values $viewValues | Out-Null
            } else {
                Add-PnPView -List $targetListName -Title $view.Title -Fields $exactViewFields -Query $view.ViewQuery -ErrorAction SilentlyContinue | Out-Null
                Set-PnPView -List $targetListName -Identity $view.Title -Values $viewValues | Out-Null
            }
        }

        Write-Host "   -> [BERHASIL] Seluruh konfigurasi & rumus status telah sinkron." -ForegroundColor Green
        Disconnect-PnPOnline

    } catch {
        Write-Host "   -> [ERROR] Terjadi kegagalan: $_" -ForegroundColor Red
        $siteReportErrors = $_.Exception.Message
        
        $siteLog.Status = "Failed"
        $siteLog.Details = "Gagal memproses site akibat kesalahan internal system."
        $siteLog.Errors = $siteReportErrors
        
        try { Disconnect-PnPOnline } catch {}
    }

    $syncReport += $siteLog
}

# ==========================================
# 5. EXPORT MASTER REPORT KE CSV FILE
# ==========================================
Write-Host "`nEkspor hasil laporan sinkronisasi massal ke CSV..." -ForegroundColor Magenta
$syncReport | Export-Csv -Path $csvReportPath -NoTypeInformation -Delimiter "," -Encoding UTF8

Write-Host "==========================================" -ForegroundColor Green
Write-Host "       PROSES MASS-DEPLOYMENT SELESAI!     " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " File Report CSV tersimpan di:" -ForegroundColor Yellow
Write-Host " $csvReportPath" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Green