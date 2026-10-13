# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# KONFIGURASI TARGET LIST
# ==========================================
$siteURL        = $Config.HubSiteUrl 
$listName       = "DATABASE WORKFLOWS" # <--- Ganti dengan nama list yang mau dicek
$csvOutputPath  = (Join-Path $Config.ReportFolder "Daftar_Internal_Name_Kolom.csv")

# Konfigurasi Akses (Menggunakan Sertifikatmu)
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# ==========================================
# PROSES PENARIKAN DATA SKEMA KOLOM
# ==========================================
Write-Host "Menghubungkan ke Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Membaca struktur kolom dari list '$listName'..." -ForegroundColor Cyan
$allFields = Get-PnPField -List $listName -ErrorAction SilentlyContinue

if (-not $allFields) {
    Write-Host "[ERROR] List tidak ditemukan atau tidak ada akses!" -ForegroundColor Red
    Disconnect-PnPOnline; exit
}

# Susun hasil data kolom ke bentuk tabel
$reportColumns = @()
foreach ($field in $allFields) {
    # Filter agar kolom bawaan sistem yang tidak penting (hidden) tidak mengotori CSV
    if ($field.Hidden -eq $false -or $field.InternalName -like "field_*" -or $field.InternalName -like "TAKA*") {
        $reportColumns += [PSCustomObject]@{
            "Nama di Layar (Title)" = $field.Title
            "Nama Asli (Internal)"  = $field.InternalName
            "Tipe Kolom (Type)"     = $field.TypeAsString
        }
    }
}

# ==========================================
# EKSPOR KE CSV
# ==========================================
Write-Host "Mengekspor daftar internal name ke CSV..." -ForegroundColor Yellow
$reportColumns | Export-Csv -Path $csvOutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

Disconnect-PnPOnline
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "[SUKSES] Daftar Internal Name selesai dibuat! Silakan buka file di:" -ForegroundColor Green
Write-Host "$csvOutputPath" -ForegroundColor Yellow
Write-Host "==========================================================================" -ForegroundColor Cyan