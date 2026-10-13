# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL (LIVE UPDATE)
# ==========================================
$siteURL        = $Config.HubSiteUrl
$listName       = "DATABASE KLIEN"
$kolomNama      = "field_1"

# Konfigurasi Akses (Certificate)
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Mode update:
#   "Update"       -> Modified & Modified By berubah, versi baru dibuat
#   "SystemUpdate" -> Modified & Modified By TIDAK berubah, versi baru tidak dibuat
$updateType     = "Update"

# Backup nilai lama (untuk rollback)
$backupPath     = "$env:USERPROFILE\Desktop\BACKUP_PT_TITIK_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# ==========================================
# 2. TARIK DATA & SIAPKAN KANDIDAT
# ==========================================
Write-Host "Menghubungkan ke Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Menarik semua item dari '$listName'..." -ForegroundColor Cyan
$allItems = Get-PnPListItem -List $listName -Fields "ID", $kolomNama -PageSize 2000

$kandidat = foreach ($item in $allItems) {
    $nama = [string]$item[$kolomNama]
    # Hanya yang diawali "PT." (case-insensitive, boleh ada spasi di depan)
    if ($nama -and $nama -match '^\s*[Pp][Tt]\.') {
        $baru = ($nama -replace '^\s*[Pp][Tt]\.\s*', 'PT ').Trim()
        if ($baru -ne $nama) {
            [PSCustomObject]@{
                ID        = $item.Id
                NilaiLama = $nama
                NilaiBaru = $baru
            }
        }
    }
}

if (-not $kandidat) {
    Write-Host "[INFO] Tidak ada item yang perlu diupdate." -ForegroundColor Yellow
    Disconnect-PnPOnline; exit
}

$kandidat = @($kandidat)
$totalItems = $kandidat.Count

# Backup dulu sebelum apa pun diubah
$kandidat | Export-Csv -Path $backupPath -NoTypeInformation -Encoding UTF8
Write-Host "Backup nilai lama disimpan: $backupPath" -ForegroundColor Cyan

# ==========================================
# 3. KONFIRMASI
# ==========================================
Write-Host "`nContoh perubahan (5 pertama):" -ForegroundColor White
$kandidat | Select-Object -First 5 | Format-Table -AutoSize

Write-Host "Total $totalItems item akan diubah (mode: $updateType)." -ForegroundColor Magenta
$konfirmasi = Read-Host "Ketik YA untuk melanjutkan"
if ($konfirmasi -ne "YA") {
    Write-Host "[BATAL] Tidak ada perubahan dilakukan." -ForegroundColor Yellow
    Disconnect-PnPOnline; exit
}

# ==========================================
# 4. PROSES UPDATE
# ==========================================
$counter = 0
$successCount = 0
$gagal = @()

foreach ($row in $kandidat) {
    $counter++
    try {
        Set-PnPListItem -List $listName -Identity $row.ID -Values @{ $kolomNama = $row.NilaiBaru } -UpdateType $updateType -ErrorAction Stop | Out-Null
        $successCount++
        Write-Host "   [$counter/$totalItems] [OK] ID $($row.ID): '$($row.NilaiLama)' -> '$($row.NilaiBaru)'" -ForegroundColor Green
    }
    catch {
        $gagal += [PSCustomObject]@{ ID = $row.ID; NilaiLama = $row.NilaiLama; Error = $_.Exception.Message }
        Write-Host "   [$counter/$totalItems] [FAIL] ID $($row.ID): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ==========================================
# 5. LAPORAN
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "        PROSES UPDATE DATA SELESAI!       " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " Total Kandidat   : $totalItems" -ForegroundColor White
Write-Host " Sukses Terupdate : $successCount" -ForegroundColor Green
Write-Host " Gagal            : $($gagal.Count)" -ForegroundColor Yellow
Write-Host " Backup           : $backupPath" -ForegroundColor Cyan

if ($gagal.Count -gt 0) {
    $gagalPath = $backupPath -replace 'BACKUP_', 'GAGAL_'
    $gagal | Export-Csv -Path $gagalPath -NoTypeInformation -Encoding UTF8
    Write-Host " Daftar gagal     : $gagalPath" -ForegroundColor Red
}
Write-Host "==========================================" -ForegroundColor Green

Disconnect-PnPOnline