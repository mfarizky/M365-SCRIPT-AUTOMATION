# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL (MENGGUNAKAN CERTIFICATE)
# ==========================================
$siteURL        = $Config.HubSiteUrl 
$listName       = "REGULATORY COMPLIANCE"

# Konfigurasi Akses
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Definisi internal name kolom SharePoint sesuai input user
$kolomFiling    = "COMPLIANCE"    # Kolom Kriteria (LKPM REPORT)
$kolomIDRC      = "IDKLIENRC"     # Kolom ID Klien RC
$kolomCLIENT    = "KLIEN"         # Kolom Klien untuk log display
$kolomDeadline  = "DEADLINEDATE"   # Kolom Target Deadline Date

# ==========================================
# 2. EKSEKUSI PENARIKAN DATA KANDIDAT
# ==========================================
Write-Host "Menghubungkan ke Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Menarik data item dengan kriteria $kolomFiling = 'LKPM REPORT'..." -ForegroundColor Cyan
$camlQuery = "<View><Query><Where><Eq><FieldRef Name='$kolomFiling'/><Value Type='Text'>LKPM REPORT</Value></Eq></Where></Query></View>"
$targetItems = Get-PnPListItem -List $listName -Query $camlQuery -ErrorAction SilentlyContinue

if (-not $targetItems) {
    Write-Host "[INFO] Tidak ditemukan item yang perlu diupdate." -ForegroundColor Yellow
    Disconnect-PnPOnline; exit
}

$totalItems = $targetItems.Count
Write-Host "Total ada $totalItems item yang akan dieksekusi perubahannya.`n" -ForegroundColor Magenta

# ==========================================
# 3. PROSES UPDATE MASSAL (LIVE EKSEKUSI)
# ==========================================
Write-Host "Memulai proses update Deadline Date ke tanggal 15..." -ForegroundColor Yellow
$counter = 0
$successCount = 0

foreach ($item in $targetItems) {
    $counter++
    $currentDateVal = $item[$kolomDeadline]
    
    # Ambil info identitas klien untuk log pencatatan
    $klienVal = $item[$kolomCLIENT]
    if ($null -ne $klienVal.LookupValue) { $klienVal = $klienVal.LookupValue }
    
    if ($null -ne $currentDateVal) {
        try {
            # 1. Konversi ke objek DateTime untuk mengambil Bulan dan Tahun asli
            $dt = [DateTime]$currentDateVal
            
            # 2. Buat target waktu baru dengan mengunci tanggal menjadi tanggal 15
            $newDateObj = [DateTime]::new($dt.Year, $dt.Month, 15, $dt.Hour, $dt.Minute, $dt.Second)
            
            # 3. Lakukan update ke SharePoint menggunakan Set-PnPListItem
            Set-PnPListItem -List $listName -Identity $item.Id -Values @{
                $kolomDeadline = $newDateObj
            } | Out-Null
            
            $successCount++
            Write-Host "   [$counter/$totalItems] [OK] Item ID $($item.Id) ($klienVal) -> Berhasil diubah ke tanggal: $($newDateObj.ToString('15-MM-yyyy'))" -ForegroundColor Green
        }
        catch {
            Write-Host "   [$counter/$totalItems] [FAIL] Gagal mengupdate Item ID $($item.Id) ($klienVal): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "   [$counter/$totalItems] [SKIP] Item ID $($item.Id) ($klienVal) dilewati karena Deadline Date awal kosong." -ForegroundColor Gray
    }
}

# ==========================================
# 4. LAPORAN SELESAI
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "        PROSES UPDATE DATA SELESAI!       " -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " Total Item Diproses : $totalItems" -ForegroundColor White
Write-Host " Sukses Terupdate    : $successCount" -ForegroundColor Green
Write-Host " Gagal/Dilewati      : $($totalItems - $successCount)" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Green

Disconnect-PnPOnline