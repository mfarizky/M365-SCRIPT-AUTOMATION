# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL
# ==========================================
$siteURL     = $Config.HubSiteUrl   # [VERIFIKASI] site tempat MONITOR KLIEN
$siteURLWF   = $siteURL                                           # ganti kalau DATABASE WORKFLOWS ada di site lain

$listMonitor = "MONITOR KLIEN"
$listWF      = "DATABASE WORKFLOWS"

$clientId    = $Config.AppClientId
$tenant      = $Config.TenantDomain
$pfxPath     = $Config.PfxPath
$pfxPassword =  $Config.PfxPassword

# Internal name kolom (cek dengan $cekKolom = $true)
$kolomStatus  = "STATUS"     # MONITOR KLIEN
$kolomLayanan = "LAYANAN"    # MONITOR KLIEN            [VERIFIKASI]
$kolomTED     = "TED"        # MONITOR KLIEN (Number)   [VERIFIKASI]
$kolomWFKey   = "field_1"    # DATABASE WORKFLOWS (nama layanan)
$kolomWFTED   = "TED"        # DATABASE WORKFLOWS       [VERIFIKASI]

# Filter status MONITOR KLIEN
$statusAktif = @("BELUM MULAI", "BERLANGSUNG", "PRIORITY")

# Konversi TED teks -> hari
$mapTED = @{
    "1 WEEK"   = 7
    "2 WEEKS"  = 14
    "3 WEEKS"  = 21
    "1 MONTH"  = 30
    "3 MONTHS" = 90
}

$testID   = $null     # ID item untuk tes 1 item saja; $null untuk semua item
$dryRun   = $false    # $false untuk eksekusi live
$cekKolom = $false   # $true untuk menampilkan internal name kolom lalu berhenti

# ==========================================
# 1b. KONEKSI (+ OPSIONAL CEK INTERNAL NAME)
# ==========================================
Write-Host "Menghubungkan ke Site..." -ForegroundColor Cyan
$connMon = Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword -ReturnConnection
$connWF  = if ($siteURLWF -eq $siteURL) { $connMon } else {
    Connect-PnPOnline -Url $siteURLWF -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword -ReturnConnection
}

if ($cekKolom) {
    foreach ($pair in @(@($listMonitor, $connMon), @($listWF, $connWF))) {
        Write-Host "`n=== $($pair[0]) ===" -ForegroundColor Yellow
        Get-PnPField -List $pair[0] -Connection $pair[1] | Where-Object { -not $_.Hidden -and -not $_.ReadOnlyField } |
            Select-Object Title, InternalName, TypeAsString | Sort-Object Title | Format-Table -AutoSize
    }
    exit
}

# Helper: ambil teks dari kolom Text / Choice / Lookup, lalu rapikan
function Get-Teks($val) {
    if ($null -eq $val) { return "" }
    if ($val -is [array]) { $val = $val[0] }
    if ($null -ne $val.LookupValue) { $val = $val.LookupValue }
    return ("$val" -replace '\s+', ' ').Trim().ToUpper()
}

# ==========================================
# 2. BACA DATABASE WORKFLOWS -> PETA LAYANAN = HARI
# ==========================================
Write-Host "Menarik DATABASE WORKFLOWS..." -ForegroundColor Cyan
$wfItems = @(Get-PnPListItem -List $listWF -Fields $kolomWFKey, $kolomWFTED -PageSize 2000 -Connection $connWF)

$mapLayanan = @{}
$konflik    = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($w in $wfItems) {
    $key = Get-Teks $w[$kolomWFKey]
    if (-not $key) { continue }

    $tedTeks = Get-Teks $w[$kolomWFTED]
    if (-not $mapTED.ContainsKey($tedTeks)) {
        Write-Host "   [WARN] WF '$key' - TED '$tedTeks' tidak dikenal, dilewati." -ForegroundColor Yellow
        continue
    }
    $hari = $mapTED[$tedTeks]

    if ($mapLayanan.ContainsKey($key) -and $mapLayanan[$key] -ne $hari) {
        [void]$konflik.Add($key)
        Write-Host "   [WARN] WF '$key' - muncul lebih dari sekali dengan TED berbeda. Cek manual." -ForegroundColor Yellow
        continue
    }
    $mapLayanan[$key] = $hari
}
Write-Host "Layanan terpetakan: $($mapLayanan.Count)`n" -ForegroundColor Magenta

# ==========================================
# 3. BACA MONITOR KLIEN + FILTER STATUS
# ==========================================
Write-Host "Menarik MONITOR KLIEN..." -ForegroundColor Cyan
$items  = @(Get-PnPListItem -List $listMonitor -Fields $kolomStatus, $kolomLayanan, $kolomTED -PageSize 2000 -Connection $connMon)
$target = @($items | Where-Object { $statusAktif -contains (Get-Teks $_[$kolomStatus]) })
if ($testID) {
    $target = @($target | Where-Object { $_.Id -eq $testID })
    if ($target.Count -eq 0) { Write-Host "Item ID $testID tidak ada / statusnya di luar filter." -ForegroundColor Yellow }
}
Write-Host "Total item: $($items.Count) | Sesuai filter status: $($target.Count)" -ForegroundColor Magenta
Write-Host "Mode: $(if ($dryRun) {'DRY RUN (simulasi)'} else {'LIVE'})$(if ($testID) {" | TES ID $testID"})`n" -ForegroundColor Magenta

# ==========================================
# 4. PROSES UPDATE TED
# ==========================================
$counter = 0; $okCount = 0; $samaCount = 0; $skipCount = 0; $failCount = 0
$layananTakAda = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($it in $target) {
    $counter++
    $prefix  = "   [$counter/$($target.Count)] ID $($it.Id)"
    $layanan = Get-Teks $it[$kolomLayanan]

    if (-not $layanan) {
        $skipCount++; Write-Host "$prefix [SKIP] LAYANAN kosong." -ForegroundColor Gray; continue
    }
    if ($konflik.Contains($layanan)) {
        $skipCount++; Write-Host "$prefix [SKIP] '$layanan' punya TED ganda di WF." -ForegroundColor Yellow; continue
    }
    if (-not $mapLayanan.ContainsKey($layanan)) {
        $skipCount++; [void]$layananTakAda.Add($layanan)
        Write-Host "$prefix [SKIP] '$layanan' tidak ada di DATABASE WORKFLOWS." -ForegroundColor Yellow; continue
    }

    $hari    = $mapLayanan[$layanan]
    $tedLama = $it[$kolomTED]
    if ("$tedLama" -eq "$hari") {
        $samaCount++; Write-Host "$prefix [SAMA] $layanan sudah $hari hari." -ForegroundColor Gray; continue
    }

    if ($dryRun) {
        $okCount++
        Write-Host "$prefix [DRY] $layanan -> TED $(if ($null -eq $tedLama) {'(kosong)'} else {$tedLama}) jadi $hari" -ForegroundColor Cyan
        continue
    }

    try {
        # SystemUpdate: tidak mengubah Modified/Editor dan tidak membuat versi baru
        Set-PnPListItem -List $listMonitor -Identity $it.Id -Values @{ $kolomTED = $hari } -UpdateType SystemUpdate -Connection $connMon | Out-Null
        $okCount++
        Write-Host "$prefix [OK] $layanan -> TED $hari" -ForegroundColor Green
    }
    catch {
        $failCount++
        Write-Host "$prefix [FAIL] $layanan : $_" -ForegroundColor Red
    }
}

# ==========================================
# 5. LAPORAN
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "        PROSES SELESAI $(if ($dryRun) {'(DRY RUN)'})" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " Item sesuai filter   : $($target.Count)" -ForegroundColor White
Write-Host " $(if ($dryRun) {'Akan diupdate'} else {'Berhasil diupdate'})     : $okCount" -ForegroundColor Green
Write-Host " Sudah sama           : $samaCount" -ForegroundColor Gray
Write-Host " Dilewati             : $skipCount" -ForegroundColor Yellow
Write-Host " Gagal                : $failCount" -ForegroundColor Red
if ($layananTakAda.Count -gt 0) {
    Write-Host "`n LAYANAN tidak ditemukan di DATABASE WORKFLOWS:" -ForegroundColor Yellow
    $layananTakAda | Sort-Object | ForEach-Object { Write-Host "   - $_" -ForegroundColor Yellow }
}
Write-Host "==========================================" -ForegroundColor Green

Disconnect-PnPOnline -Connection $connMon
if ($connWF -ne $connMon) { Disconnect-PnPOnline -Connection $connWF }