# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI GLOBAL
# ==========================================
$siteURL        = $Config.HubSiteUrl
$listName       = "REGULATORY COMPLIANCE"

$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Internal name kolom (CEK DULU yang bertanda [VERIFIKASI] - lihat bagian 1b)
$kolomIDRC      = "IDKLIENRC"
$kolomCLIENT    = "KLIEN"
$kolomFiling    = "COMPLIANCE"      # COMPLIANCE FILING
$kolomPeriod    = "PERIOD"    # [VERIFIKASI] FILING PERIOD
$kolomYear      = "TAHUN"            # [VERIFIKASI] YEAR
$kolomDeadline  = "DEADLINEDATE"
$kolomStatus    = "STATUS"          # [VERIFIKASI] STATUS

# Nilai row baru
$nilaiFiling    = "4"
$nilaiPeriod    = "-"
$nilaiYear      = "2026"
$nilaiDeadline  = [DateTime]::new(2026, 11, 30, 0, 0, 0, [DateTimeKind]::Local)
$nilaiStatus    = "SCHEDULED"

$jumlahTarget   = 6        # PT dengan jumlah row ini yang akan ditambah
$dryRun         = $false   # $false untuk eksekusi live
$cekKolom       = $false   # $true untuk menampilkan daftar internal name kolom lalu berhenti

# ==========================================
# 1b. KONEKSI (+ OPSIONAL CEK INTERNAL NAME)
# ==========================================
Write-Host "Menghubungkan ke Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

if ($cekKolom) {
    Get-PnPField -List $listName | Where-Object { -not $_.Hidden -and -not $_.ReadOnlyField } |
        Select-Object Title, InternalName, TypeAsString | Sort-Object Title | Format-Table -AutoSize
    Disconnect-PnPOnline; exit
}

# ==========================================
# 2. TARIK SEMUA ROW
# ==========================================
Write-Host "Menarik semua item..." -ForegroundColor Cyan
$items = @(Get-PnPListItem -List $listName -Fields $kolomIDRC, $kolomCLIENT, $kolomFiling -PageSize 2000)
Write-Host "Total item: $($items.Count)`n" -ForegroundColor Magenta

# Kumpulan semua ID yang sudah terpakai (untuk cegah duplikat)
$semuaID = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($i in $items) { if ($i[$kolomIDRC]) { [void]$semuaID.Add("$($i[$kolomIDRC])".Trim()) } }

# Helper: ambil nama klien (Text / Lookup)
function Get-NamaKlien($val) {
    if ($null -eq $val) { return $null }
    if ($val -is [array]) { $val = $val[0] }
    if ($null -ne $val.LookupValue) { return $val.LookupValue.Trim() }
    return "$val".Trim()
}

# Kelompokkan per PT
$grupPT = $items | Where-Object { Get-NamaKlien $_[$kolomCLIENT] } |
    Group-Object { (Get-NamaKlien $_[$kolomCLIENT]).ToUpper() }

$kandidat = @($grupPT | Where-Object { $_.Count -eq $jumlahTarget })
Write-Host "PT dengan $jumlahTarget row: $($kandidat.Count)" -ForegroundColor Magenta
Write-Host "Mode: $(if ($dryRun) {'DRY RUN (simulasi)'} else {'LIVE'})`n" -ForegroundColor Magenta

# ==========================================
# 3. PROSES ADD ROW
# ==========================================
$counter = 0; $okCount = 0; $skipCount = 0; $failCount = 0

foreach ($g in $kandidat) {
    $counter++
    $contoh   = $g.Group[0]
    $namaPT   = Get-NamaKlien $contoh[$kolomCLIENT]
    $prefix   = "   [$counter/$($kandidat.Count)]"

    # Skip jika PT sudah punya row KBLI
    $sudahAda = $g.Group | Where-Object { "$($_[$kolomFiling])".Trim() -eq $nilaiFiling }
    if ($sudahAda) {
        $skipCount++
        Write-Host "$prefix [SKIP] $namaPT - sudah punya row KBLI." -ForegroundColor Gray
        continue
    }

    # Cari ID terbesar milik PT ini
    $ids = @($g.Group | ForEach-Object { "$($_[$kolomIDRC])".Trim() } | Where-Object { $_ -match '^\d+$' })
    if ($ids.Count -eq 0) {
        $skipCount++
        Write-Host "$prefix [SKIP] $namaPT - ID KLIEN RC tidak ditemukan / bukan angka." -ForegroundColor Yellow
        continue
    }
    $maxID  = $ids | Sort-Object { [long]$_ } | Select-Object -Last 1
    $newID  = ([long]$maxID + 1).ToString().PadLeft($maxID.Length, '0')   # jaga angka 0 di depan

    if ($semuaID.Contains($newID)) {
        $skipCount++
        Write-Host "$prefix [SKIP] $namaPT - ID baru $newID sudah dipakai item lain. Cek manual." -ForegroundColor Yellow
        continue
    }

    # Nilai KLIEN: pakai LookupId kalau kolom Lookup, kalau Text pakai nama
    $klienRaw = $contoh[$kolomCLIENT]
    if ($klienRaw -is [array]) { $klienRaw = $klienRaw[0] }
    $nilaiKlien = if ($null -ne $klienRaw.LookupId) { $klienRaw.LookupId } else { $namaPT }

    $values = @{
        $kolomIDRC     = $newID
        $kolomCLIENT   = $nilaiKlien
        $kolomFiling   = $nilaiFiling
        $kolomPeriod   = $nilaiPeriod
        $kolomYear     = $nilaiYear
        $kolomDeadline = $nilaiDeadline
        $kolomStatus   = $nilaiStatus
    }

    if ($dryRun) {
        $okCount++
        Write-Host "$prefix [DRY] $namaPT -> ID terakhir $maxID, row baru ID $newID" -ForegroundColor Cyan
        continue
    }

    try {
        $new = Add-PnPListItem -List $listName -Values $values
        [void]$semuaID.Add($newID)
        $okCount++
        Write-Host "$prefix [OK] $namaPT -> row baru ID $newID (Item ID $($new.Id))" -ForegroundColor Green
    }
    catch {
        $failCount++
        Write-Host "$prefix [FAIL] $namaPT : $_" -ForegroundColor Red
    }
}

# ==========================================
# 4. LAPORAN
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "        PROSES SELESAI $(if ($dryRun) {'(DRY RUN)'})" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host " PT kandidat ($jumlahTarget row) : $($kandidat.Count)" -ForegroundColor White
Write-Host " $(if ($dryRun) {'Akan ditambah'} else {'Berhasil ditambah'}) : $okCount" -ForegroundColor Green
Write-Host " Dilewati             : $skipCount" -ForegroundColor Yellow
Write-Host " Gagal                : $failCount" -ForegroundColor Red
Write-Host "==========================================" -ForegroundColor Green

Disconnect-PnPOnline