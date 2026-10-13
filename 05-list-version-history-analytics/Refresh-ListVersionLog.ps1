# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ============================================================
# Refresh Incremental - Ambil perubahan BARU sejak run terakhir
# Melengkapi Backfill_MonitorKlienLog.ps1 (yang narik histori lama)
# Jalankan script ini berkala (manual atau lewat Task Scheduler)
# ============================================================

# --- KREDENSIAL APP-ONLY AUTHENTICATION (ISI SENDIRI, JANGAN SHARE KE SIAPAPUN) ---
$siteUrl      = $Config.HubSiteUrl
$tenant       = $Config.TenantDomain
$clientId     = $Config.AppClientId
$pfxPath      = $Config.PfxPath
$pfxPassword  = $Config.PfxPassword

# --- KONFIGURASI ---
$listName            = "MONITOR KLIEN"
$outputCsv           = (Join-Path $PSScriptRoot "data\MonitorKlienLog_Backfill.csv")       # CSV YANG SAMA dengan hasil backfill
$watermarkFile        = (Join-Path $PSScriptRoot "data\watermark.txt")
$exportedVersionsFile = (Join-Path $PSScriptRoot "data\exported_versions.txt")             # jejak ItemID|VersionID yang sudah pernah masuk CSV

# Mapping Internal Name SharePoint -> Nama Kolom di Dataverse (SAMAKAN dengan script backfill)
$fieldMap = @{
    KLIEN               = "KLIEN"
    DITUGASKAN          = "DITUGASKAN"
    TGLDokumenLengkap    = "TGL_DOKUMEN_LENGKAP"
    TGLSelesai           = "TGL_SELESAI"
    STATUS              = "STATUS"
    CO                  = "CO"
    PIC                 = "PIC"
    DEPARTMENT          = "DEPARTMENT"
    LAYANAN             = "LAYANAN"
    INSTANSI            = "INSTANSI"
    TAKA                = "TAKA"
    NOTE                = "NOTE"
    PROGRESS            = "PROGRESS_x0020_x0028_x0025_x0029_"
}

# --- CATAT WAKTU MULAI (dipakai sebagai watermark BARU di akhir, bukan waktu selesai) ---
# Ini penting: kalau ada perubahan data yang terjadi SAAT script ini masih berjalan,
# watermark baru tetap "aman" nyakup perubahan itu di run berikutnya.
$runStartTime = (Get-Date).ToUniversalTime()

# --- BACA WATERMARK ---
if (-not (Test-Path $watermarkFile)) {
    Write-Error "STOP: File watermark '$watermarkFile' tidak ditemukan. Kalau ini run PERTAMA setelah backfill selesai, buat file ini manual dulu, isi dengan waktu backfill terakhir kali selesai (format: yyyy-MM-ddTHH:mm:ssZ, contoh: 2026-08-27T12:00:00Z)."
    exit 1
}

$watermarkRaw = Get-Content $watermarkFile -Raw
$watermark = [datetime]::Parse($watermarkRaw.Trim(), $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
$watermarkIso = $watermark.ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "Watermark terakhir: $watermarkIso" -ForegroundColor Cyan
Write-Host "Mencari perubahan sejak waktu itu..." -ForegroundColor Cyan

# --- BACA DAFTAR VERSI YANG SUDAH PERNAH DI-EXPORT (biar gak dobel di CSV) ---
$exportedVersions = [System.Collections.Generic.HashSet[string]]::new()
if (Test-Path $exportedVersionsFile) {
    Get-Content $exportedVersionsFile | ForEach-Object { $exportedVersions.Add($_) | Out-Null }
}

# --- KONEKSI ---
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

# --- QUERY: cuma item yang Modified >= watermark ---
$camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Geq>
        <FieldRef Name='Modified' />
        <Value Type='DateTime' IncludeTimeValue='TRUE'>$watermarkIso</Value>
      </Geq>
    </Where>
  </Query>
  <RowLimit Paged='TRUE'>500</RowLimit>
</View>
"@

$items = Get-PnPListItem -List $listName -Query $camlQuery -PageSize 500

Write-Host "Item yang berubah sejak watermark: $($items.Count)" -ForegroundColor Green

if ($items.Count -eq 0) {
    Write-Host "Tidak ada perubahan baru. Selesai." -ForegroundColor Yellow
    # Tetap update watermark meskipun gak ada perubahan, biar run berikutnya gak query ulang rentang yang sama
    $runStartTime.ToString("yyyy-MM-ddTHH:mm:ssZ") | Set-Content -Path $watermarkFile
    exit 0
}

$results = [System.Collections.Generic.List[Object]]::new()
$counter = 0
$anyFatalError = $false

foreach ($item in $items) {
    $counter++
    Write-Progress -Activity "Cek perubahan" -Status "Item $counter dari $($items.Count) (ItemID: $($item.Id))" -PercentComplete (($counter / $items.Count) * 100)

    try {
        Get-PnPProperty -ClientObject $item -Property Versions

        foreach ($version in $item.Versions) {
            $versionKey = "$($item.Id)|$($version.VersionId)"

            # Skip kalau versi ini sudah pernah ke-export sebelumnya
            if ($exportedVersions.Contains($versionKey)) { continue }

            # Skip juga versi yang lebih lama dari watermark (bukan bagian dari perubahan baru)
            if ($version.Created -lt $watermark) { continue }

            $fv = $version.FieldValues

            $row = [PSCustomObject]@{
                ItemID              = $item.Id
                VersionID           = $version.VersionId
                KLIEN               = $fv[$fieldMap.KLIEN]
                DITUGASKAN          = $fv[$fieldMap.DITUGASKAN]
                TGL_DOKUMEN_LENGKAP = $fv[$fieldMap.TGLDokumenLengkap]
                TGL_SELESAI         = $fv[$fieldMap.TGLSelesai]
                STATUS              = $fv[$fieldMap.STATUS]
                CO                  = $fv[$fieldMap.CO]
                PIC                 = if ($fv[$fieldMap.PIC]) { $fv[$fieldMap.PIC].LookupValue } else { $null }
                DEPARTMENT          = $fv[$fieldMap.DEPARTMENT]
                LAYANAN             = $fv[$fieldMap.LAYANAN]
                INSTANSI            = $fv[$fieldMap.INSTANSI]
                TAKA                = $fv[$fieldMap.TAKA]
                NOTE                = $fv[$fieldMap.NOTE]
                PROGRESS            = $fv[$fieldMap.PROGRESS]
                ModifiedBy          = $version.CreatedBy.LookupValue
                TimestampPerubahan  = $version.Created
                Created             = $item.FieldValues["Created"]
            }

            $results.Add($row)
            Add-Content -Path $exportedVersionsFile -Value $versionKey
        }
    }
    catch {
        Write-Warning "Gagal proses ItemID $($item.Id): $($_.Exception.Message)"
        $anyFatalError = $true
    }
}

if ($results.Count -gt 0) {
    $results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8 -Append
    Write-Host "Ditambahkan $($results.Count) baris baru ke CSV." -ForegroundColor Green
} else {
    Write-Host "Tidak ada baris baru yang perlu ditambahkan (semua versi sudah pernah ke-capture)." -ForegroundColor Yellow
}

if ($anyFatalError) {
    Write-Warning "Ada item yang gagal diproses. Watermark TIDAK di-update, biar item yang gagal itu dicoba lagi di run berikutnya."
} else {
    # Cuma update watermark kalau semua item berhasil diproses tanpa error
    $runStartTime.ToString("yyyy-MM-ddTHH:mm:ssZ") | Set-Content -Path $watermarkFile
    Write-Host "Watermark diupdate ke: $($runStartTime.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -ForegroundColor Green
}

Write-Host "Selesai." -ForegroundColor Green
