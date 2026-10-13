# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ============================================================
# Backfill Version History dari SharePoint List ke CSV
# Untuk kemudian di-import ke Dataverse: MonitorKlienLog
# ============================================================


# --- KONFIGURASI (WAJIB DISESUAIKAN) ---
$listName       = "MONITOR KLIEN"
$outputCsv      = (Join-Path $PSScriptRoot "data\MonitorKlienLog_Backfill.csv")
$checkpointFile = (Join-Path $PSScriptRoot "data\checkpoint_processed_items.txt")
$batchSize      = 25   # simpan progress tiap 25 item, biar gak kehilangan banyak kalau crash


# Mapping Internal Name SharePoint -> Nama Kolom di Dataverse
# GANTI value di sebelah kanan sesuai hasil Get-PnPField langkah 3
$fieldMap = @{
    KLIEN               = "KLIEN"                 # ganti kalau internal name beda
    DITUGASKAN          = "DITUGASKAN"
    TGLDokumenLengkap    = "TGL_DOKUMEN_LENGKAP"    # contoh: internal name sering tanpa spasi
    TGLSelesai           = "TGL_SELESAI"
    STATUS              = "STATUS"
    CO                  = "CO"
    PIC                 = "PIC"
    DEPARTMENT          = "DEPARTMENT"             # CEK ULANG - ini yang di-rename
    LAYANAN             = "LAYANAN"
    INSTANSI            = "INSTANSI"
    TAKA                = "TAKA"
    NOTE                = "NOTE"
    PROGRESS            = "PROGRESS_x0020_x0028_x0025_x0029_"  # contoh internal name dgn karakter khusus
}

# --- KREDENSIAL APP-ONLY AUTHENTICATION (ISI SENDIRI, JANGAN SHARE KE SIAPAPUN) ---

$siteUrl      = $Config.HubSiteUrl
$tenant       = $Config.TenantDomain
$clientId     = $Config.AppClientId
$pfxPath      = $Config.PfxPath
$pfxPassword  = $Config.PfxPassword



# --- KONEKSI ---
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Mengambil item dari list '$listName' dengan filter..." -ForegroundColor Cyan

# Filter: DITUGASKAN tahun ini (>= 1 Jan tahun berjalan) ATAU STATUS bukan SELESAI
# GANTI 'DITUGASKAN' dan 'STATUS' di bawah kalau Internal Name-nya beda (cek lagi via Get-PnPField)
$tahunIni = (Get-Date).Year
$camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Or>
        <Geq>
          <FieldRef Name='DITUGASKAN' />
          <Value Type='DateTime'>$tahunIni-01-01T00:00:00Z</Value>
        </Geq>
        <Neq>
          <FieldRef Name='STATUS' />
          <Value Type='Text'>SELESAI</Value>
        </Neq>
      </Or>
    </Where>
  </Query>
  <RowLimit Paged='TRUE'>500</RowLimit>
</View>
"@

$items = Get-PnPListItem -List $listName -Query $camlQuery -PageSize 500

Write-Host "Total item yang lolos filter (ditugaskan $tahunIni ATAU belum selesai): $($items.Count)" -ForegroundColor Green

# Pastikan folder output ada
$outputFolder = Split-Path $outputCsv -Parent
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

# Baca checkpoint - ItemID yang sudah pernah diproses sebelumnya
$processedIds = [System.Collections.Generic.HashSet[int]]::new()
if (Test-Path $checkpointFile) {
    Get-Content $checkpointFile | ForEach-Object {
        if ($_ -match '^\d+$') { $processedIds.Add([int]$_) | Out-Null }
    }

    # PENGAMAN: kalau checkpoint bilang ada item yang sudah diproses,
    # tapi file CSV output-nya gak ada/hilang, STOP - jangan lanjut diam-diam,
    # karena itu artinya data yang sudah diproses sebelumnya HILANG.
    if ($processedIds.Count -gt 0 -and -not (Test-Path $outputCsv)) {
        Write-Error "STOP: Checkpoint mencatat $($processedIds.Count) item sudah diproses, tapi file CSV '$outputCsv' tidak ditemukan. Data sebelumnya kemungkinan hilang/terhapus/pindah lokasi. Kalau memang mau mulai ULANG DARI NOL, hapus dulu file checkpoint '$checkpointFile' secara manual, baru jalankan script ini lagi."
        exit 1
    }

    Write-Host "Ditemukan checkpoint: $($processedIds.Count) item sudah pernah diproses sebelumnya, akan di-skip." -ForegroundColor Yellow
}

$results = [System.Collections.Generic.List[Object]]::new()
$counter = 0
$csvHeaderWritten = (Test-Path $outputCsv) -and ($processedIds.Count -gt 0)

foreach ($item in $items) {
    $counter++

    if ($processedIds.Contains($item.Id)) {
        continue  # sudah pernah diproses, skip
    }

    Write-Progress -Activity "Menarik version history" -Status "Item $counter dari $($items.Count) (ItemID: $($item.Id))" -PercentComplete (($counter / $items.Count) * 100)

    try {
        # Ambil semua versi untuk item ini
        Get-PnPProperty -ClientObject $item -Property Versions

        foreach ($version in $item.Versions) {
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
        }

        # Tandai item ini sudah selesai diproses
        Add-Content -Path $checkpointFile -Value $item.Id
    }
    catch {
        Write-Warning "Gagal proses ItemID $($item.Id): $($_.Exception.Message)"
        # SENGAJA tidak ditambahkan ke checkpoint, biar item ini dicoba lagi di run berikutnya
    }

    # Simpan progress tiap $batchSize item, biar gak numpuk semua di memory & aman kalau crash
    if ($results.Count -ge $batchSize) {
        if (-not $csvHeaderWritten) {
            $results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
            $csvHeaderWritten = $true
        } else {
            $results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8 -Append
        }
        Write-Host "  -> Progress disimpan: $($results.Count) baris ditambahkan ke CSV." -ForegroundColor DarkGray
        $results.Clear()
    }
}

# Simpan sisa data yang belum ke-flush di batch terakhir
if ($results.Count -gt 0) {
    if (-not $csvHeaderWritten) {
        $results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
    } else {
        $results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8 -Append
    }
}

Write-Host "Selesai. File tersimpan di: $outputCsv" -ForegroundColor Green
Write-Host "Kalau ada item yang gagal (lihat warning di atas), cukup jalankan ulang script ini - item yang sudah sukses akan otomatis di-skip." -ForegroundColor Yellow
Write-Host "Langkah berikutnya: import CSV ini ke tabel MonitorKlienLog di Dataverse." -ForegroundColor Yellow
