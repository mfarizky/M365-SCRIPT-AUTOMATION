# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI KREDENSIAL & PARAMETER
# ==========================================
$siteUrl        = $Config.SharePointRoot
$tenant         = $Config.TenantDomain
$clientId       = $Config.AppClientId
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

$outputPath     = (Join-Path $Config.ReportFolder "Internal_Sites_FileSize_Summary.csv")

# Daftar 5 Site Internal yang ingin diperiksa saja
$targetSites = @(
    @{ Name = "GKI LEGAL";       UrlPath = "/sites/legal" },
    @{ Name = "GKI SERVICES";    UrlPath = "/sites/services" }, # Mengcover /sites/service atau /sites/services
    @{ Name = "GKI FINANCE";     UrlPath = "/sites/finance" },
    @{ Name = "GKI DIGIHUB";      UrlPath = "/sites/digihub" },
    @{ Name = "GKI REAL ESTATE"; UrlPath = "/sites/real" } # Mengcover /sites/realestate atau /sites/real-estate
)

# ==========================================
# 2. KONEKSI KE SHAREPOINT MENGGUNAKAN SERTIFIKAT
# ==========================================
Write-Host "Menghubungkan ke SharePoint Tenant..." -ForegroundColor Cyan

Connect-PnPOnline -Url $siteUrl `
                  -ClientId $clientId `
                  -Tenant $tenant `
                  -CertificatePath $pfxPath `
                  -CertificatePassword $pfxPassword

Write-Host "Berhasil terhubung!" -ForegroundColor Green

# ==========================================
# 3. QUERY PENCARIAN TERFOKUS PADA 5 SITE INTERNAL
# ==========================================
# Buat filter KQL Path khusus untuk 5 site internal tersebut
$pathFilters = ($targetSites | ForEach-Object { "Path:`"$siteUrl$($_.UrlPath)*`"" }) -join " OR "
$searchQuery = "IsContainer=false AND Size>=3145728 AND ($pathFilters)"

Write-Host "Mengumpulkan data file >= 3MB khusus dari 5 Site Internal..." -ForegroundColor Yellow

$allFiles = @()
$startRow = 0
$pageSize = 500

do {
    $results = Submit-PnPSearchQuery -Query $searchQuery `
                                     -SelectProperties "Title", "Path", "FileType", "Size", "SPSiteUrl" `
                                     -StartRow $startRow `
                                     -MaxResults $pageSize
    
    if ($results.ResultRows.Count -gt 0) {
        foreach ($row in $results.ResultRows) {
            $sitePath = $row["SPSiteUrl"].ToLower()
            
            # Cocokkan nama site berdasarkan URL
            $matchedSiteName = "INTERNAL LAINNYA"
            foreach ($ts in $targetSites) {
                $checkPath = $ts.UrlPath.ToLower().TrimEnd('s') # fleksibel untuk service/services, realestate/real-estate
                if ($sitePath -like "*$checkPath*") {
                    $matchedSiteName = $ts.Name
                    break
                }
            }

            # Ambil ekstensi file
            $ext = $row["FileType"]
            if (-not $ext) { $ext = "NO EXTENSION" } else { $ext = $ext.ToLower() }

            $allFiles += [PSCustomObject]@{
                SiteName = $matchedSiteName
                SiteUrl  = $row["SPSiteUrl"]
                FileType = $ext
                SizeBytes= [long]$row["Size"]
            }
        }
        $startRow += $results.ResultRows.Count
    }
} while ($results.ResultRows.Count -eq $pageSize)

Write-Host "Total file teridentifikasi di 5 site internal: $($allFiles.Count) file." -ForegroundColor Green

# ==========================================
# 4. REKAPITULASI & PENGELOMPOKAN
# ==========================================
$summaryReport = @()

# Ambil daftar unik site dan ekstensi dari hasil pencarian
$uniqueSitesFound = $allFiles | Select-Object -ExpandProperty SiteName -Unique
$allExtensionsFound = $allFiles | Select-Object -ExpandProperty FileType -Unique | Sort-Object

foreach ($siteName in ($targetSites.Name | Select-Object -Unique)) {
    foreach ($ext in $allExtensionsFound) {
        
        # Filter file berdasarkan Site & Ekstensi
        $filtered = $allFiles | Where-Object { $_.SiteName -eq $siteName -and $_.FileType -eq $ext }

        if ($filtered.Count -gt 0) {
            $c3to5   = ($filtered | Where-Object { $_.SizeBytes -ge 3145728 -and $_.SizeBytes -le 5242880 }).Count
            $c5to8   = ($filtered | Where-Object { $_.SizeBytes -gt 5242880 -and $_.SizeBytes -le 8388608 }).Count
            $cAbove8 = ($filtered | Where-Object { $_.SizeBytes -gt 8388608 }).Count

            $summaryReport += [PSCustomObject]@{
                "Nama Site / Unit" = $siteName
                "Ekstensi File"   = $ext.ToUpper()
                "3 - 5 MB"         = $c3to5
                "5 - 8 MB"         = $c5to8
                "> 8 MB"           = $cAbove8
                "Total File"       = ($c3to5 + $c5to8 + $cAbove8)
            }
        }
    }
}

# Tambahkan Baris TOTAL 5 SITE INTERNAL di Bagian Bawah
foreach ($ext in $allExtensionsFound) {
    $filteredExt = $allFiles | Where-Object { $_.FileType -eq $ext }
    if ($filteredExt.Count -gt 0) {
        $tot3to5   = ($filteredExt | Where-Object { $_.SizeBytes -ge 3145728 -and $_.SizeBytes -le 5242880 }).Count
        $tot5to8   = ($filteredExt | Where-Object { $_.SizeBytes -gt 5242880 -and $_.SizeBytes -le 8388608 }).Count
        $totAbove8 = ($filteredExt | Where-Object { $_.SizeBytes -gt 8388608 }).Count

        $summaryReport += [PSCustomObject]@{
            "Nama Site / Unit" = "=== TOTAL 5 SITE INTERNAL ==="
            "Ekstensi File"   = $ext.ToUpper()
            "3 - 5 MB"         = $tot3to5
            "5 - 8 MB"         = $tot5to8
            "> 8 MB"           = $totAbove8
            "Total File"       = ($tot3to5 + $tot5to8 + $totAbove8)
        }
    }
}

# ==========================================
# 5. EKSPOR HASIL KE CSV
# ==========================================
if ($summaryReport.Count -gt 0) {
    $folderPath = Split-Path $outputPath
    if (-not (Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    }

    $summaryReport | Export-Csv -Path $outputPath -NoTypeInformation -Encoding utf8BOM
    Write-Host "`n[SUKSES] Ringkasan data 5 site internal berhasil dibuat!" -ForegroundColor Green
    Write-Host "File laporan tersimpan di: $outputPath" -ForegroundColor Green
} else {
    Write-Host "`n[INFO] Tidak ditemukan file >= 3MB pada 5 site internal tersebut." -ForegroundColor Yellow
}

Disconnect-PnPOnline