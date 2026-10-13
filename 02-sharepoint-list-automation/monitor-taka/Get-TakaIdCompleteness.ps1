# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI TARGET SITE & LIST
# ==========================================
$siteUrl        = $Config.HubSiteUrl
$listName       = "MONITOR TAKA"

# Credentials & Cert Path
$tenant         = $Config.TenantDomain
$clientId       = $Config.AppClientId
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Path Output CSV untuk Unique Values
$csvReportFolder = $Config.ReportFolder
$csvReportPath   = "$csvReportFolder/IDMK_IDMK0_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

if (-not (Test-Path $csvReportFolder)) { New-Item -ItemType Directory -Path $csvReportFolder | Out-Null }

# ==========================================
# 2. KONEKSI DAN PENARIKAN DATA
# ==========================================
Write-Host "Connecting to SharePoint Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Fetching items from '$listName'..." -ForegroundColor Cyan
$allItems = Get-PnPListItem -List $listName -PageSize 2000 -ErrorAction SilentlyContinue

if (-not $allItems) {
    Write-Host "[ERROR] List '$listName' tidak ditemukan atau kosong!" -ForegroundColor Red
    Disconnect-PnPOnline; exit
}

# ==========================================
# 3. ANALISIS DATA (IDMK0 & ID MK / Lookup)
# ==========================================
$totalRows      = $allItems.Count

$bothFilled     = 0
$bothEmpty      = 0
$onlyIDMK0      = 0
$onlyIDMK       = 0

$idmk0List      = @()
$idmkList       = @()
$combinedReport = @()

foreach ($item in $allItems) {
    # 1. Ambil nilai IDMK0 (Teks Biasa)
    $valIDMK0 = [string]$item["IDMK0"]
    $isIDMK0Empty = [string]::IsNullOrWhiteSpace($valIDMK0)

    # 2. Ambil nilai 'ID MK' (Lookup / Expanded Reference)
    $rawIDMK = $item["ID_x0020_MK"] # Internal Name default SharePoint untuk "ID MK"
    if ($null -eq $rawIDMK) { $rawIDMK = $item["ID MK"] }
    
    $valIDMK = ""
    if ($rawIDMK -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
        $valIDMK = [string]$rawIDMK.LookupValue
    } elseif ($rawIDMK) {
        $valIDMK = [string]$rawIDMK
    }
    $isIDMKEmpty = [string]::IsNullOrWhiteSpace($valIDMK)

    # 3. Pengelompokan Kategori Kosong/Terisi
    if (-not $isIDMK0Empty -and -not $isIDMKEmpty) {
        $bothFilled++
    } elseif ($isIDMK0Empty -and $isIDMKEmpty) {
        $bothEmpty++
    } elseif (-not $isIDMK0Empty -and $isIDMKEmpty) {
        $onlyIDMK0++
    } else {
        $onlyIDMK++
    }

    # 4. Simpan ke List untuk analisis Unique Value
    if (-not $isIDMK0Empty) { $idmk0List += $valIDMK0.Trim() }
    if (-not $isIDMKEmpty)  { $idmkList  += $valIDMK.Trim() }

    # Array untuk ekspor detail
    $combinedReport += [PSCustomObject]@{
        ItemID       = $item.Id
        IDMK0_Value  = $valIDMK0
        IDMK_Value   = $valIDMK
        Status       = if (-not $isIDMK0Empty -and -not $isIDMKEmpty) { "Lengkap" } 
                       elseif ($isIDMK0Empty -and $isIDMKEmpty) { "Keduanya Kosong" }
                       elseif ($isIDMK0Empty) { "IDMK0 Kosong" } 
                       else { "ID MK Kosong" }
    }
}

# Grouping Unique Values
$uniqueIDMK0 = $idmk0List | Group-Object | Select-Object @{N='Value';E={$_.Name}}, @{N='Jumlah_Row_IDMK0';E={$_.Count}}
$uniqueIDMK  = $idmkList  | Group-Object | Select-Object @{N='Value';E={$_.Name}}, @{N='Jumlah_Row_IDMK';E={$_.Count}}

Disconnect-PnPOnline

# ==========================================
# 4. EKSPOR KE CSV
# ==========================================
if ($combinedReport) {
    $combinedReport | Export-Csv -Path $csvReportPath -NoTypeInformation -Delimiter "," -Encoding UTF8
}
Write-Host "`nRINGKASAN ITEM DENGAN STATUS KOSONG:" -ForegroundColor Red
$combinedReport | Where-Object { $_.Status -ne "Lengkap" } | Select-Object ItemID, IDMK0_Value, IDMK_Value, Status | Format-Table -AutoSize

Write-Host "Detail lengkap (termasuk status per item) telah diekspor ke CSV:" -ForegroundColor Green
Write-Host "$csvReportPath" -ForegroundColor White
# ==========================================
# 5. TAMPILKAN HASIL RINGKASAN
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "     RINGKASAN ANALISIS MONITOR TAKA      " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " 1. Total Baris (List Item) : $totalRows row" -ForegroundColor White
Write-Host " 2. Status Kelengkapan Field:" -ForegroundColor White
Write-Host "    - Keduanya Terisi        : $bothFilled row" -ForegroundColor Green
Write-Host "    - Keduanya KOSONG        : $bothEmpty row" -ForegroundColor Red
Write-Host "    - Cuma IDMK0 Terisi      : $onlyIDMK0 row" -ForegroundColor Yellow
Write-Host "    - Cuma ID MK Terisi      : $onlyIDMK row" -ForegroundColor Yellow
Write-Host "------------------------------------------" -ForegroundColor Cyan
Write-Host " 3. Total Unique Value:" -ForegroundColor White
Write-Host "    - Unique Value IDMK0     : $($uniqueIDMK0.Count) nilai" -ForegroundColor Green
Write-Host "    - Unique Value ID MK     : $($uniqueIDMK.Count) nilai" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

