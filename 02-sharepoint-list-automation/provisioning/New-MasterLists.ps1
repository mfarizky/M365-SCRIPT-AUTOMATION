# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

$ClientId = $Config.InteractiveClientId
$Url = $Config.HubSiteUrl

Connect-PnPOnline -Url $Url -ClientId $ClientId

# =========================================================
# 1. DB_LAYANAN (Master Kategori Layanan)
# =========================================================
New-PnPList -Title "DB LAYANAN" -Template GenericList
Add-PnPField -List "DB LAYANAN" -DisplayName "LAYANAN_ID" -InternalName "LAYANAN_ID" -Type Text
Add-PnPField -List "DB LAYANAN" -DisplayName "LAYANAN_NAME" -InternalName "LAYANAN_NAME" -Type Text
Add-PnPField -List "DB LAYANAN" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean

# =========================================================
# 2. DB_TAKA (Kamus Tahapan & SLA)
# =========================================================
New-PnPList -Title "DB TAKA" -Template GenericList
Add-PnPField -List "DB TAKA" -DisplayName "TAKA_ID" -InternalName "TAKA_ID" -Type Text
Add-PnPField -List "DB TAKA" -DisplayName "TAKA_NAME" -InternalName "TAKA_NAME" -Type Text
Add-PnPField -List "DB TAKA" -DisplayName "DIVISION" -InternalName "DEPARTMENT" -Type Choice -Choices "FINANCE", "ASSIST", "SERVICE", "TECHNOLOGY", "LEGAL"


# =========================================================
# 3. DB_LAYANAN_TAKA_MAP (Peta Jalan / Urutan Taka)
# =========================================================
New-PnPList -Title "DB LAYANAN TAKA" -Template GenericList
Add-PnPField -List "DB LAYANAN TAKA" -DisplayName "LAYANAN_ID" -InternalName "LAYANAN_ID" -Type Text
Add-PnPField -List "DB LAYANAN TAKA" -DisplayName "TAKA_ID" -InternalName "TAKA_ID" -Type Text
Add-PnPField -List "DB LAYANAN TAKA" -DisplayName "URUTAN" -InternalName "URUTAN" -Type Number

# =========================================================
# 4. DB_LOG_DURASI (Mesin Analytics / History)
# =========================================================
New-PnPList -Title "LOG MONITOR KLIEN" -Template GenericList
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "WORKFLOW_ID" -InternalName "WORKFLOW_ID" -Type Text
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "TAKA_NAME" -InternalName "TAKA_NAME" -Type Text
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "PIC" -InternalName "PIC" -Type Text
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "WAKTU_MULAI" -InternalName "WAKTU_MULAI" -Type DateTime
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "WAKTU_SELESAI" -InternalName "WAKTU_SELESAI" -Type DateTime
Add-PnPField -List "LOG MONITOR KLIEN" -DisplayName "STATUS_TAKA" -InternalName "STATUS_TAKA" -Type Text

# Setup Navigasi Sidebar
$dbHeader = Add-PnPNavigationNode -Title "DIGITAL OPERATIONAL CONTROL" -Location QuickLaunch -External -Url $Url
$lists = @("DB LAYANAN", "DB TAKA", "DB LAYANAN TAKA MAP", "LOG MONITOR KLIEN")
foreach ($list in $lists) {
    Add-PnPNavigationNode -Title $list -Location QuickLaunch -Url "$Url/Lists/$list" -Parent $dbHeader
}