# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# 1. Koneksi ke SharePoint
$ClientId = $Config.InteractiveClientId
$Url = $Config.HubSiteUrl

# Menggunakan Interactive agar lebih mudah jika ClientId butuh auth context
Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive

# 2. Nama List
$listName = "DATABASE_WORKFLOWS"
$takaMasterList = "DB TAKA"
$layananMasterList = "DB LAYANAN"

# Ambil ID dari list master (Wajib ada sebelum membuat kolom)
$takaListId = (Get-PnPList -Identity $takaMasterList).Id
$layananListId = (Get-PnPList -Identity $layananMasterList).Id

# 3. Buat List Baru
Write-Host "Membuat List $listName..." -ForegroundColor Cyan
New-PnPList -Title $listName -Template GenericList -ErrorAction SilentlyContinue

# 4. Tambah Kolom LAYANAN (Lookup ke DB LAYANAN)
Write-Host "Menambahkan Kolom Lookup LAYANAN..." -ForegroundColor Yellow
$xmlLayanan = "<Field Type='Lookup' Name='LAYANAN' DisplayName='LAYANAN' List='{$layananListId}' ShowField='Title' Required='TRUE' />"
Add-PnPFieldFromXml -List $listName -FieldXml $xmlLayanan

# 5. Tambah Kolom TAKA 1 s/d TAKA 20 (Lookup ke DB TAKA)
for ($i = 1; $i -le 20; $i++) {
    $displayName = "TAKA $i"
    $internalName = "TAKA_$i" # Internal name tanpa spasi agar aman di Power BI
    Write-Host "Menambahkan Kolom Lookup $displayName..." -ForegroundColor Green
    
    # Gunakan XML untuk presisi menunjuk ke list DB TAKA
    $xmlTaka = "<Field Type='Lookup' Name='$internalName' DisplayName='$displayName' List='{$takaListId}' ShowField='TAKA_NAME' />"
    
    Add-PnPFieldFromXml -List $listName -FieldXml $xmlTaka
}

Write-Host "Selesai! List $listName berhasil dibuat dengan 20 kolom Lookup." -ForegroundColor White -BackgroundColor DarkGreen