# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ============================================================
# MONITOR TAKA VALIDATOR
# READ-ONLY SHAREPOINT VALIDATION
# ============================================================
#
# TUJUAN:
# Membandingkan:
#
# DATABASE WORKFLOW
#        VS
# MONITOR TAKA
#
# BERDASARKAN:
# - ID LAYANAN
# - IDMK / ID MK
#
# SCRIPT INI 100% READ-ONLY
#
# TIDAK ADA:
# - CREATE
# - UPDATE
# - DELETE
# ============================================================


# ============================================================
# 1. CONFIGURATION
# ============================================================

$siteUrl = $Config.HubSiteUrl

$tenant = $Config.TenantDomain

$clientId = $Config.AppClientId

$pfxPath = $Config.PfxPath

$masterListName = "DATABASE WORKFLOW"

$monitorListName = "MONITOR TAKA"



# ============================================================
# 2. PFX PASSWORD
# ============================================================

# Password PFX diambil dari config; jika kosong, diminta secara interaktif.
$pfxPassword = $Config.PfxPassword
if (-not $pfxPassword) { $pfxPassword = Read-Host "Masukkan PFX Password" -AsSecureString }


# ============================================================
# 3. HEADER
# ============================================================

Clear-Host

Write-Host ""
Write-Host ("=" * 70)
Write-Host "MONITOR TAKA VALIDATOR"
Write-Host ("=" * 70)
Write-Host ""
Write-Host "READ-ONLY SHAREPOINT VALIDATION"
Write-Host "No data will be created, updated, or deleted."
Write-Host ""


# ============================================================
# 4. CHECK CERTIFICATE
# ============================================================

Write-Host "CHECK CERTIFICATE"
Write-Host ("-" * 70)

Write-Host "CERT PATH : $pfxPath"

if (-not (Test-Path -LiteralPath $pfxPath)) {
    throw "Certificate tidak ditemukan: $pfxPath"
}

Write-Host "EXISTS    : True"
Write-Host ""


# ============================================================
# 5. CONNECT SHAREPOINT
# ============================================================

Write-Host "CONNECTING TO SHAREPOINT..."
Write-Host ("-" * 70)

try {

    Connect-PnPOnline `
        -Url $siteUrl `
        -ClientId $clientId `
        -Tenant $tenant `
        -CertificatePath $pfxPath `
        -CertificatePassword $pfxPassword `
        -ErrorAction Stop

}
catch {

    throw "Gagal connect ke SharePoint. Error: $($_.Exception.Message)"
}


# ============================================================
# 6. VERIFY CONNECTION
# ============================================================

try {

    $web = Get-PnPWeb -ErrorAction Stop

}
catch {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    throw "Authentication berhasil tetapi gagal membaca SharePoint Web. Error: $($_.Exception.Message)"
}


Write-Host ""
Write-Host "SHAREPOINT CONNECTION SUCCESS"
Write-Host "Site Title : $($web.Title)"
Write-Host "Site URL   : $siteUrl"
Write-Host ""


# ============================================================
# 7. CHECK LISTS
# ============================================================

Write-Host ("=" * 70)
Write-Host "CHECK SHAREPOINT LISTS"
Write-Host ("=" * 70)

try {

    $masterList = Get-PnPList `
        -Identity $masterListName `
        -ErrorAction Stop

    Write-Host "MASTER LIST : $($masterList.Title)"

}
catch {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    throw "List '$masterListName' tidak ditemukan."
}


try {

    $monitorList = Get-PnPList `
        -Identity $monitorListName `
        -ErrorAction Stop

    Write-Host "MONITOR LIST: $($monitorList.Title)"

}
catch {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    throw "List '$monitorListName' tidak ditemukan."
}


# ============================================================
# 8. INPUT ID LAYANAN
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "INPUT"
Write-Host ("=" * 70)

$idLayananInput = Read-Host "Masukkan ID LAYANAN"

if ([string]::IsNullOrWhiteSpace($idLayananInput)) {
    throw "ID LAYANAN tidak boleh kosong."
}

$idLayanan = 0

if (-not [int]::TryParse($idLayananInput, [ref]$idLayanan)) {
    throw "ID LAYANAN harus berupa angka."
}


# ============================================================
# 9. READ DATABASE WORKFLOW
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "READ DATABASE WORKFLOW"
Write-Host ("=" * 70)

$masterItems = Get-PnPListItem `
    -List $masterListName `
    -PageSize 2000 `
    -ErrorAction Stop

Write-Host "TOTAL MASTER ITEMS : $($masterItems.Count)"


# ============================================================
# 10. FIND MASTER ITEM
# ============================================================

$masterItem = $masterItems |
    Where-Object {
        $_.Id -eq $idLayanan
    } |
    Select-Object -First 1

if ($null -eq $masterItem) {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    throw "ID LAYANAN $idLayanan tidak ditemukan di DATABASE WORKFLOW."
}


# ============================================================
# 11. GET LAYANAN
# ============================================================

$layanan = [string]$masterItem["field_1"]

if ($null -eq $layanan) {
    $layanan = ""
}

$layanan = $layanan.Trim()


# ============================================================
# 12. GET MASTER TAKA
# ============================================================

$masterTaka = @()

for ($number = 1; $number -le 70; $number++) {

    # --------------------------------------------------------
    # TAKA 1 - TAKA 11
    # --------------------------------------------------------

    if ($number -le 11) {

        $internalName = "field_$($number + 1)"
    }

    # --------------------------------------------------------
    # TAKA 12 - TAKA 65
    # --------------------------------------------------------

    elseif ($number -le 65) {

        $internalName = "TAKA$number"
    }

    # --------------------------------------------------------
    # TAKA 66 - TAKA 70
    # --------------------------------------------------------

    else {

        $internalName = "TAKA_x0020_$number"
    }


    $value = $masterItem[$internalName]

    if ($null -ne $value) {

        $value = ([string]$value).Trim()

        if (-not [string]::IsNullOrWhiteSpace($value)) {

            $masterTaka += $value
        }
    }
}


# ============================================================
# 13. DISPLAY MASTER WORKFLOW
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "MASTER WORKFLOW"
Write-Host ("=" * 70)

Write-Host "ID LAYANAN : $idLayanan"
Write-Host "LAYANAN    : $layanan"
Write-Host ""

for ($i = 0; $i -lt $masterTaka.Count; $i++) {

    $displayNo = $i + 1

    Write-Host ("{0,3}  {1}" -f $displayNo, $masterTaka[$i])
}

Write-Host ""
Write-Host "TOTAL MASTER: $($masterTaka.Count)"


# ============================================================
# 14. INPUT IDMK
# ============================================================

Write-Host ""

$idmk = Read-Host "Masukkan IDMK / ID MK MONITOR KLIEN"

if ([string]::IsNullOrWhiteSpace($idmk)) {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    throw "IDMK tidak boleh kosong."
}


# ============================================================
# 15. READ MONITOR TAKA
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "READ MONITOR TAKA"
Write-Host ("=" * 70)

$monitorItems = Get-PnPListItem `
    -List $monitorListName `
    -PageSize 2000 `
    -ErrorAction Stop

Write-Host "TOTAL MONITOR ITEMS : $($monitorItems.Count)"


# ============================================================
# 16. FILTER BY ID MK LOOKUP
# ============================================================

$monitorForIdMk = @()

foreach ($item in $monitorItems) {

    $rawIdMk = $item["ID_x0020_MK"]

    if ($null -eq $rawIdMk) {
        continue
    }


    $lookupValue = ""

    if ($rawIdMk -is [Microsoft.SharePoint.Client.FieldLookupValue]) {

        $lookupValue = [string]$rawIdMk.LookupValue
    }

    elseif ($rawIdMk -is [System.Array]) {

        foreach ($lookup in $rawIdMk) {

            if ($lookup -is [Microsoft.SharePoint.Client.FieldLookupValue]) {

                $lookupValue = [string]$lookup.LookupValue
                break
            }
        }
    }

    else {

        $lookupValue = [string]$rawIdMk
    }


    $lookupValue = $lookupValue.Trim()


    if ($lookupValue -eq $idmk) {

        $monitorForIdMk += $item
    }
}


# ============================================================
# 17. DISPLAY MONITOR TAKA
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "MONITOR TAKA FOR IDMK"
Write-Host ("=" * 70)

Write-Host "IDMK               : $idmk"
Write-Host "TOTAL MONITOR TAKA : $($monitorForIdMk.Count)"
Write-Host ""


# ============================================================
# 18. EXTRACT MONITOR DATA
# ============================================================

$monitorTaka = @()

foreach ($item in $monitorForIdMk) {

    $rawNo = $item["NO_x002e_"]

    $rawTaka = $item["TAKA"]

    $noText = ""

    if ($null -ne $rawNo) {

        $noText = ([string]$rawNo).Trim()
    }


    $takaText = ""

    if ($null -ne $rawTaka) {

        $takaText = ([string]$rawTaka).Trim()
    }


    $numericNo = $null

    $parsedNo = 0

    if ([int]::TryParse($noText, [ref]$parsedNo)) {

        $numericNo = $parsedNo
    }


    $monitorTaka += [PSCustomObject]@{
        ItemId   = $item.Id
        NoText   = $noText
        No       = $numericNo
        Taka     = $takaText
    }
}


# ============================================================
# 19. SORT MONITOR TAKA
# ============================================================

$monitorTakaSorted = $monitorTaka |
    Sort-Object `
        @{ Expression = {
            if ($null -eq $_.No) {
                [int]::MaxValue
            }
            else {
                $_.No
            }
        } },
        ItemId


# ============================================================
# 20. DISPLAY MONITOR DATA
# ============================================================

foreach ($row in $monitorTakaSorted) {

    if ($null -eq $row.No) {

        Write-Host ("{0,3}  {1}" -f "?", $row.Taka)
    }

    else {

        Write-Host ("{0,3}  {1}" -f $row.No, $row.Taka)
    }
}


# ============================================================
# 21. VALIDATION
# ============================================================

Write-Host ""
Write-Host ("=" * 70)
Write-Host "VALIDATION"
Write-Host ("=" * 70)


# ============================================================
# COUNT
# ============================================================

$masterCount = $masterTaka.Count

$monitorCount = $monitorTaka.Count

$countValid = ($masterCount -eq $monitorCount)

if ($countValid) {

    Write-Host "COUNT          : VALID ($masterCount / $monitorCount)"
}

else {

    Write-Host "COUNT          : MISMATCH (Expected $masterCount / Actual $monitorCount)"
}


# ============================================================
# DUPLICATE NO
# ============================================================

$validNoRows = $monitorTaka |
    Where-Object {
        $null -ne $_.No
    }

$duplicateNoGroups = $validNoRows |
    Group-Object No |
    Where-Object {
        $_.Count -gt 1
    }


if ($duplicateNoGroups.Count -eq 0) {

    Write-Host "DUPLICATE NO   : NONE"
}

else {

    Write-Host "DUPLICATE NO   : YES"

    foreach ($group in $duplicateNoGroups) {

        Write-Host "                 NO $($group.Name) = $($group.Count) items"
    }
}


# ============================================================
# INVALID / BLANK NO
# ============================================================

$invalidNoRows = $monitorTaka |
    Where-Object {
        $null -eq $_.No
    }


if ($invalidNoRows.Count -eq 0) {

    Write-Host "INVALID NO     : NONE"
}

else {

    Write-Host "INVALID NO     : $($invalidNoRows.Count) item(s)"

    foreach ($row in $invalidNoRows) {

        Write-Host "                 Item ID $($row.ItemId) | NO='$($row.NoText)' | TAKA='$($row.Taka)'"
    }
}


# ============================================================
# NO SEQUENCE
# ============================================================

$sequenceValid = $true

$sequenceMessage = "N/A"

if ($validNoRows.Count -gt 0) {

    $observedStart = ($validNoRows | Measure-Object No -Minimum).Minimum

    $expectedNo = $observedStart

    foreach ($row in ($validNoRows | Sort-Object No)) {

        if ($row.No -ne $expectedNo) {

            $sequenceValid = $false
            break
        }

        $expectedNo++
    }


    if ($sequenceValid) {

        $sequenceMessage = "VALID (starts at $observedStart)"
    }

    else {

        $sequenceMessage = "INVALID"
    }
}


Write-Host "NO SEQUENCE    : $sequenceMessage"


# ============================================================
# DUPLICATE TAKA
# ============================================================

$duplicateTakaGroups = $monitorTaka |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.Taka)
    } |
    Group-Object Taka |
    Where-Object {
        $_.Count -gt 1
    }


if ($duplicateTakaGroups.Count -eq 0) {

    Write-Host "DUPLICATE TAKA  : NONE"
}

else {

    Write-Host "DUPLICATE TAKA  : YES"

    foreach ($group in $duplicateTakaGroups) {

        Write-Host "                  '$($group.Name)' = $($group.Count) items"
    }
}


# ============================================================
# ORDER VALIDATION
# ============================================================

$masterOrder = @($masterTaka)

$monitorOrder = @(
    $monitorTakaSorted |
    Where-Object {
        $null -ne $_.No
    } |
    ForEach-Object {
        $_.Taka
    }
)


$orderValid = $true

$orderMessage = "VALID"


if ($masterOrder.Count -ne $monitorOrder.Count) {

    $orderValid = $false

    $orderMessage = "MISMATCH - COUNT DIFFERENT"
}

else {

    for ($i = 0; $i -lt $masterOrder.Count; $i++) {

        if ($masterOrder[$i] -ne $monitorOrder[$i]) {

            $orderValid = $false

            $orderMessage = "MISMATCH"

            break
        }
    }
}


Write-Host "TAKA ORDER      : $orderMessage"


# ============================================================
# MISSING TAKA
# ============================================================

$missingTaka = @()

foreach ($taka in $masterOrder) {

    if ($monitorOrder -notcontains $taka) {

        $missingTaka += $taka
    }
}


if ($missingTaka.Count -eq 0) {

    Write-Host "MISSING TAKA    : NONE"
}

else {

    Write-Host "MISSING TAKA    : $($missingTaka.Count)"

    foreach ($taka in $missingTaka) {

        Write-Host "                  $taka"
    }
}


# ============================================================
# EXTRA TAKA
# ============================================================

$extraTaka = @()

foreach ($taka in $monitorOrder) {

    if ($masterOrder -notcontains $taka) {

        $extraTaka += $taka
    }
}


if ($extraTaka.Count -eq 0) {

    Write-Host "EXTRA TAKA      : NONE"
}

else {

    Write-Host "EXTRA TAKA      : $($extraTaka.Count)"

    foreach ($taka in $extraTaka) {

        Write-Host "                  $taka"
    }
}


# ============================================================
# FINAL RESULT
# ============================================================

$duplicateNoValid = ($duplicateNoGroups.Count -eq 0)

$invalidNoValid = ($invalidNoRows.Count -eq 0)

$missingValid = ($missingTaka.Count -eq 0)

$extraValid = ($extraTaka.Count -eq 0)

$finalValid = (
    $countValid -and
    $duplicateNoValid -and
    $invalidNoValid -and
    $sequenceValid -and
    $orderValid -and
    $missingValid -and
    $extraValid
)


Write-Host ""
Write-Host ("=" * 70)

if ($finalValid) {

    Write-Host "RESULT: SYNCHRONIZED"
}

else {

    Write-Host "RESULT: NOT SYNCHRONIZED"
}

Write-Host ("=" * 70)

Write-Host ""
Write-Host "READ-ONLY CONFIRMATION"
Write-Host "No SharePoint data was created, updated, or deleted."
Write-Host ""


# ============================================================
# 22. DISCONNECT
# ============================================================

Disconnect-PnPOnline -ErrorAction SilentlyContinue
