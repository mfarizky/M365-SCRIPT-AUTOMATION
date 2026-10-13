# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

$siteUrl        = $Config.HubSiteUrl

# Konfigurasi App Registration & Sertifikat Spesifik
$clientId       = $Config.AppClientId
$tenant         = $Config.TenantDomain
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

$ListName = "MONITOR KLIEN"
$OutputCsv = (Join-Path $Config.ReportFolder "MonitorKlien_Report.csv")
$itemId    = 88

# ==========================================
# 2. Koneksi ke SharePoint Online
# ==========================================
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Remove-Item $OutputCsv -ErrorAction SilentlyContinue

# ==== Mapping display name -> internal name ====
$fields = Get-PnPField -List $ListName
function Get-InternalName($displayName) {
    $f = $fields | Where-Object { $_.Title -eq $displayName }
    if (-not $f) { Write-Warning "Kolom '$displayName' TIDAK DITEMUKAN!" }
    return $f.InternalName
}
$colKlien    = Get-InternalName "KLIEN"
$colStatus   = Get-InternalName "STATUS"
$colLayanan  = Get-InternalName "LAYANAN"
$colTaka     = Get-InternalName "TAKA"
$colProgress = Get-InternalName "PROGRESS (%)"
$colCO       = Get-InternalName "CO"
$colPIC      = Get-InternalName "PIC"
$colDitugaskan = Get-InternalName "DITUGASKAN"
$colInstansi = Get-InternalName "INSTANSI"

# ==== Format tanggal ke WITA ====
function Format-WitaDate($isoDate) {
    if (-not $isoDate) { return "" }
    $utcDate = [DateTime]::Parse($isoDate, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
    $witaDate = $utcDate.AddHours(8)
    return $witaDate.ToString("dd/MM/yyyy HH:mm") + " WITA"
}

# ==== Format tanggal saja (tanpa jam), untuk kolom seperti DITUGASKAN ====
function Format-DateOnly($val) {
    if (-not $val) { return "" }
    try {
        $d = [DateTime]$val
        return $d.ToString("dd/MM/yyyy")
    } catch {
        return $val
    }
}

# ==== Resolve mention: index-based, property "name" langsung dari REST ====
function Resolve-CommentMentions($comment) {
    $commentText = [System.Net.WebUtility]::HtmlDecode($comment.text)
    if ($comment.mentions -and $comment.mentions.Count -gt 0) {
        for ($m = 0; $m -lt $comment.mentions.Count; $m++) {
            $mName = $comment.mentions[$m].name
            if (-not $mName) { $mName = "unknown-user" }
            $commentText = $commentText -replace "@mention\{$m\}", "@$mName"
        }
    }
    return $commentText
}

# ==== Ambil comments via REST, dengan pagination ====
function Get-ListItemCommentsRest($listName, $id) {
    $allComments = @()
    $restUrl = "$siteUrl/_api/web/lists/getbytitle('$listName')/items($id)/comments?`$top=100"

    while ($restUrl) {
        $resp = Invoke-PnPSPRestMethod -Url $restUrl -Method Get
        if ($resp.value) { $allComments += $resp.value }

        if ($resp.nextHref) {
            $restUrl = $resp.nextHref
            if ($restUrl -notmatch '^https?://') { $restUrl = "$siteUrl$restUrl" }
        } elseif ($resp.'odata.nextLink') {
            $restUrl = $resp.'odata.nextLink'
        } else {
            $restUrl = $null
        }
    }
    return $allComments
}

# ==== Retry export ====
function Export-RowsWithRetry($rows, $path, $append) {
    $maxRetry = 5
    $attempt = 0
    while ($attempt -lt $maxRetry) {
        try {
            $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8 -Append:$append
            return
        } catch {
            $attempt++
            Write-Warning "File lagi terpakai, coba lagi ($attempt/$maxRetry)..."
            Start-Sleep -Seconds 2
        }
    }
    Write-Error "GAGAL menulis untuk item ID $($rows[0].ID) setelah $maxRetry percobaan."
}

# ==== Filter STATUS sesuai view ACTIVE ====
$camlQuery = @"
<View>
  <Query>
    <Where>
      <In>
        <FieldRef Name='$colStatus' />
        <Values>
          <Value Type='Text'>BELUM MULAI</Value>
          <Value Type='Text'>BERLANGSUNG</Value>
          <Value Type='Text'>PRIORITY</Value>
        </Values>
      </In>
    </Where>
  </Query>
</View>
"@

$items = Get-PnPListItem -List $ListName -Query $camlQuery
Write-Host "Total item yang match filter: $($items.Count)"

$isFirstWrite = $true

foreach ($item in $items) {
    $currentId = $item["ID"]
    $itemComments = Get-ListItemCommentsRest -listName $ListName -id $currentId

    $baseRow = @{
        ID       = $currentId
        KLIEN    = $item[$colKlien].LookupValue
        DITUGASKAN = Format-DateOnly $item[$colDitugaskan]
        STATUS   = $item[$colStatus]
        LAYANAN  = $item[$colLayanan]
        TAKA     = $item[$colTaka]
        PROGRESS = $item[$colProgress]
        INSTANSI = $item[$colInstansi]
        CO       = $item[$colCO].LookupValue
        PIC      = $item[$colPIC].LookupValue
    }

    $rowsForItem = if ($itemComments -and $itemComments.Count -gt 0) {
        for ($i = 0; $i -lt $itemComments.Count; $i++) {
            $cleanText = Resolve-CommentMentions $itemComments[$i]
            $authorName = $itemComments[$i].author.name
            $createdFriendly = Format-WitaDate $itemComments[$i].createdDate
            [PSCustomObject]($baseRow + @{ AUTHOR = $authorName; "CREATED DATE" = $createdFriendly; KOMENTAR = $cleanText }) | Select-Object ID,KLIEN,DITUGASKAN,STATUS,LAYANAN,TAKA,PROGRESS,INSTANSI,CO,PIC,AUTHOR,"CREATED DATE",KOMENTAR
        }
    } else {
        [PSCustomObject]($baseRow + @{ AUTHOR = ""; "CREATED DATE" = ""; KOMENTAR = "" }) | Select-Object ID,KLIEN,DITUGASKAN,STATUS,LAYANAN,TAKA,PROGRESS,INSTANSI,CO,PIC,AUTHOR,"CREATED DATE",KOMENTAR
    }

    Export-RowsWithRetry -rows $rowsForItem -path $OutputCsv -append (-not $isFirstWrite)
    $isFirstWrite = $false

    Write-Host "Item $currentId ($($item[$colStatus])) selesai - $($itemComments.Count) komentar"
}

Write-Host "SELESAI. File tersimpan di: $OutputCsv"
