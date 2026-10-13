# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# --- CONFIGURATION ---
$clientId = $Config.InteractiveClientId
$tenantDomain = $Config.SharePointHost

# 1. Login
Connect-PnPOnline -Url "https://$tenantDomain" -Interactive -ClientId $clientId

# 2. Get & Filter Sites (Tetap pakai logic Anda)
$searchQuery = "contentclass:STS_Site OR contentclass:STS_Web"
$allSites = Submit-PnPSearchQuery -Query $searchQuery -All -SelectProperties "Title", "Path", "SiteId"

$filteredSites = $allSites.ResultRows | Where-Object {
    $url = $_["Path"].ToLower()
    $title = $_["Title"].ToLower()
    -not ($url -like "*content*") -and -not ($url -like "*app*") -and
    -not ($url -like "*vivahome*") -and -not ($title -like "*gki*") -and
    -not ($title -like "*portal*") -and -not ($url -like "*gki*") -and
    -not ($url -like "*ganawiki*") -and -not ($url -like "*portal*") -and
    -not ($title -like "*fill gaps site*") -and -not ($title -like "*communication*") -and
    -not ($title -like "*gana*")
}

$userEmail = (Get-PnPContext).Web.CurrentUser.Email
if (!$userEmail) { $userEmail = whoami /upn }

# 3. Eksekusi Sync - Meniru persis perilaku tombol Sync di UI
foreach ($site in $filteredSites) {
    try {
        $siteUrl = $site["Path"]
        $siteTitle = $site["Title"] 
        
        $conn = Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId -ReturnConnection
        
        # Identifikasi Library (PORTAL atau Documents)
        $list = Get-PnPList -Identity "PORTAL" -Connection $conn -ErrorAction SilentlyContinue
        $lName = "PORTAL"
        if ($null -eq $list) {
            $list = Get-PnPList -Identity "Documents" -Connection $conn -ErrorAction SilentlyContinue
            $lName = "Documents"
        }
        
        $web = Get-PnPWeb -Connection $conn -Includes Id, Title

        if ($list) {
            $sId = $site["SiteId"].ToString().Trim("{}")
            $wId = $web.Id.ToString().Trim("{}")
            $lId = $list.Id.ToString().Trim("{}")
            
            # Format display: "<NAMA KLIEN> - PORTAL"
            $fullDisplayName = "$siteTitle - $lName"
            
            $encodedUrl = [uri]::EscapeDataString($siteUrl)
            $encodedName = [uri]::EscapeDataString($fullDisplayName)
            
            # Tambahan parameter internal SharePoint (isSyncRequirementMet & isLegacyService)
            # Ini memaksa OneDrive Client untuk memproses metadata secara penuh
            $odProtocol = "odopen://sync/?siteId=$sId&webId=$wId&listId=$lId&userEmail=$userEmail&webUrl=$encodedUrl&siteName=$encodedName&listName=$encodedName&isSyncRequirementMet=1&isLegacyService=0"
            
            Write-Host "Memicu Sync UI-Style untuk: $fullDisplayName" -ForegroundColor Green
            Start-Process $odProtocol
            
            # Kasih jeda 10 detik. Jangan dipercepat. 
            # Tombol Sync di UI SharePoint juga lambat prosesnya di background.
            Start-Sleep -Seconds 10
        }
    } catch {
        Write-Host "Gagal di site: $siteTitle" -ForegroundColor Red
    }
}