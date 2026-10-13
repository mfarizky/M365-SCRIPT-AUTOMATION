# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==========================================
# AUDIT VIEW TIME ZONES (READ-ONLY) - VIA MS GRAPH
# ==========================================

# ==========================================
# KONFIGURASI KREDENSIAL & SERTIFIKAT
# ==========================================
$clientId      = $Config.AppClientId
$tenant        = $Config.TenantDomain
$pfxPath       = $Config.PfxPath
$pfxPassword   = $Config.PfxPassword
$exportFolder  = $Config.ReportFolder

if (-not (Test-Path -Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
}

$timestamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$auditCsvPath  = Join-Path -Path $exportFolder -ChildPath "ViewTimeZonesCheck_$timestamp.csv"

# ==========================================
# 1. CONNECT VIA CERTIFICATE
# ==========================================
Write-Host "Mendapatkan Access Token dari Azure AD..." -ForegroundColor Cyan

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, $pfxPassword)

if (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue) {
    Connect-MgGraph -ClientId $clientId -TenantId $tenant -Certificate $cert | Out-Null
}
else {
    Write-Host "Modul Microsoft.Graph belum terinstall/loaded." -ForegroundColor Red
    return
}

# ==========================================
# 2. AMBIL DATA USER & CHECK VIEW TIME ZONES
# ==========================================
Write-Host "Mengambil data User & View Time Zones dari Microsoft Graph..." -ForegroundColor Cyan

$allUsers = Get-MgUser -All -Property "id", "displayName", "userPrincipalName", "accountEnabled"

$results = @()
$total   = $allUsers.Count
$count   = 0

foreach ($user in $allUsers) {
    $count++
    if ($user.AccountEnabled -eq $false) { continue }

    Write-Progress -Activity "Checking Calendar View Time Zones" -Status ("[$count/$total] " + $user.UserPrincipalName) -PercentComplete (($count / $total) * 100)

    $primaryTZ     = "N/A"
    $viewTimeZones = "None"
    $tzCount       = 0
    $status        = "Success"

    try {
        # Endpoint Graph untuk membaca regional & language / calendar settings
        $uri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/mailboxSettings/regionalAndLanguageSettings"
        $regionalSettings = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        if ($regionalSettings) {
            # Time zone utama akun
            if ($regionalSettings.defaultTimeZone) {
                $primaryTZ = $regionalSettings.defaultTimeZone.name
            }

            # Daftar View Time Zones tambahan di kalender
            if ($regionalSettings.timeZones) {
                $tzNames       = $regionalSettings.timeZones | ForEach-Object { $_.name }
                $viewTimeZones = $tzNames -join " | "
                $tzCount       = $regionalSettings.timeZones.Count
            }
        }
    }
    catch {
        $status = "No Exchange Mailbox / Access Denied"
    }

    $results += [PSCustomObject]@{
        DisplayName        = $user.DisplayName
        UPN                = $user.UserPrincipalName
        UserId             = $user.Id
        PrimaryTimeZone    = $primaryTZ
        ViewTimeZones      = $viewTimeZones
        AdditionalTZCount  = $tzCount
        Status             = $status
    }
}

Write-Progress -Activity "Checking Calendar View Time Zones" -Completed

# ==========================================
# 3. EXPORT AUDIT & SUMMARY
# ==========================================
$results | Export-Csv -Path $auditCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "`nHasil check tersimpan di: $auditCsvPath" -ForegroundColor Green

Write-Host ""
Write-Host "===== SUMMARY CHECK =====" -ForegroundColor Yellow
Write-Host "Total user aktif dicek        : $($results.Count)"
Write-Host "User dengan View TZ Tambahan  : $(($results | Where-Object { $_.AdditionalTZCount -gt 0 }).Count)"
Write-Host "User tanpa View TZ Tambahan   : $(($results | Where-Object { $_.AdditionalTZCount -eq 0 -and $_.Status -eq 'Success' }).Count)"
Write-Host "Tanpa Mailbox / Access Error  : $(($results | Where-Object { $_.Status -ne 'Success' }).Count)"
Write-Host "========================="

Write-Host "`nPreview hasil check:" -ForegroundColor Cyan
$results | Select-Object DisplayName, UPN, PrimaryTimeZone, ViewTimeZones, Status | Format-Table -AutoSize

Write-Host "`nProses check selesai." -ForegroundColor Cyan