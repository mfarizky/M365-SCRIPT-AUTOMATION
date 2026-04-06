# ============================================================
# Import-ICSToM365.ps1
# Import ICS holiday calendar into Microsoft 365 via Graph API
#
# Triggered by  : user (manual run)
# Requires      : PowerShell 7+ | Microsoft.Graph module
# Usage         : .\Import-ICSToM365.ps1
#                 .\Import-ICSToM365.ps1 -EnvFile ".\custom.env"
# ============================================================

param(
    [string]$EnvFile = ".\.env"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ==============================
# DEPENDENCIES CHECK
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " DEPENDENCY CHECK" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

# Minimum required versions
$requiredModules = @(
    @{ Name = "Microsoft.Graph.Authentication"; MinVersion = "2.0.0" }
    @{ Name = "Microsoft.Graph.Calendar";       MinVersion = "2.0.0" }
)

foreach ($mod in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $mod.Name |
                 Sort-Object Version -Descending |
                 Select-Object -First 1

    if ($null -eq $installed) {
        Write-Host "[ MISSING ] $($mod.Name) — installing..." -ForegroundColor Yellow
        Install-Module -Name $mod.Name -MinimumVersion $mod.MinVersion `
                       -Scope CurrentUser -Force -AllowClobber
        Write-Host "[   OK    ] $($mod.Name) installed." -ForegroundColor Green
    }
    elseif ($installed.Version -lt [version]$mod.MinVersion) {
        Write-Host "[ OUTDATED] $($mod.Name) v$($installed.Version) — upgrading to $($mod.MinVersion)+..." -ForegroundColor Yellow
        Update-Module -Name $mod.Name -Force
        Write-Host "[   OK    ] $($mod.Name) updated." -ForegroundColor Green
    }
    else {
        Write-Host "[   OK    ] $($mod.Name) v$($installed.Version)" -ForegroundColor Green
    }

    Import-Module $mod.Name -MinimumVersion $mod.MinVersion -ErrorAction Stop
}

# ==============================
# LOAD .ENV
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " LOADING CONFIG" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

if (-not (Test-Path $EnvFile)) {
    Write-Host "ERROR: .env file not found at '$EnvFile'" -ForegroundColor Red
    Write-Host "       Copy .env.example to .env and fill in your values." -ForegroundColor Yellow
    exit 1
}

# Parse .env — skip comments and blank lines
Get-Content $EnvFile | ForEach-Object {
    $line = $_.Trim()
    if ($line -and $line -notmatch "^#") {
        $parts = $line -split "=", 2
        if ($parts.Count -eq 2) {
            $name  = $parts[0].Trim()
            $value = $parts[1].Trim().Trim('"').Trim("'")
            Set-Variable -Name $name -Value $value -Scope Script
        }
    }
}

# Validate required variables
$required = @("TenantId","ClientId","ClientSecret","TargetMailbox","ICSUrl","TimeZone")
$missing  = $required | Where-Object { -not (Get-Variable -Name $_ -Scope Script -ErrorAction SilentlyContinue) }

if ($missing.Count -gt 0) {
    Write-Host "ERROR: Missing required variables in .env:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "       - $_" -ForegroundColor Red }
    exit 1
}

Write-Host "Config loaded from   : $EnvFile" -ForegroundColor Green
Write-Host "Target Mailbox       : $TargetMailbox" -ForegroundColor Green
Write-Host "Timezone             : $TimeZone" -ForegroundColor Green

# ==============================
# CONNECT TO GRAPH
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " CONNECTING TO MICROSOFT GRAPH" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

$SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$Credential   = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)

Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential
Write-Host "Connected successfully!" -ForegroundColor Green

# ==============================
# DOWNLOAD ICS
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " DOWNLOADING ICS" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

Write-Host "URL: $ICSUrl" -ForegroundColor Yellow

$response = Invoke-WebRequest -Uri $ICSUrl -UseBasicParsing
$content  = $response.Content

Write-Host "Status Code    : $($response.StatusCode)" -ForegroundColor Green
Write-Host "Content Length : $($content.Length) characters" -ForegroundColor Green

# ==============================
# FIX ICS FOLDED LINES & PARSE
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " PARSING ICS" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

# Unfold: line ending with CRLF+space = continuation
$content = $content -replace "`r`n ", ""
$lines   = $content -split "`r`n"

$rawEventCount = ($lines | Where-Object { $_ -eq "BEGIN:VEVENT" }).Count
Write-Host "VEVENT blocks found  : $rawEventCount" -ForegroundColor Yellow

$events       = @()
$currentEvent = @{}
$inEvent      = $false

foreach ($line in $lines) {
    $line = $line.Trim()

    if ($line -eq "BEGIN:VEVENT") {
        $currentEvent = @{}
        $inEvent = $true
        continue
    }

    if ($line -eq "END:VEVENT") {
        $events += $currentEvent
        $inEvent = $false
        continue
    }

    if ($inEvent -and $line -match ":") {
        $parts = $line -split ":", 2
        $key   = $parts[0].Trim()
        $value = $parts[1].Trim()

        if ($key -match "^DTSTART") { $key = "DTSTART" }
        if ($key -match "^DTEND")   { $key = "DTEND"   }
        if ($key -match "^SUMMARY") { $key = "SUMMARY" }

        $currentEvent[$key] = $value
    }
}

Write-Host "Events parsed        : $($events.Count)" -ForegroundColor Yellow

if ($events.Count -eq 0) {
    Write-Host "ERROR: No events parsed. Check ICS format or URL." -ForegroundColor Red
    exit 1
}

# ==============================
# IMPORT TO M365 CALENDAR
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " IMPORTING TO M365 CALENDAR" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

$addedCount   = 0
$skippedCount = 0
$errorCount   = 0

foreach ($e in $events) {
    # Skip incomplete events
    if (-not $e.ContainsKey("DTSTART") -or
        -not $e.ContainsKey("DTEND")   -or
        -not $e.ContainsKey("SUMMARY")) {
        Write-Host "  SKIP (incomplete): missing DTSTART/DTEND/SUMMARY" -ForegroundColor DarkGray
        $skippedCount++
        continue
    }

    $startRaw = $e["DTSTART"].Trim()
    $endRaw   = $e["DTEND"].Trim()
    $subject  = $e["SUMMARY"].Trim()

    try {
        $start = [DateTime]::ParseExact($startRaw.Substring(0,8), "yyyyMMdd", $null)
        $end   = [DateTime]::ParseExact($endRaw.Substring(0,8),   "yyyyMMdd", $null)
    }
    catch {
        Write-Host "  SKIP (bad date): $subject — $startRaw / $endRaw" -ForegroundColor DarkYellow
        $skippedCount++
        continue
    }

    $eventBody = @{
        subject = "[HOLIDAY] $subject"
        start   = @{
            dateTime = $start.ToString("yyyy-MM-ddT00:00:00")
            timeZone = $TimeZone
        }
        end     = @{
            dateTime = $end.ToString("yyyy-MM-ddT00:00:00")
            timeZone = $TimeZone
        }
        isAllDay = $true
        showAs   = "oof"
    }

    try {
        New-MgUserEvent -UserId $TargetMailbox -BodyParameter $eventBody | Out-Null
        Write-Host "  ADDED : $subject ($($start.ToString("yyyy-MM-dd")))" -ForegroundColor Green
        $addedCount++
    }
    catch {
        Write-Host "  ERROR : $subject — $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}

# ==============================
# SUMMARY
# ==============================
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host " IMPORT FINISHED" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host "Mailbox Target : $TargetMailbox"
Write-Host "Added          : $addedCount"    -ForegroundColor Green
Write-Host "Skipped        : $skippedCount"  -ForegroundColor Yellow
Write-Host "Errors         : $errorCount"    -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host "==============================`n" -ForegroundColor Cyan

Disconnect-MgGraph | Out-Null
