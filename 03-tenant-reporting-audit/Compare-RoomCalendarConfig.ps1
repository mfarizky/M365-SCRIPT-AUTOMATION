# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# 1. Pastikan modul ExchangeOnlineManagement sudah terpasang dan terhubung
if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
}

# Hubungkan ke Exchange Online (jika belum terhubung)
try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
} catch {
    Write-Host "Menghubungkan ke Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline
}

# 2. Definisikan Identity (diatur di config.local.ps1)
$UserIdentity = $Config.CalendarCheckUser   # UPN user yang mengalami masalah booking
$RoomIdentity = $Config.CalendarCheckRoom   # Nama / UPN room mailbox

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "   MEMERIKSA WORKING HOURS & TIMEZONE (USER vs ROOM)    " -ForegroundColor Cyan
Write-Host "========================================================`n" -ForegroundColor Cyan

# 3. Ambil data konfigurasi kalender keduanya
try {
    $CalendarConfigs = Get-MailboxCalendarConfiguration -Identity $UserIdentity, $RoomIdentity -ErrorAction Stop

    # Tampilkan hasil perbandingan dalam format tabel
    $CalendarConfigs | Select-Object @{N="Mailbox / Identity"; E={$_.Identity}},
                                     @{N="Start Time"; E={$_.WorkingHoursStartTime}},
                                     @{N="End Time"; E={$_.WorkingHoursEndTime}},
                                     @{N="Working Days"; E={$_.WorkingDays}},
                                     @{N="Working Time Zone"; E={$_.WorkingHoursTimeZone}} | Format-Table -AutoSize

    Write-Host "[HINT] Analisis Penyebab Unavailable:" -ForegroundColor Yellow
    Write-Host "1. Pastikan 'Working Time Zone' User & Room bernilai sama (misal: Singapore Standard Time / SE Asia Standard Time)."
    Write-Host "2. Jika slot meeting berada di luar rentang 'Start Time' - 'End Time' atau di luar 'Working Days' Sam, status ruangan bisa terbaca Unavailable."
} catch {
    Write-Error "Gagal mengambil data. Pastikan UPN/Email user sudah benar dan Anda memiliki hak akses Admin."
}