Write-Host "--- TAHAP PRA-SCRIPT: INSTALASI MODUL ---" -ForegroundColor Cyan

# Set izin eksekusi script agar tidak diblokir Windows
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# Update Kurir (NuGet)
Write-Host "Menyiapkan Package Provider..." -ForegroundColor Yellow
#Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null

# Install PnP PowerShell (Hanya jika belum ada)
if (!(Get-Module -ListAvailable PnP.PowerShell)) {
    Write-Host "Menginstall PnP.PowerShell (Mohon tunggu, laptop mungkin agak berat sebentar)..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -AllowClobber -Force -Confirm:$false
    Write-Host "Instalasi Modul Berhasil!" -ForegroundColor Green
} else {
    Write-Host "Modul sudah terinstall, tidak perlu install ulang." -ForegroundColor Green
}

Write-Host "--- PRA-SCRIPT SELESAI. SIAP MENJALANKAN SCRIPT UTAMA ---" -ForegroundColor Cyan