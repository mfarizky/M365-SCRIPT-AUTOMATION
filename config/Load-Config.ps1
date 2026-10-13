<#
.SYNOPSIS
    Memuat konfigurasi environment (tenant, app registration, sertifikat, path) ke variabel $Config.

.DESCRIPTION
    Semua script di repo ini TIDAK menyimpan kredensial secara hard-code.
    Nilai asli disimpan di `config/config.local.ps1` (di-ignore oleh git).
    Buat file tersebut dengan menyalin `config/config.example.ps1`.
#>

$localConfig   = Join-Path $PSScriptRoot "config.local.ps1"
$exampleConfig = Join-Path $PSScriptRoot "config.example.ps1"

if (-not (Test-Path -LiteralPath $localConfig)) {
    throw "config.local.ps1 tidak ditemukan. Salin '$exampleConfig' menjadi '$localConfig' lalu isi nilainya."
}

. $localConfig

# Password PFX: ambil dari environment variable bila tersedia, lalu ubah ke SecureString
$pfxPlain = if ($env:PNP_PFX_PASSWORD) { $env:PNP_PFX_PASSWORD } else { $Config.PfxPasswordPlain }
if ($pfxPlain) {
    $Config.PfxPassword = ConvertTo-SecureString $pfxPlain -AsPlainText -Force
}
$Config.Remove("PfxPasswordPlain")
Remove-Variable pfxPlain -ErrorAction SilentlyContinue

# Folder report default (di-ignore oleh git)
if (-not $Config.ReportFolder) {
    $Config.ReportFolder = Join-Path (Split-Path $PSScriptRoot -Parent) "output"
}
if (-not (Test-Path -LiteralPath $Config.ReportFolder)) {
    New-Item -ItemType Directory -Path $Config.ReportFolder -Force | Out-Null
}
