# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==============================================================================
# KONFIGURASI KREDENSIAL ZOHO
# ==============================================================================
$clientId     = $Config.Zoho.ClientId
$clientSecret = $Config.Zoho.ClientSecret
$refreshToken = $Config.Zoho.RefreshToken
$orgId        = $Config.Zoho.OrganizationId

# ==============================================================================
# LANGKAH 1: Refresh Access Token (OAuth2)
# ==============================================================================
$tokenUri = "https://accounts.zoho.com/oauth/v2/token"

# Menyusun body dengan format application/x-www-form-urlencoded
$body = @{
    refresh_token = $refreshToken
    grant_type    = "refresh_token"
    redirect_uri  = "http://www.zoho.com/books"
}

# Membuat Basic Authentication Header dari Client ID dan Client Secret
$pair = "$($clientId):$($clientSecret)"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)
$headers = @{
    "Authorization" = "Basic $base64"
    "Content-Type"  = "application/x-www-form-urlencoded"
}

Write-Host "Sedang memperbarui access token..." -ForegroundColor Cyan

try {
    # Melakukan request POST untuk refresh token
    $tokenResponse = Invoke-RestMethod -Uri $tokenUri -Method Post -Headers $headers -Body $body
    
    # Mengambil access token dari hasil respon JSON
    $accessToken = $tokenResponse.access_token

    if (-not $accessToken) {
        throw "Access token tidak ditemukan dalam respon. Periksa kredensial Anda."
    }
    
    Write-Host "Access token berhasil diperbarui!" -ForegroundColor Green
}
catch {
    Write-Error "Gagal me-refresh token: $_"
    exit
}

# ==============================================================================
# LANGKAH 2: List All Reporting Tags menggunakan Token Baru
# ==============================================================================
$apiUri = "https://www.zohoapis.com/books/v3/reportingtags?organization_id=$orgId"

# Menyusun header dengan format 'Zoho-oauthtoken [token]' sesuai dokumentasi Zoho
$apiHeaders = @{
    "Authorization" = "Zoho-oauthtoken $accessToken"
}

Write-Host "Sedang mengambil data Reporting Tags..." -ForegroundColor Cyan

try {
    # Melakukan request GET ke Zoho Books API
    $apiResponse = Invoke-RestMethod -Uri $apiUri -Method Get -Headers $apiHeaders
    
    # Menampilkan hasil output dalam format JSON yang rapi (Pretty Print)
    Write-Host "`nHasil Data Reporting Tags:" -ForegroundColor Green
    $apiResponse | ConvertTo-Json -Depth 10
}
catch {
    Write-Error "Gagal mengambil data dari Zoho Books API: $_"
}