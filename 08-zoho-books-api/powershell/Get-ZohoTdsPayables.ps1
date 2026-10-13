# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# KONFIGURASI ZOHO API
# ==========================================
$clientId     = $Config.Zoho.ClientId
$clientSecret = $Config.Zoho.ClientSecret
$refreshToken = $Config.Zoho.RefreshToken
$organizationId = $Config.Zoho.OrganizationId

# Domain Zoho (Gunakan .com, .in, .eu, atau .com.au sesuai lokasi akun Zoho Anda)
$accountsDomain = "https://accounts.zoho.com" 
$apiDomain      = "https://www.zohoapis.com"

# Path Output File CSV
$outputCsvPath  = "$PSScriptRoot\Zoho_TDS_Payables_Report.csv"

# ==========================================
# 1. GET ACCESS TOKEN BARU
# ==========================================
Write-Host "Mengambil Access Token baru dari Zoho..." -ForegroundColor Cyan

$tokenUrl = "$accountsDomain/oauth/v2/token"
$tokenBody = @{
    refresh_token = $refreshToken
    client_id     = $clientId
    client_secret = $clientSecret
    grant_type    = "refresh_token"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
    $accessToken   = $tokenResponse.access_token

    if (-not $accessToken) {
        throw "Access token tidak ditemukan dalam response."
    }
    Write-Host "Access Token berhasil didapatkan!" -ForegroundColor Green
}
catch {
    Write-Host "Gagal mendapatkan Access Token: $_" -ForegroundColor Red
    exit
}

# ==========================================
# 2. GET TDS PAYABLES REPORT DATA
# ==========================================
Write-Host "Mengambil data laporan TDS Payables..." -ForegroundColor Cyan

# Endpoint API Zoho Books untuk TDS Payables Summary
$reportUrl = "$apiDomain/books/v3/reports/tdspayables?organization_id=$organizationId"

$headers = @{
    "Authorization" = "Zoho-oauthtoken $accessToken"
}

try {
    $reportResponse = Invoke-RestMethod -Uri $reportUrl -Method Get -Headers $headers

    # Memeriksa struktur JSON response dari Zoho
    if ($reportResponse.code -eq 0) {
        Write-Host "Data berhasil diambil!" -ForegroundColor Green
        
        # Ambil baris data tabel (biasanya ada di properti .tdspayablessummary atau .tds_payables)
        $tdsData = $reportResponse.tdspayablessummary
        
        if ($null -eq $tdsData) {
            # Fallback jika nama properti berupa daftar/array di JSON response
            $tdsData = $reportResponse
        }

        # Ekspor data ke CSV
        $tdsData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "File laporan berhasil disimpan di: $outputCsvPath" -ForegroundColor Yellow
    }
    else {
        Write-Host "Zoho API Error: $($reportResponse.message)" -ForegroundColor Red
    }
}
catch {
    Write-Host "Gagal mengambil data dari Zoho API: $_" -ForegroundColor Red
}