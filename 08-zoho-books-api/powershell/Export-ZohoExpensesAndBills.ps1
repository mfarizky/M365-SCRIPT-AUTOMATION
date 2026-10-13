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
# LANGKAH 1: Refresh Access Token
# ==============================================================================
$tokenUri = "https://accounts.zoho.com/oauth/v2/token"
$body = @{
    refresh_token = $refreshToken
    grant_type    = "refresh_token"
    redirect_uri  = "http://www.zoho.com/books"
}
$pair = "$($clientId):$($clientSecret)"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)
$headers = @{
    "Authorization" = "Basic $base64"
    "Content-Type"  = "application/x-www-form-urlencoded"
}

Write-Host "Sedang memperbarui access token..." -ForegroundColor Cyan
try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenUri -Method Post -Headers $headers -Body $body
    $accessToken = $tokenResponse.access_token
    if (-not $accessToken) { throw "Token kosong." }
    Write-Host "Access token berhasil diperbarui!" -ForegroundColor Green
} catch {
    Write-Error "Gagal me-refresh token: $_"
    exit
}

# Inisialisasi Header API Zoho Books
$apiHeaders = @{
    "Authorization" = "Zoho-oauthtoken $accessToken"
    "Accept"        = "application/json"
}

$allDetailedExpenses = [System.Collections.Generic.List[Object]]::new()
$allDetailedBills    = [System.Collections.Generic.List[Object]]::new()

# ==============================================================================
# LANGKAH 2: Ambil List & Detail EXPENSES (Menggunakan WebRequest agar Murni Teks)
# ==============================================================================
$listExpensesUri = "https://www.zohoapis.com/books/v3/expenses?organization_id=$orgId"
Write-Host "`n[EXPENSES] Mengambil daftar awal..." -ForegroundColor Cyan

try {
    # Ambil konten mentah berupa string murni menggunakan Invoke-WebRequest
    $rawExpRequest = Invoke-WebRequest -Uri $listExpensesUri -Method Get -Headers $apiHeaders -UseBasicParsing
    $rawExpText = $rawExpRequest.Content
    
    # Amankan format expense_id menggunakan regex string standar
    $safeExpText = $rawExpText -replace '"expense_id":\s*(\d+)', '"expense_id": "$1"'
    
    # Sekarang aman dikonversi ke Objek
    $expensesResponse = ConvertFrom-Json $safeExpText
    $expensesList = $expensesResponse.expenses

    Write-Host "[EXPENSES] Ditemukan $($expensesList.Count) item. Menarik detail item satu per satu..." -ForegroundColor Yellow

    foreach ($exp in $expensesList) {
        $expenseId = $exp.expense_id
        if (-not $expenseId) { continue }

        Write-Host "   -> GET Detail Expense ID: $expenseId" -ForegroundColor Gray
        $detailExpUri = "https://www.zohoapis.com/books/v3/expenses/${expenseId}?organization_id=$orgId"
        
        Start-Sleep -Milliseconds 600 
        try {
            $rawDetailRequest = Invoke-WebRequest -Uri $detailExpUri -Method Get -Headers $apiHeaders -UseBasicParsing
            $safeDetailText = $rawDetailRequest.Content -replace '"expense_id":\s*(\d+)', '"expense_id": "$1"'
            
            $expenseDetail = (ConvertFrom-Json $safeDetailText).expense
            $allDetailedExpenses.Add($expenseDetail)
        } catch {
            Write-Host "   x Gagal mengambil detail Expense ID $expenseId" -ForegroundColor Red
        }
    }
} catch {
    Write-Error "Gagal memproses kelompok data Expenses: $_"
}

# ==============================================================================
# LANGKAH 3: Ambil List & Detail BILLS (Menggunakan WebRequest agar Murni Teks)
# ==============================================================================
$listBillsUri = "https://www.zohoapis.com/books/v3/bills?organization_id=$orgId"
Write-Host "`n[BILLS] Mengambil daftar awal..." -ForegroundColor Cyan

try {
    # Ambil konten mentah berupa string murni menggunakan Invoke-WebRequest
    $rawBillRequest = Invoke-WebRequest -Uri $listBillsUri -Method Get -Headers $apiHeaders -UseBasicParsing
    $rawBillText = $rawBillRequest.Content
    
    # Amankan format bill_id menggunakan regex string standar (-replace bawaan PowerShell)
    $safeBillText = $rawBillText -replace '"bill_id":\s*(\d+)', '"bill_id": "$1"'
    
    $billsResponse = ConvertFrom-Json $safeBillText
    $billsList = $billsResponse.bills

    Write-Host "[BILLS] Ditemukan $($billsList.Count) item. Menarik detail item satu per satu..." -ForegroundColor Yellow

    foreach ($bill in $billsList) {
        $billId = $bill.bill_id
        if (-not $billId) { continue }

        Write-Host "   -> GET Detail Bill ID: $billId" -ForegroundColor Gray
        $detailBillUri = "https://www.zohoapis.com/books/v3/bills/${billId}?organization_id=$orgId"
        
        Start-Sleep -Milliseconds 600
        try {
            $rawDetailBillRequest = Invoke-WebRequest -Uri $detailBillUri -Method Get -Headers $apiHeaders -UseBasicParsing
            $safeDetailBillText = $rawDetailBillRequest.Content -replace '"bill_id":\s*(\d+)', '"bill_id": "$1"'
            
            $billDetail = (ConvertFrom-Json $safeDetailBillText).bill
            $allDetailedBills.Add($billDetail)
        } catch {
            Write-Host "   x Gagal mengambil detail Bill ID $billId" -ForegroundColor Red
        }
    }
} catch {
    Write-Error "Gagal memproses kelompok data Bills: $_"
}

# ==============================================================================
# LANGKAH 4: Penggabungan Akhir Menjadi 1 JSON Master
# ==============================================================================
Write-Host "`nMenyusun seluruh data ke dalam 1 master JSON..." -ForegroundColor Cyan

$masterOutput = @{
    total_expenses = $allDetailedExpenses.Count
    total_bills    = $allDetailedBills.Count
    expenses       = $allDetailedExpenses
    bills          = $allDetailedBills
} | ConvertTo-Json -Depth 10

# Menyimpan hasil file dump JSON di direktori eksekusi skrip
$outputPath = ".\zoho_master_dump.json"
$masterOutput | Out-File -FilePath $outputPath -Encoding utf8

Write-Host "Selesai! Seluruh data sukses disimpan di: $outputPath" -ForegroundColor Green