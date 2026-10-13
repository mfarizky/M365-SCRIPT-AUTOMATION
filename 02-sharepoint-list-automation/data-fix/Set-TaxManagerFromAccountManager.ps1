# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

# ==========================================
# 1. KONFIGURASI TARGET SITE & LIST
# ==========================================
$siteUrl        = $Config.HubSiteUrl
$listName       = "FINANCIAL COMPLIANCE"

# Credentials & Cert Path
$tenant         = $Config.TenantDomain
$clientId       = $Config.AppClientId
$pfxPath        = $Config.PfxPath
$pfxPassword    = $Config.PfxPassword

# Aturan mapping Account Manager -> Tax Manager (diatur di config.local.ps1)
# Contoh: @{ AccountManagerKeyword = "ANDI"; TaxManagerUpn = "budi@contoso.com"; TaxManagerName = "BUDI" }
$rules          = $Config.TaxManagerRules

# ==========================================
# 2. KONEKSI DAN PENARIKAN DATA
# ==========================================
Write-Host "Connecting to SharePoint Site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $pfxPath -CertificatePassword $pfxPassword

Write-Host "Fetching items from '$listName'..." -ForegroundColor Cyan
$allItems = Get-PnPListItem -List $listName -PageSize 2000 -ErrorAction SilentlyContinue

if (-not $allItems) {
    Write-Host "[ERROR] List '$listName' tidak ditemukan atau kosong!" -ForegroundColor Red
    Disconnect-PnPOnline; exit
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   RUNNING AUTO UPDATE: TAX MANAGER       " -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan

$updatedCount = 0
$skippedCount = 0

# ==========================================
# 3. LOGIKA UPDATE
# ==========================================
foreach ($item in $allItems) {
    # Ambil nilai dari internal field ACCOUNTMANAGER
    $accountMgrVal = $item["ACCOUNTMANAGER"]
    
    # Ekstrak string Nama/Email dari field ACCOUNTMANAGER
    $accountMgrStr = ""
    if ($accountMgrVal -is [Microsoft.SharePoint.Client.FieldUserValue]) {
        $accountMgrStr = $accountMgrVal.LookupValue # Menangkap Display Name
    } elseif ($null -ne $accountMgrVal) {
        $accountMgrStr = $accountMgrVal.ToString()
    }

    # CONDITIONAL UPDATES (HANYA YANG COCOK DENGAN RULE, SELAIN ITU DI-SKIP)
    $targetTaxMgr = $null
    $matchedRule  = $rules | Where-Object { $accountMgrStr -like "*$($_.AccountManagerKeyword)*" } | Select-Object -First 1
    if ($matchedRule) { $targetTaxMgr = $matchedRule.TaxManagerUpn }

    # Eksekusi Update jika kondisi terpenuhi
    if ($null -ne $targetTaxMgr) {
        try {
            Set-PnPListItem -List $listName -Identity $item.Id -Values @{ "TAXMANAGER" = $targetTaxMgr } -ErrorAction Stop
            Write-Host "[UPDATED] ID $($item.Id): ACCOUNTMANAGER = '$accountMgrStr' -> TAX MANAGER = '$targetTaxMgr'" -ForegroundColor Green
            $updatedCount++
        }
        catch {
            # Fallback jika kolom TAX MANAGER bertipe Text biasa
            $namaSimpel = $matchedRule.TaxManagerName
            try {
                Set-PnPListItem -List $listName -Identity $item.Id -Values @{ "TAX MANAGER" = $namaSimpel } -ErrorAction Stop
                Write-Host "[UPDATED-TEXT] ID $($item.Id): ACCOUNTMANAGER = '$accountMgrStr' -> TAX MANAGER = '$namaSimpel'" -ForegroundColor Green
                $updatedCount++
            }
            catch {
                Write-Host "[ERROR] Gagal update ID $($item.Id): $_" -ForegroundColor Red
            }
        }
    } else {
        # Tidak cocok dengan rule manapun -> Di-skip tanpa diproses
        $skippedCount++
    }
}

Disconnect-PnPOnline

Write-Host "==========================================" -ForegroundColor Green
Write-Host " [SELESAI] Total $updatedCount item di-update." -ForegroundColor Green
Write-Host " [SKIPPED] Total $skippedCount item di-skip." -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Green