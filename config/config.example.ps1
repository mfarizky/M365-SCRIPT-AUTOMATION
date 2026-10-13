# =====================================================================
# CONTOH KONFIGURASI — salin menjadi `config.local.ps1` lalu isi nilainya.
# `config.local.ps1` sudah ada di .gitignore dan TIDAK BOLEH di-commit.
# =====================================================================

$Config = @{
    # --- Tenant Microsoft 365 ---
    TenantId            = "00000000-0000-0000-0000-000000000000"
    TenantDomain        = "contoso.onmicrosoft.com"
    SharePointHost      = "contoso.sharepoint.com"
    SharePointRoot      = "https://contoso.sharepoint.com"
    AdminUrl            = "https://contoso-admin.sharepoint.com"

    # --- Site utama ---
    HubSiteUrl          = "https://contoso.sharepoint.com/sites/hub"            # site internal (master lists)
    TemplatePortalUrl   = "https://contoso.sharepoint.com/sites/template-portal" # portal klien yang jadi sumber template
    ExcludedSiteUrl     = "https://contoso.sharepoint.com/sites/excluded"       # site yang dikecualikan dari mass update
    InternalTitleKeyword = "CONTOSO"                                            # kata kunci judul site internal (di-skip)

    # --- Entra ID App Registration ---
    AppClientId         = "00000000-0000-0000-0000-000000000000"  # app-only (sertifikat)
    InteractiveClientId = "00000000-0000-0000-0000-000000000000"  # delegated / interactive login
    PfxPath             = "C:\Secure\Cert\PnPAutomationCert.pfx"  # simpan DI LUAR repo
    PfxPasswordPlain    = ""   # kosongkan & pakai env var PNP_PFX_PASSWORD (disarankan)

    # --- Output ---
    ReportFolder        = ""   # kosong = <repo>\output

    # --- Exchange / Calendar ---
    SharedMailbox       = "shared@contoso.com"
    CalendarCheckUser   = "user@contoso.com"
    CalendarCheckRoom   = "Meeting Room 1"

    # --- Recovery (04-onedrive-file-compression/recovery) ---
    ActivityUserUpn         = "user@contoso.com"
    ActivityUserDisplayName = "User Name"

    # --- Data fix: mapping Account Manager -> Tax Manager ---
    TaxManagerRules = @(
        @{ AccountManagerKeyword = "ANDI"; TaxManagerUpn = "budi@contoso.com";  TaxManagerName = "BUDI" }
        @{ AccountManagerKeyword = "SARI"; TaxManagerUpn = "dewi@contoso.com";  TaxManagerName = "DEWI" }
    )

    # --- File compression ---
    Compression = @{
        ImageScanRoot    = "C:\Users\<you>\OneDrive - Contoso\Real Estate"
        ImageSpBaseUrl   = "https://contoso.sharepoint.com/sites/real/Shared%20Documents"
        PdfScanRoot      = "C:\Users\<you>\OneDrive - Contoso\Legal"
        PdfSpBaseUrl     = "https://contoso.sharepoint.com/sites/legal/Shared%20Documents"
        PdfCorporateRoot = "C:\Users\<you>\OneDrive - Contoso\Legal\ACCOUNTS\CORPORATE"
        PdfBackupRoot      = "D:\Backup\PDF\Ghostscript\CORPORATE"
        PdfBackupRootAdobe = "D:\Backup\PDF\Adobe\CORPORATE"
        OneDriveMarker   = "OneDrive - Contoso\"
        GhostscriptPath  = "C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe"
    }

    # --- Adobe PDF Services API ---
    Adobe = @{
        ClientId     = "<adobe-client-id>"
        ClientSecret = "<adobe-client-secret>"
    }

    # --- Zoho Books API ---
    Zoho = @{
        ClientId       = "1000.XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
        ClientSecret   = "<zoho-client-secret>"
        RefreshToken   = "<zoho-refresh-token>"
        OrganizationId = "<zoho-organization-id>"
    }
}
