# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\config\Load-Config.ps1")

# ==============================================
# Shared Mailbox Unified Audit Report
# Mailbox : diatur lewat $Config.SharedMailbox
#
# GMT+8 : 20 Jul 2026 00:00 -> 22 Jul 2026 23:59:59
# UTC   : 19 Jul 2026 16:00 -> 22 Jul 2026 15:59:59
# ==============================================

Connect-ExchangeOnline

$Mailbox = $Config.SharedMailbox

$Start = Get-Date "2026-07-19T16:00:00Z"
$End   = Get-Date "2026-07-22T15:59:59Z"

$ExportPath = (Join-Path $Config.ReportFolder "SharedMailbox_Audit_20260720_20260722.csv")

Write-Host "Collecting Unified Audit Logs..."

$Logs = Search-UnifiedAuditLog `
    -StartDate $Start `
    -EndDate $End `
    -ResultSize 5000

$Results = foreach ($Log in $Logs)
{
    try{
        $Data = $Log.AuditData | ConvertFrom-Json
    }
    catch{
        continue
    }

    if(($Data.MailboxOwnerUPN -eq $Mailbox) -or ($Log.AuditData -match [regex]::Escape($Mailbox)))
    {

        $Recipients = @()

        if($Data.DestFolderPathName){}

        if($Data.Recipients){
            $Recipients += ($Data.Recipients -join "; ")
        }

        if($Data.To){
            $Recipients += ($Data.To -join "; ")
        }

        if($Data.ToRecipients){
            $Recipients += ($Data.ToRecipients -join "; ")
        }

        $Recipients = $Recipients -join "; "

        $RestoreAction = switch ($Log.Operations)
        {
            "Restore"             {"YES"}
            "MailItemsRestored"   {"YES"}
            "Recover"             {"YES"}
            "RecoverDeletedItems" {"YES"}
            default               {"NO"}
        }

        [PSCustomObject]@{

            UTC_Time        = $Log.CreationDate.ToUniversalTime()
            GMT8_Time       = $Log.CreationDate.ToUniversalTime().AddHours(8)

            Operation       = $Log.Operations
            Restore         = $RestoreAction
            Result          = $Data.ResultStatus

            User            = $Log.UserIds
            Mailbox         = $Data.MailboxOwnerUPN
            LogonType       = $Data.LogonType

            Subject         = $Data.ItemSubject
            Sender          = $Data.Sender
            From            = $Data.From
            Recipients      = $Recipients

            Folder          = $Data.FolderPathName

            Client          = $Data.ClientInfoString
            ClientIP        = $Data.ClientIPAddress

            AppId           = $Data.AppId
            SessionId       = $Data.SessionId

            InternetMessageId = $Data.InternetMessageId

            OrganizationId  = $Data.OrganizationId

            RawOperation    = $Log.Operations
        }
    }
}

$Results |
Sort-Object UTC_Time |
Export-Csv $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "==========================================="
Write-Host "Mailbox : $Mailbox"
Write-Host "GMT+8   : 20 Jul 2026 00:00 - 22 Jul 2026 23:59:59"
Write-Host "UTC     : 19 Jul 2026 16:00 - 22 Jul 2026 15:59:59"
Write-Host "Total Events : $($Results.Count)"
Write-Host "CSV : $ExportPath"
Write-Host "==========================================="