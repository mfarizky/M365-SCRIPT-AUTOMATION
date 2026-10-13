# Load environment-specific settings (see config/README.md)
. (Join-Path $PSScriptRoot "..\..\config\Load-Config.ps1")

$ClientId = $Config.InteractiveClientId
Connect-PnPOnline -Url $Config.HubSiteUrl -ClientId $ClientId

$TargetListName  = "DB WORKFLOWS"
$SourceListName  = "DATABASE TAKA LEGAL"
$SourceFieldName = "TAKA"   # INTERNAL NAME

$ctx = Get-PnPContext

# GET SOURCE LIST ID (AUTO, NO HARDCODE)
$SourceList = Get-PnPList -Identity $SourceListName
$SourceListId = $SourceList.Id.ToString()

for ($i = 1; $i -le 20; $i++) {

    $DisplayName  = "TAKA $i"
    $InternalName = "TAKA_$i"

    Write-Host "Creating $DisplayName..."

    $field = Add-PnPField `
        -List $TargetListName `
        -DisplayName $DisplayName `
        -InternalName $InternalName `
        -Type Lookup

    $lookup = $field.TypedObject
    $lookup.LookupList  = $SourceListId
    $lookup.LookupField = $SourceFieldName

    $lookup.Update()
}

# EXECUTE SEKALI SAJA
$ctx.ExecuteQuery()
