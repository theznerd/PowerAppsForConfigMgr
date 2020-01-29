param(
    [Parameter(mandatory=$true)]
    [string]$SharePointSite,

    [Parameter(mandatory=$true)]
    [string]$ListName
)
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
Import-Module SharePointPnPPowerShellOnline
$connection = Connect-PnPOnline -Url $SharePointSite -Credentials (Get-Credential)
New-PnPList -Title $ListName -Connection $connection -Template GenericList -OnQuickLaunch:$False -EnableVersioning:$False -Hidden:$true
Set-PnPList -Identity $ListName -EnableAttachments:$false
Set-PnPField -List $ListName -Identity "Title" -Values @{Title="Collection Name"} -Connection $connection
Add-PnPField -List $ListName -DisplayName "Collection Type" -InternalName "CollectionType" -Type Choice -Choices @("User","Device") -Required:$true -AddToDefaultView:$true -Connection $connection
Add-PnPField -List $ListName -DisplayName "Collection Description" -InternalName "CollectionDescription" -Type Text -Required:$false -AddToDefaultView:$true -Connection $connection