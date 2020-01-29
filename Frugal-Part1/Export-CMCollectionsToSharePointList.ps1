param(
    [Parameter(mandatory=$true)]
    [string]$SiteCode,

    [Parameter(mandatory=$true)]
    [string]$ProviderMachineName,

    [Parameter(mandatory=$true)]
    [string]$SharePointSite,

    [Parameter(mandatory=$true)]
    [string]$SharePointListName,

    [string]$CollectionQuery = "*"
)
#Configure this section
$Username = ""
$PasswordSecureString = ""
#End Configuration

$PasswordSecureString = ConvertTo-SecureString $PasswordSecureString
$Credentials = New-Object System.Management.Automation.PSCredential($Username,$PasswordSecureString)

Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
Import-Module SharePointPnPPowerShellOnline

#Connect to SPOnline
$connection = Connect-PnPOnline -Url $SharePointSite -Credentials $Credentials
$Credentials = $null

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\"

$collections = Get-CMCollection -Name $CollectionQuery | Where-Object {$_.CollectionID -notlike "SMS*"}
$CurrentList = Get-PnPListItem -List $SharePointListName
$CurrentListTitles = @()
foreach($l in $CurrentList){ $CurrentListTitles += $l["Title"] }

# Remove Missing Collections
foreach($l in $CurrentList)
{
    if($l["Title"] -notin $collections.Name)
    {
        $l.DeleteObject()
    }
}

# Add New Collections
foreach($c in $collections)
{
    if($c.Name -notin $CurrentListTitles)
    {
        if($c.CollectionType -eq 2)
        {
            $colType = "Device"
        }else{
            $colType = "User"
        }
        $ch = @{
            "Title"="$($c.Name)";
            "CollectionType"="$($colType)";
            "CollectionDescription"="$($c.Comment)"
        }
        Add-PnPListItem -List $SharePointListName -Values $ch -Connection $connection
    }
}