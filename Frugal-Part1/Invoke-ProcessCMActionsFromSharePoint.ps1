param(
    [Parameter(mandatory=$true)]
    [string]$SiteCode,

    [Parameter(mandatory=$true)]
    [string]$ProviderMachineName,

    [Parameter(mandatory=$true)]
    [string]$SharePointSite,

    [Parameter(mandatory=$true)]
    [string]$SharePointListName
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

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\"

# Get Actions that are not yet completed
$actions = Get-PnPListItem -List $SharePointListName -Query "<View><Query><Where><Eq><FieldRef Name='Completed'/><Value Type='Text'>false</Value></Eq></Where></Query></View>"

# Execute Actions
foreach($a in $actions)
{
    if($a["ResourceType"] -eq "Device")
    {
        try
        {
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $a["Title"] -ResourceId $a["ResourceId"]
            Set-PnPListItem -List $SharePointListName -Identity $a -Values @{"Completed" = "true"}
        }
        catch
        {
            Set-PnPListItem -List $SharePointListName -Identity $a -Values @{"Completed" = "error"; "Details" = "$($error[0].ToString())"}
        }
    }
    else
    {
        try
        {
            Add-CMUserCollectionDirectMembershipRule -CollectionName $a["Title"] -ResourceId $a["ResourceId"]
            Set-PnPListItem -List $SharePointListName -Identity $a -Values @{"Completed" = "true"}
        }
        catch
        {
            Set-PnPListItem -List $SharePointListName -Identity $a -Values @{"Completed" = "error"; "Details" = "$($error[0].ToString())"}
        }
    }
}