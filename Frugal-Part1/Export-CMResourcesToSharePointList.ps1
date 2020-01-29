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

$Resources = Get-CMResource -Fast
$Resources = $Resources | Where-Object {$_.Name -notlike "*Unknown*"} | Where-Object {$_.Name -notlike "*Provisioning Device*"}
$CurrentList = Get-PnPListItem -List $SharePointListName

$CurrentListDeviceResourceIds = @()
$CurrentListUserResourceIds = @()
$CurrentListUserGroupResourceIds = @()

# Build Current Resource Id Lists
foreach($l in $CurrentList){ 
    if($l["ResourceType"] -eq "User")
    {
        $CurrentListUserResourceIds += $l["ResourceId"]
    }
    elseif($l["ResourceType"] -eq "User Group")
    {
        $CurrentListUserGroupResourceIds += $l["ResourceId"]
    }
    elseif($l["ResourceType"] -eq "Device")
    {
        $CurrentListDeviceResourceIds += $l["ResourceId"]
    }
}

# Remove Missing Objects
foreach($l in $CurrentListDeviceResourceIds)
{
    if($l["ResourceType"] -eq "Device")
    {
        if($l["Title"] -notin ($Resources | Where {$_.SmsProviderObjectPath -like "SMS_R_System*"}).ResourceId)
        {
            $l.DeleteObject()
        }
    }
    elseif($l["ResourceType"] -eq "User")
    {
        if($l["Title"] -notin ($Resources | Where {$_.SmsProviderObjectPath -like "SMS_R_User.*"}).ResourceId)
        {
            $l.DeleteObject()
        }
    }
    elseif($l["ResourceType"] -eq "User Group")
    {
        if($l["Title"] -notin ($Resources | Where {$_.SmsProviderObjectPath -like "SMS_R_UserGroup*"}).ResourceId)
        {
            $l.DeleteObject()
        }
    }
}

# Add New Collections
foreach($r in $Resources)
{
    if(($r.ResourceId -notin $CurrentListDeviceResourceIds) -and ($r.ResourceId -notin $CurrentListUserResourceIds) -and ($r.ResourceId -notin $CurrentListUserGroupResourceIds))
    {
        if($r.SmsProviderObjectPath -like "SMS_R_User.*"){ $rType = "User" }
        elseif($r.SmsProviderObjectPath -like "SMS_R_UserGroup*"){ $rType = "User Group" }
        elseif($r.SmsProviderObjectPath -like "SMS_R_System*"){ $rType = "Device" }
        $rh = @{
            "Title"="$($r.Name)";
            "ResourceType"="$($rType)";
            "ResourceId"="$($r.ResourceId)"
        }
        Add-PnPListItem -List $SharePointListName -Values $rh -Connection $connection
    }
}