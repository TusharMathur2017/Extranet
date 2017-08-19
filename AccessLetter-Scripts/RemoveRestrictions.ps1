clear
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null)
{ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }
        
#$web = Get-SPWebApplication "http://connect.bain.com"
#
#$web.ClientCallableSettings.AnonymousRestrictedTypes.Remove([Microsoft.SharePoint.SPList], "GetItems")
#$web.Update()

$site = Get-SPWeb "http://connect.bain.com/al" 
$list = $site.Lists["Consent"] 
write-host $list.AnonymousPermMask64
