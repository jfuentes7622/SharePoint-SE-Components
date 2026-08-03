$siteUrl = "https://goldentroutfarms.sharepoint.com"
$appPath = "C:\Source\BannerClock\sharepoint\solution\branding.sppkg"
$appTitle = "branding-client-side-solution"

if($siteUrl -like '*ADFS_SITE_URL.com*') {
    Connect-PnPOnline -url $siteUrl -UseWebLogin
}
else {
    Connect-PnPOnline -url $siteUrl -UseWebLogin # -Credentials (Get-Credential)
}

#scope of site is for site collections, removeing scope does tenant but only if you control tenant
Add-PnPApp -Path $appPath -Scope Site -Publish

$app = Get-PnPApp | Where-Object { $_.Title -eq $appTitle }

install-pnpapp $app

#to uninstall and remove
uninstall-pnpapp $app
remove-pnpapp $app


#Following is to add the settings for the site collections tenant wide setting 
#this works tenant wide if the app catalog is tenant wide, or site collection wide if the app catalog is per site collection
#keep in mind this would have to be done for each site collection with app catalog


Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/appcatalog" -Interactive

Add-PnPListItem -List "Tenant Wide Extensions" -Values @{
     Title = "Banner-Clock-Branding";
     Location = "ClientSideExtension.ApplicationCustomizer";
     ComponentId = "6fad0110-ac0c-46b4-a16b-337c1c6df3c4";
     ComponentProperties = '{ "rootUrl":"https://siteurlhere"}'
}
