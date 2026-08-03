$siteUrl = Read-Host -Prompt "Enter the site url"

if($siteUrl -like '*ADFS_SITE_URL.com*') {
    Connect-PnPOnline -url $siteUrl -UseWebLogin
}
else {
    Connect-PnPOnline -url $siteUrl -UseWebLogin # -Credentials (Get-Credential)
}
function UpdateFile ($localFile, $file, $folder) {
    $currentSiteServerRelativeUrl = Get-PnPSite -Includes ServerRelativeUrl | ForEach-Object { $_.ServerRelativeUrl }
    $currentSiteServerRelativeUrl = $currentSiteServerRelativeUrl.TrimEnd("/") 

     Get-PnPFolderItem -FolderSiteRelativeUrl "$folder/$file" -ItemType File
     Write-Host "File Located" -ForegroundColor Green
     Write-Host "Deploy Style Library Assets" -ForegroundColor Green
     Set-PnPFileCheckedOut -Url "$currentSiteServerRelativeUrl/$folder/$file"
     Add-PnPFile -Path $localFile -Folder $folder | Out-Null
     Set-PnPFileCheckedIn -Url "$currentSiteServerRelativeUrl/$folder/$file"
     Write-Host "Update Complete" -ForegroundColor Blue
}

$currentSiteServerRelativeUrl = Get-PnPSite -Includes ServerRelativeUrl | ForEach-Object { $_.ServerRelativeUrl }
$currentSiteServerRelativeUrl = $currentSiteServerRelativeUrl.TrimEnd("/")
UpdateFile "./banner.js" "banner.js" "SiteAssets/bannerapp/js"

#This is to turn on scripts if its turned off
Set-PnPSite -Identity $siteUrl -NoScriptSite $false

Add-PnPJavaScriptLink -Name "BannerApp" -Url "$currentSiteServerRelativeUrl/SiteAssets/bannerapp/js/banner.js" -Scope Site
