# UserProfile Scripts

This is a userprofile related sample scripts. However this could be replaced with PnP cmdlet, such as Get-PnpUserProfileProperty and Set-PnpUserProfileProperty

https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpuserprofileproperty?view=sharepoint-ps

https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnpuserprofileproperty?view=sharepoint-ps

## Example 1
.\GetUserProfile.ps1 -mySiteUrl https://tenant-my.sharepoint.com | Out-GridView

## Example 2
.\GetUserProfileProperty.ps1 -mySiteUrl https://tenant-my.sharepoint.com -filter @domain.com -propertyName WorkEmail | Export-CSV .\test.csv -NoTypeInformation -Encoding ASCII

## Example 3
.\GetUserProfileProperty.ps1 -mySiteUrl https://tenant-my.sharepoint.com -filter admin@test.onmicrosoft.com -propertyName WorkEmail | .\SetUserProfilePropertyViaPipe.ps1 -propertyValue admin@outlook.com

## Example 4
.\SetUserProfileProperty.ps1 -adminUrl https://tenant-admin.sharepoint.com -accountName "i:0#.f|membership|admin@tenant.onmicrosoft.com" -propertyValue admin@outlook.com

