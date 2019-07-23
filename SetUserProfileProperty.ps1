<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 

 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
    [Parameter(Mandatory=$true)]
    $adminSiteUrl,
    [Parameter(Mandatory=$true)]
    $accountName,
    [Parameter(Mandatory=$true)]
    $propertyName,
    [Parameter(Mandatory=$true)]
    $propertyValue,
    [parameter(DontShow)]
    $oldPropertyValue,
    $ResetCred = $false
)

$ErrorActionPreference = "Stop"

[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.UserProfiles, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")

if (($global:UPSCred -eq $null) -or $ResetCred)
{
  $global:UPSCred = Get-Credential
}

$context = New-Object Microsoft.SharePoint.Client.ClientContext($adminSiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:UPSCred.UserName, $global:UPSCred.Password) 
$context.Credentials = $credentials

$pm = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)

if ($oldPropertyValue -eq $null)
{
    $prop = $pm.GetUserProfilePropertyFor($accountName, $propertyName)
    $context.ExecuteQuery()
    $oldPropertyValue = $prop.Value
}

if ($oldPropertyValue -ne $propertyValue)
{
    $pm.SetSingleValueProfileProperty($accountName, $propertyName, $propertyValue)
    $context.ExecuteQuery()

    ($accountName + " : " + $propertyName + " changed from '" + $oldPropertyValue + "' to '" + $propertyValue + "'")
}
