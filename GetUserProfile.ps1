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
    $mySiteUrl,
    [Parameter(Mandatory=$false)]   
    $filter,
    $ResetCred = $false
)


[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.UserProfiles, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")

if (($global:UPSCred -eq $null) -or $ResetCred)
{
  $global:UPSCred = Get-Credential
}

$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($mySiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:UPSCred.UserName, $global:UPSCred.Password) 
$script:context.Credentials = $credentials


$script:context.add_ExecutingWebRequest({
    param ($source, $eventArgs);
    $request = $eventArgs.WebRequestExecutor.WebRequest;
    $request.UserAgent = "NONISV|Contoso|Application/1.0";
  })
  
  function ExecuteQueryWithIncrementalRetry {
    param (
        [parameter(Mandatory = $true)]
        [int]$retryCount
    );
  
    $DefaultRetryAfterInMs = 120000;
    $RetryAfterHeaderName = "Retry-After";
    $retryAttempts = 0;
  
    if ($retryCount -le 0) {
        throw "Provide a retry count greater than zero."
    }
  
    while ($retryAttempts -lt $retryCount) {
        try {
            $script:context.ExecuteQuery();
            return;
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response
  
            if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {
                $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
                $retryAfterInMs = $DefaultRetryAfterInMs;
  
                if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                    if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInMs)) {
                        $retryAfterInMs = $DefaultRetryAfterInMs;
                    }
                    else {
                        $retryAfterInMs *= 1000;
                    }
                }
  
                Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInMs / 1000))
                #Add delay.
                Start-Sleep -m $retryAfterInMs
                #Add to retry count.
                $retryAttempts++;
            }
            else {
                throw;
            }
        }
    }
  
    throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
  }
  

$script:context.Load($script:context.Web.SiteUsers)
ExecuteQueryWithIncrementalRetry -retryCount 5

foreach ($user in $script:context.Web.SiteUsers)
{
    if ($user.LoginName.Contains($filter))
    {
        $pm = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($script:context)
        $properties = $pm.GetPropertiesFor($user.LoginName)
        $script:context.Load($properties)
        ExecuteQueryWithIncrementalRetry -retryCount 5
        if (![System.String]::IsNullOrEmpty($properties.AccountName))
        {
            $obj = New-Object -TypeName PSObject

            foreach($key in $properties.UserProfileProperties.Keys)
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name $key -Value $properties.UserProfileProperties[$key]
            }
            $obj
        }
    }
}





