# Author - Tim Wheeler (http://blog.timwheeler.io)
# This script gets a twitter feed and publishes it to a list.
# Process: 
# One Off: Register as a developer in twitter.  Create a new app and generate the keys.
#          Call this with the consumer details.  Copy and paste the bearer token.
# Scheduled Task: Create a scheduled task.  Pass in the details with the bearer token. (Consumer credentials no longer needed)
param(
    $siteUrl = $(throw "SiteURL not specified"),
    $spOnline = $true,
    $username = $null,
    $password = $null,
    $maxItems = 10,
    $screenName = $(throw "You must provide a screenname"),
    $authToken = $null, #optional - saves an auth check - If you use this you do not need the consumer key and secret.
    $consumerKey = $null, #Use the consumer credentials or the authToken
    $consumerSecret = $null
    
)
Add-Type -LiteralPath ($PSScriptRoot + "\CSOM\Microsoft.SharePoint.Client.dll") -PassThru | out-null
Add-Type -LiteralPath ($PSScriptRoot + "\CSOM\Microsoft.SharePoint.Client.Runtime.dll") -PassThru | out-null
Add-Type -AssemblyName System.Web

function ToBase64($toEncode)
{
    Add-Type -AssemblyName "System.Text.Encoding" | out-null
    return [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($toEncode))
}

Function GetauthToken($consumerKey, $consumerSecret, $authUrl)
{
    Write-Host "Attempting to `aquire a bearer token using the app credentials"
    $response = Invoke-RestMethod  $authUrl -Headers @{"Authorization"="Basic " + $encodedCredentials} -Body "grant_type=client_credentials" -Method Post -TimeoutSec 60 -ErrorAction:Stop -ContentType "application/x-www-form-urlencoded;charset=UTF-8"
    if($response.token_type -ne "bearer")
    {
        throw "Error: the returned token is not a bearer token, we got $($response.token_type)"
        return
    }
    Write-Host "Bearer Token: $($response.access_token)"
    return $response.access_token
}
Function GetTwitterFeed($url, $authToken)
{
    $bearerCredentials = "Bearer " + $authToken
    $head = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $head = @{}
    $head.Add("Authorization",  $bearerCredentials)
    $head.Add("Accept-Encoding","gzip")
    $response = Invoke-WebRequest  $url -Headers $head -Method Get -TimeoutSec 60 -ErrorAction:Stop 
    if($response.StatusCode -ne 200)
    {
      throw "Got a $($response.StatusCode) response from twitter, expected 200"
      return $null
    }
    return $response
}
Function Login-Web
{
  [CmdletBinding(DefaultParameterSetName="SiteUrl")]
  param(
    [Parameter(Mandatory=$true, Position=0, ParameterSetName="SiteUrl", 
                   ValueFromPipeline=$true, 
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Url to SharePoint Site")]
        [ValidateNotNullOrEmpty()]
        [string]
        $siteUrl, 
    [Parameter(Mandatory=$false, Position=1, 
                   HelpMessage="Username for login")]
        [string]
        $username,
    
      [Parameter(Mandatory=$false, Position=2,
                   HelpMessage="Password string to login with")]
        [string]
        $password,
       [Parameter(Mandatory=$false, Position=3, HelpMessage="Password string to login with")]
       
        $spOnline = $true
  )
   
  $context = Create-Context $siteUrl $username $password $spOnline
  $context.Load($context.Web)
  $context.Load($context.Site)
  try
  {
    $context.ExecuteQuery()
    Write-Host ("Connected to " + $context.Web.Url) -foregroundcolor black -backgroundcolor Green
    return $context
  }
  catch
  {
    throw "Unable to reach $siteUrl : $_"
  }
}
Function Create-Context($siteUrl, $username, $password, $spOnline = $true) 
{
[System.Security.SecureString] $securePassword = $null
  if(($spOnline -eq $true -and [string]::IsNullOrEmpty($password)) -or ($spOnline -eq $false -and [string]::IsNullOrEmpty($username) -ne $true -and [string]::IsNullOrEmpty($password) -eq $true))
  {
    $psCredential = $null
    if($username -ne $null )
    {
        $psCredential = Get-Credential $username
    }
    else {
        $psCredential = Get-Credential
    }
    if($psCredential -eq $null)
    {
      write-error "Exit due to no credentials"
      return $null
    }
    $securePassword = ConvertTo-SecureString -String ($psCredential.GetNetworkCredential().Password) -AsPlainText -Force 
    $username = $psCredential.GetNetworkCredential().UserName
    #$domain = $psCredential.GetNetworkCredential().Domain
  }
  elseif([string]::IsNullOrEmpty($password) -eq $false) {
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
  }
  $context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($siteUrl)
  if($spOnline)
  {
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword) -ErrorAction:Stop
  }
  elseif($spOnline -eq $false -and [string]::IsNullOrEmpty($username) -ne $true)
  {
    $context.Credentials = New-Object System.Net.NetworkCredential($username, $securePassword) -ErrorAction:Stop
  }
  return $context
}
Function UploadFeed($context, $response)
{
  write-host "Connecting to site to upload feed"
    
    [Microsoft.SharePoint.Client.Web] $web = $context.Site.RootWeb
    $list = $web.Lists.GetByTitle("Twitter Cache")
    [Microsoft.SharePoint.Client.CamlQuery] $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = "<View><ViewFields> <FieldRef Name='Title'></FieldRef><FieldRef Name='Twitter_x0020_Cache'></FieldRef></ViewFields><Query></Query><RowLimit>10</RowLimit></View>"
    $items = $list.GetItems($query)
    $context.Load($items)
    $context.ExecuteQuery()
    if($items.Count -gt 1)
    {
      Write-Warning "Found more than 1 item in the Twitter Cache List.  Removing extra items."
      #should only ever be 1, so delete all but the first.
      for($i = $items.Count - 1; $i -gt 0; $i--)
      {
          $items[$i].DeleteObject()
      }
      $context.ExecuteQuery()
      $items = $list.GetItems($query) #reload
      $context.Load($items)
      $context.ExecuteQuery()
    }
    $listItem = $null
    if($items.Count -eq 0)
    {
      #create new
      [Microsoft.SharePoint.Client.ListItemCreationInformation] $itemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
      $listItem = $list.AddItem($itemInfo)
    }
    else  {
      $listItem = $items[0]
    }
    $listItem["Title"] = ("Automated feed upload $(Get-Date)")
    
    $listItem["Twitter_x0020_Cache"] = $response.Content
    $listItem.Update()
    $context.ExecuteQuery()
    write-host "Updated twitter feed - Saved to list Twitter Cache"
}

if($authToken -eq $null)
{
  $consumerKey = [System.Web.HttpUtility]::UrlEncode($consumerKey)
  $consumerSecret = [System.Web.HttpUtility]::UrlEncode($consumerSecret)
  $consumerCredentials = $consumerKey + ":" + $consumerSecret
  $authUrl = "https://api.twitter.com/oauth2/token"
  $encodedCredentials = ToBase64 $consumerCredentials    
  $authToken = GetauthToken $consumerKey $consumerSecret $authUrl
}
if($authToken -ne $null)
{
  $feedUrl = "https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=$screenName&count=$maxItems"
  $response = GetTwitterFeed $feedUrl $authToken
}
if($response -ne $null)
{
  try 
  {
    [Microsoft.SharePoint.Client.ClientContext] $context = Login-Web $siteUrl $username $password -SPOnline:$spOnline
    UploadFeed $context $response
  }
  finally
  {
    if($context -ne $null)
    {
      $context.Dispose()
    } 
  }
}




