#  Author: Tim Wheeler (http://blog.timwheeler.tech/)
$ensureFolderHistory = new-object "System.Collections.ArrayList"
$productName =" sp-clientapps automated release process"
<#
    .SYNOPSIS
    Uses SP Online or on premise ad.  Depends on $spOnline parameter.
    .DESCRIPTION
    .PARAMETER SiteUrl
    The Url to the site eg : http://mycompany.sharepoint.com
    .PARAMETER UserName
    (Optional UserName - you will be prompted)
    .PARAMETER UserName
    (Optional Password - you will be prompted)
    .EXAMPLE
    C:\PS> Login-Web -SiteUrl "https://mycompany.sharepoint.com" -UserName = "myuser@mycompany.onmicrosoft.com"
    <Description of example>
    .NOTES
    Author: Tim Wheeler (http://blog.timwheeler.tech/)
#>
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
    [Parameter(Mandatory=$true, Position=1, 
                   HelpMessage="Username for login")]
        [string]
        $username,
    
      [Parameter(Mandatory=$false, Position=2,
                   
                   HelpMessage="Password string to login with")]
        [string]
        $password,
       [Parameter(Mandatory=$false, Position=3,
                   HelpMessage="Password string to login with")]
       
        $spOnline = $true
  )
   
  $context = Create-Context $siteUrl $username $password $spOnline
  $context.Load($context.Web)
  $context.Load($context.Site)
  $context.Load($context.Site.RootWeb)
  try
  {
    $context.ExecuteQuery()
    Write-Host "Connected to " $context.Web.Url -foregroundcolor green
    return $context
  }
  catch
  {
    throw "Unable to reach $siteUrl : $_"
  }
}
Function Is-True($value) {
    try 
    {
      $result = [System.Convert]::ToBoolean($value) 
    } 
    catch [FormatException] 
    {
      $result = $false
    }
    return $result
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
Function Ensure-Folder()
{
  Param(

    [Parameter(Mandatory=$True)]
    [Microsoft.SharePoint.Client.Web]$web,

    [Parameter(Mandatory=$True)]
    [Microsoft.SharePoint.Client.Folder]$parentFolder, 

    [Parameter(Mandatory=$True)]
    [String]$folderUrl

  )
    #history check to avoid duplicate calls
    if($ensureFolderHistory.Contains($folderUrl))
    {
        return;
    }
    Create-Folder $web $parentFolder $folderUrl
    $ensureFolderHistory.Add($folderUrl)
}
Function Create-Folder()
{
  Param(

    [Parameter(Mandatory=$True)]
    [Microsoft.SharePoint.Client.Web]$web,

    [Parameter(Mandatory=$True)]
    [Microsoft.SharePoint.Client.Folder]$parentFolder, 

    [Parameter(Mandatory=$True)]
    [String]$folderUrl
  )
    $folderNames = $folderUrl.Trim().Split("/",[System.StringSplitOptions]::RemoveEmptyEntries)
    $folderName = $folderNames[0]
    if([string]::IsNullOrEmpty($folderName))
    {
        return;#root level, no folder to create.
    }
    #Write-Host "Creating folder [$folderName] ..."
    $curFolder = $parentFolder.Folders.Add($folderName)
    $Web.Context.Load($curFolder)
    $web.Context.ExecuteQuery()

    if ($folderNames.Length -gt 1)
    {
        $curFolderUrl = [System.String]::Join("/", $folderNames, 1, $folderNames.Length - 1)
        Create-Folder -Web $Web -ParentFolder $curFolder -FolderUrl $curFolderUrl
    }
}
Function ReplaceInFile($filePath, $textToFind, $textToReplace){
    (Get-Content $filePath | out-string) -ireplace $textToFind,$textToReplace | Set-Content $filePath
}

Function PublishFile($context, $web, $fileUrl, $versioningEnabled, $publishingEnabled)
{
    if($versioningEnabled -eq $false -and $publishingEnabled -eq $false)
    {
        return
    }
  $file = $web.GetFileByServerRelativeUrl($fileUrl)
  $context.Load($file)
  $context.ExecuteQuery();
  if($versioningEnabled -and $file.Level -ne "Published" -and $file.CheckOutType -ne [Microsoft.SharePoint.Client.CheckOutType]::None)
  {
            $file.CheckIn("Checked in by $productName", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
            Write-Host "$($file.Name) has been Checked In" 
    }
  if($publishingEnabled -and $file.Level -ne "Published")
  {
    $file.Publish("Approved by $productName")
    Write-Host "$($file.Name) has been Approved" 
  }
  $context.ExecuteQuery();
}
Function CheckOutFile($context, $web, $fileUrl)
{
  $file = $web.GetFileByServerRelativeUrl($fileUrl)
  $context.Load($file)
  $context.Load($file.ListItemAllFields)
  try {
      $context.ExecuteQuery();
      if($file.Exists -and $file.Level -eq "Checkout" -or $file.CheckOutType -ne [Microsoft.SharePoint.Client.CheckOutType]::None) 
      {
        $file.CheckIn("Checked in by $productName", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
        $context.ExecuteQuery();
      }
      if($file.Exists)
      {
        $file.CheckOut()
        $context.ExecuteQuery();
        Write-Host "$($file.Name) has been Checked Out" 
      }
    }
  catch
  {
    Write-Host $_
    #ignore this as the call to get the file triggers an error
    #todo: check if filenot found otherwise throw.
  }
}
function New-TempFile($fromFile)
{
    $tempFileName = [System.IO.Path]::GetTempFileName()
    Remove-Item $tempFileName -Force
    if($fromFile -ne $null)
    {
        Copy-Item $fromFile $tempFileName
    }
    return $tempFileName
}
function UploadFile([Microsoft.SharePoint.Client.ClientContext] $context, [Microsoft.SharePoint.Client.Web] $web,
    [System.IO.FileInfo]$file, [String] $fileUrl, [System.Xml.XmlElement] $properties, [Boolean] $versioningEnabled, [Boolean] $publishingEnabled, [Boolean] $skipIfExists = $false  )
{
    try {
        Write-Host "Uploading file $($file.Name) - $fileUrl ..."
    $fileExists = FileExists $context $web $fileUrl
    if($skipIfExists -and $fileExists)
    {
      Write-Host "$fileUrl exists, skipping"
      return
    }
        if($versioningEnabled -and $fileExists)
        {
            CheckOutFile $context $web $fileUrl
        }
        $stream = $null

        if(Is-TextFile $fileUrl)
        {
            #copy and modify the file and replace ~sitecollection with the real web server relative url.
            
            $tempFileName = New-TempFile $file.FullName
            DecodeCommonTokens $context $tempFileName
                        
            $tempFile = (Get-ChildItem $tempFileName)
            try 
            {
                $stream = $tempFile.OpenRead()
                [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($context, $fileUrl, $stream, $true)
            }
            finally {
                $stream.Dispose()
                Remove-Item $tempFileName -Force
            }
        }
        else {
            try 
            {
                $stream = $file.OpenRead()
                [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($context, $fileUrl, $stream, $true)
            }
            finally {
                $stream.Dispose()
            }
        }
		
        if($properties -ne $null)
        {
            $webFile = $web.GetFileByServerRelativeUrl($fileUrl)
            $context.Load($webFile)
            $context.Load($webFile.ListItemAllFields)
            $context.ExecuteQuery()
            foreach($attribute in $properties.Attributes)
            {
                if($webFile.ListItemAllFields.FieldValues.ContainsKey($attribute.name))
                {
                    $webFile.ListItemAllFields[$attribute.Name] = $attribute.Value
                }
            }
            $webFile.ListItemAllFields.Update()
            $context.ExecuteQuery()
        }
        PublishFile $context $web $fileUrl $versioningEnabled $publishingEnabled
    }
    catch {
       write-error "An error occured while uploading file [$($file.FullName)] $_"
    }
}
function FileExists([Microsoft.SharePoint.Client.ClientContext] $context, [Microsoft.SharePoint.Client.Web] $web, [String] $fileUrl)
    {
        try 
        {
            $webFile = $context.Site.RootWeb.GetFileByServerRelativeUrl($fileUrl)
            $context.Load($webFile)
            $context.ExecuteQuery()   
            return $true
        }
        catch
        {
            if($_.ToString().Contains("File Not Found"))
            {
                return $false
            }
            throw
        }
    }
function Test-SPFolder([Microsoft.SharePoint.Client.ClientContext] $context, [String] $folderUrl)
{
    try 
    {
        $folder = $web.GetFolderByServerRelativeUrl($folderUrl)
      $context.Load($folder)
      $context.ExecuteQuery()
        return $true
    }
    catch
    {
        if($_.ToString().Contains("Not Found"))
        {
            return $false
        }
        throw
    }
}
Function ClearDirectory([String] $path)
{
  if(Test-Path $path)
  {t
    Remove-Item -Path $path -Confirm:$false -Recurse:$true -Force
  }
  New-Item -ItemType directory -Path $path | out-null
}
Function EnsureDirectory([String] $path)
{
  if((Test-Path $path) -eq $false)
  {
    New-Item -ItemType directory -Path $path | out-null
  }
}	
	
Function ReplaceInFile($filePath, $textToFind, $textToReplace){
    
    $content = ((Get-Content $filePath | out-string) -ireplace $textToFind,$textToReplace)
    $content | Set-Content $filePath -force
    
    #DevNote: If you find the script erroring here it is most likely that visual studio is open and monitoring files and the file is locked.
    #todo: need a retry
}
Function DeleteFile($path)
{
  if(Test-Path $path)
  {
    Remove-Item -Path $path -Confirm:$false -Force
  }
}
Function GetLastModifiedBy($context, [Microsoft.SharePoint.Client.File] $file )
{
  $context.Load($file.ModifiedBy)
  $context.ExecuteQuery()
  return $file.ModifiedBy
}
Function LoadXmlFileFromSP($context, $web, $fileUrl)
{
  $file = $web.GetFileByServerRelativeUrl($fileUrl)
  $context.Load($file)
  $context.ExecuteQuery();
  if($file.Exists -eq $false)
  {
    throw "File does not exist $fileUrl"
  }
  try {
        if($context.HasPendingRequest)
        {
          $context.ExecuteQuery()
        }
        $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($context, $fileUrl)
        [xml] $doc = New-Object System.Xml.XmlDocument
        $doc.Load($fileInfo.Stream)
        return $doc
  }
  finally {
    if($fileInfo -ne $null -and $fileInfo.Stream -ne $null)
    {
      $fileInfo.Stream.Dispose()
    }
		
  }

}	

function Set-NewExperience{
    <#
      .Synopsis
       Sets the document library experience for a site or web
      .DESCRIPTION
       Sets the document library experience for a site or web
      .EXAMPLE
       The following would disable the new experience for an entire site collection
       Set-NewExperience -Url "https://tenant.sharepoint.com/sites/somesite" -Scope Site -State Disabled
      .EXAMPLE
       The following would disable the new experience for a single web
       Set-NewExperience -Url "https://tenant.sharepoint.com/sites/somesite" -Scope Web -State Disabled
      .EXAMPLE
       The following would enable the new experience for an entire site collection
       Set-NewExperience -Url "https://tenant.sharepoint.com/sites/somesite" -Scope Site -State Enabled -Context $clientContext
      .Link
      https://support.office.com/en-us/article/Switch-the-default-for-document-libraries-from-new-or-classic-66dac24b-4177-4775-bf50-3d267318caa9
    #>
    Param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)][ValidateSet("Site", "Web")]
        [string]$Scope,
        [Parameter(Mandatory=$true)][ValidateSet("Enabled", "Disabled")]
        [string]$State,
    [Parameter(Mandatory=$true)]
        [Microsoft.SharePoint.Client.ClientContext]$context
    )

    Process{
        if($Scope -eq "Site"){
            $site = $context.Site
            $featureguid = new-object System.Guid "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4"
        }
        else{
            $site = $context.Web
            $featureguid = new-object System.Guid "52E14B6F-B1BB-4969-B89B-C4FAA56745EF" 
        }
        if($State -eq "Disabled")
        {
            $site.Features.Add($featureguid, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
            $message = "New library experience has been disabled on $URL"
        }
        else{
            $site.Features.Remove($featureguid, $true)
            $message = "New library experience has been enabled on $URL"
        }
        try{
            $context.ExecuteQuery()
            write-host -ForegroundColor Green $message
        }
        catch{
            Write-Host -ForegroundColor Red $_.Exception.Message
        }
    }
    
}
Function Combine-Url([string] $path1, [string] $path2, [string] $path3)
{
  $path1 = $path1 + ""
  $path2 = $path2 + ""
  if([String]::IsNullOrEmpty($path1) -eq $false-and $path1.EndsWith("/"))
  {
    $path1 = $path1.TrimEnd("/")
  }
  if([String]::IsNullOrEmpty($path2) -eq $false -and $path2.EndsWith("/"))
  {
    $path2 = $path2.TrimEnd("/")
  }
  if([String]::IsNullOrEmpty($path3) -eq $false -and $path3.EndsWith("/"))
  {
    $path3 = $path3.TrimEnd("/")
  }
  if([String]::IsNullOrEmpty($path2) -eq $false -and $path2.StartsWith("/") -eq $false)
  {
    $path2 = "/" + $path2
  }
  if([String]::IsNullOrEmpty($path3) -eq $false -and $path3.StartsWith("/") -eq $false)
  {
    $path3 = "/" + $path3
  }
  return ($path1 + $path2 + $path3)
}

function Get-List($context, $name)
{
  [Microsoft.SharePoint.Client.List] $list = $context.Web.Lists.GetByTitle($name);
  $context.Load($list)
  $context.Load($list.Fields)
  $context.ExecuteQuery()
  return $list
}
function Get-ContentType([Microsoft.SharePoint.Client.ClientContext]$context, [string] $name)
{
  $cts = $context.Site.RootWeb.AvailableContentTypes;
  $context.Load($cts)
  $context.ExecuteQuery()
  $ct = $cts | Where-Object {$_.Name -ieq $name}
  if($ct -ne $null)
  {
    $context.Load($ct[0])
    $context.Load($ct[0].FieldLinks)
    $context.Load($ct[0].Fields)
    $context.ExecuteQuery()
    return $ct[0]
  }
  write-warning "$name not found"
  return $null
}
Function Load-SiteManifest($destination)
{
    #Extra settings for a site build are specified here, such as lists to grab.
    $xmlFilePath = [System.IO.Path]::Combine($destination, "manifest.xml")
    if((Test-Path $xmlFilePath) -eq $false)
    {
        write-host "No Site Manifest"
        return $null
    }
    [xml]$settingsFile = Get-Content $xmlFilePath
    return $settingsFile
}

Function Ensure-Directory($path)
{
  if((Test-Path $path) -eq $false)
    {
        New-Item -path $path -type:Directory	
    }
}
Function Convert-ContentTypeIdToArray([string]$contentTypeId)
{
    if([string]::IsNullOrEmpty($contentTypeId))
    {
        throw "Invalid content type id"
    }
    $ct = @()
    $index = $contentTypeId.IndexOf("00")
    $ct += $contentTypeId.Substring(0,$index); #Primary top level parent
    for($i = $index + 2; $i -lt $contentTypeId.Length; $i)
    {
        $block = [Math]::Min($contentTypeId.Length - $i,32)
        $id = $contentTypeId.Substring($i, $block)
        $ct += $id
        $i += $block
        $i = $i + 2
    }
    return $ct
}
Function Get-ParentContentTypeId([string]$contentTypeId)
{
    $ids = Convert-ContentTypeIdToArray $contentTypeId
    if($ids.Length -eq 1)
    {
        Write-Warning "There is no parent for this content type id $contentTypeId"
        return $ids[0]
    }
    if($ids.Length -gt 1)
    {
        return ($ids[0..($ids.Length - 2)]) -join '00'
        
    }
    throw "Invalid content type id"
}
Function Strip-Start([string]$stringIn, [string] $toRemove)
{
    $i = $stringIn.IndexOf($toRemove)
    if($i -ge 0)
    {
        return $stringIn.Substring($i + $toRemove.Length)
    }
    return $stringIn
}
Function Is-TextFile([string]$filename)
{
    $extension = [IO.Path]::GetExtension($filename).Replace('.','').ToLower()
    $textFiles = "txt","css","master","html","aspx","js","ts","htm","xml","json","dwp","webpart"
    return ($textFiles -icontains $extension)
}
function Convert-FileTo2013($fullFilePath)
{
    write-host "Generating 2013 compatible file"
    ReplaceInFile $fullFilePath "Version=16.0.0.0" "Version=15.0.0.0"
    ReplaceInFile $fullFilePath "<SharePoint:IECompatibleMetaTag runat=`"server`" />" ""
    Write-Host "Converted $fullFilePath for use with SP 2013"
}
function Convert-FilesTo2013($path)
{
    $files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
    foreach($file in $files)
    {
        Convert-FileTo2013 $file.Path
    }
}
Function EncodeCommonTokens($context, $fileName)
{
    if($context.Site.ServerRelativeUrl.Length -le 1)
    {
        Write-Warning "Warning, you are building from a root site.  This means we can't correctly detect the site collection urls that exist in some files.  You should build from a /sites/somesite location."
    }
    $siteServerRelativeUrl = $context.Site.ServerRelativeUrl
    ReplaceInFile $fileName $siteServerRelativeUrl "[~sitecollection]"
} 
Function DecodeCommonTokens($context, $fileName)
{
    if($context.Site.ServerRelativeUrl.Length -le 1)
    {
        ReplaceInFile $tempFileName "\[~sitecollection\]" "" #top level gets replaced to empty string.
    }
    else {
        ReplaceInFile $tempFileName "\[~sitecollection\]" $context.Site.ServerRelativeUrl 
    }
} 
Export-ModuleMember -Function '*'