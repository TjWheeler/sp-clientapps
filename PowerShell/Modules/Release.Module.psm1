<# This module is part of the automated release process for SharePoint solutions.  It uses the Client Side Object Model.
#  Author: Tim Wheeler (http://blog.timwheeler.tech/)
#  Attribution: Some functions and approaches have been taken (and may have been updated) from other sources online.  
#  Unfortunately I have not kept those links.  If you know of the source locations please let me know and I'll happilly attribute. 
#> 

$productName =" sp-clientapps automated release process"

Function Install-DisplayTemplates($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "DisplayTemplates")
	$xmlFilePath = [System.IO.Path]::Combine($path, "DisplayTemplates.xml")
    if((Test-Path $path) -eq $false -or (Test-Path $xmlFilePath) -eq $false)
    {
      write-host "No Display Templates to deploy"
      return
    }    

	$list = $context.Site.GetCatalog([Microsoft.SharePoint.Client.ListTemplateType]::MasterPageCatalog)
	$context.Load($list)
    $context.Load($list.RootFolder)

	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
	
	[xml]$settingsFile = Get-Content $xmlFilePath
	$files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
	foreach($file in $files)
	{
		if($file.Name -eq "DisplayTemplates.xml" -or $file.Extension -eq ".js")
		{
			continue
		}
		Write-Host "Uploading Display Template $($file.Name)"
		$properties = $settingsFile.xml.DisplayTemplates.DisplayTemplate | where-object { $_.FileName -eq $file.Name}
        $folderUrl = "/Display Templates" + ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl = ($folderUrl + "/" + $file.Name)
        if([string]::IsNullOrEmpty($folderUrl) -eq $false)
        {
		    Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl | out-null
        }
        $fullFileUrl = Combine-Url $list.RootFolder.ServerRelativeUrl $fileUrl
        $properties.Attributes.RemoveNamedItem("FileName") | out-null
        UploadFile $context $web $file $fullFileUrl $properties $versioningEnabled $publishingEnabled | out-null
	}
}
Function Install-WebParts($context, $web, $location, $csomVersion)
{
	$path = [System.IO.Path]::Combine($location, "WebParts")
    if((Test-Path $path) -eq $false )
    {
        write-host "No Web Parts to deploy"
        return
    }

	$list = $context.Site.GetCatalog([Microsoft.SharePoint.Client.ListTemplateType]::WebPartCatalog)
	$context.Load($list)
    $context.Load($list.RootFolder)

	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
	
	$files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
	foreach($file in $files)
	{
		Write-Host "Uploading Web Part $($file.Name)"
        if($context.HasPendingRequest)
        {
            $context.ExecuteQuery()
        }
        $folderUrl =  ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl = ($folderUrl + "/" + $file.Name)
        $fullFileUrl = $list.RootFolder.ServerRelativeUrl + $fileUrl
        [xml]$properties = '<xml><WebPart Group="Client Apps"/></xml>' 
		$replaceWith = $web.ServerRelativeUrl
        if($replaceWith -eq "/")
		{
			$replaceWith = "";
		}
        if($csomVersion -eq "15")
        {
            $tempFileName = New-TempFile $file.FullName
            Convert-FileTo2013 $tempFileName
            UploadFile $context $web (Get-ChildItem $tempFileName) $fullFileUrl $properties.xml.WebPart $versioningEnabled $publishingEnabled | out-null
            Remove-Item $tempFileName -Force
        }
        else 
        {
            UploadFile $context $web $file $fullFileUrl $properties.xml.WebPart $versioningEnabled $publishingEnabled | out-null
        }   
	}
}
Function Install-StyleLibrary($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "Style Library")
  if((Test-Path $path) -eq $false)
  {
      write-host "No Style Library assets to deploy"
      return
  }
	$list = $web.Lists.GetByTitle("Style Library");
	$context.Load($list)
  $context.Load($list.RootFolder)
	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
  $files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
  foreach($file in $files)
	{
		$folderUrl =  ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl = ($folderUrl + "/" + $file.Name)
		if([string]::IsNullOrEmpty($folderUrl) -eq $false)
		{
			Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl | out-null
		}
        $fullFileUrl = $list.RootFolder.ServerRelativeUrl + $fileUrl
        UploadFile $context $web $file $fullFileUrl $null $versioningEnabled $publishingEnabled | out-null
	}
}
Function Install-CommonFiles($context, $web, $location)
{
    throw "Obsolete - use the common client app"
    Write-Host "Deploying common files to Style Library"
	$path = $location
      if((Test-Path $path) -eq $false)
      {
          write-host "No Style Library\Client Apps\Common assets to deploy"
          return
      }
      
	$list = $web.Lists.GetByTitle("Style Library");
	$context.Load($list)
  $context.Load($list.RootFolder)
	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
  $files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
  $commonFolderUrl = "Client Apps/Common"
  Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $commonFolderUrl | out-null
  foreach($file in $files)
	{
		$folderUrl =  ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl = ($folderUrl + "/" + $file.Name)
        $folderUrl = Combine-Url $commonFolderUrl $folderUrl
		if([string]::IsNullOrEmpty($folderUrl) -eq $false)
		{
			Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl | out-null
		}
        $fullFileUrl = Combine-Url $list.RootFolder.ServerRelativeUrl $commonFolderUrl $fileUrl
        UploadFile $context $web $file $fullFileUrl $null $versioningEnabled $publishingEnabled | out-null
	}
}
Function Install-PublishingImages($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "PublishingImages")
    if((Test-Path $path) -eq $false)
    {
        write-host "No Publishing to deploy"
        return
    }
	$list = $web.Lists.GetByTitle("Images");
	$context.Load($list)
	$context.Load($list.RootFolder)
	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning

    $files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
    foreach($file in $files)
	{
		$folderUrl = $file.SubFolder #  ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl =  ($folderUrl + "/" + $file.Name)
		if([string]::IsNullOrEmpty($folderUrl) -eq $false)
		{
			Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl
		}
        $fullFileUrl = $list.RootFolder.ServerRelativeUrl + $fileUrl
        UploadFile $context $web $file $fullFileUrl $null $versioningEnabled $publishingEnabled | out-null
	}
}
#This creates the lists only.  not fields or cts 
Function Install-Lists($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "Lists\Lists.xml")
    if((Test-Path $path) -eq $false)
    {
        Write-Host "No lists to deploy"
        return 
    }
    $xmlFilePath = [System.IO.Path]::Combine($location, "Lists\Lists.xml")
    [xml] $listXml = Get-Content $xmlFilePath
    $lists = $listXml.xml.Lists.List 
    $rootWeb = $context.Site.RootWeb
    $context.Load($web.Lists)
    $context.Load($rootWeb.Lists)
    $context.ExecuteQuery();
    foreach($listDef in $lists)
    {
        $targetWeb = $web
        if(Is-True $listDef.rootWeb)
        {
            $targetWeb = $rootWeb
        }
        else {
            $targetWeb = $web
        }
            
        $list = $targetWeb.Lists | Where-Object { $_.Title -eq $listDef.Title}
        
        if($list -eq $null -or $list.Count -eq 0)
        {
            $listInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
            $listInfo.Title = $listDef.Title
            $listInfo.TemplateType = $listDef.TemplateType
            $list = $targetWeb.Lists.Add($listInfo)
            $list.Description = $listDef.Description
            $list.Update()
            write-host "Creating List {$list.Title}"
            $Context.ExecuteQuery()
        }
    }
}

#Provision the list, add the list content types and reorder CT's.  Last CT added wins.
Function Install-ConfigureLists($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "Lists\Lists.xml")
    if((Test-Path $path) -eq $false)
    {
        Write-Host "No lists to deploy"
        return 
    }
    $xmlFilePath = [System.IO.Path]::Combine($location, "Lists\Lists.xml")
    [xml] $listXml = Get-Content $xmlFilePath
    $lists = $listXml.xml.Lists.List 
    $rootWeb = $context.Site.RootWeb
    $context.Load($web.Lists)
    $context.Load($rootWeb.Lists)
    $context.ExecuteQuery();
    foreach($listDef in $lists)
    {
        $targetWeb = $web
        if(Is-True $listDef.rootWeb)
        {
            $targetWeb = $rootWeb
        }
        else {
            $targetWeb = $web
        }
            
        $list = $targetWeb.Lists | Where-Object { $_.Title -eq $listDef.Title}
        
        if($list -eq $null -or $list.Count -eq 0)
        {
            $listInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
            $listInfo.Title = $listDef.Title
            $listInfo.TemplateType = $listDef.TemplateType
            $list = $targetWeb.Lists.Add($listInfo)
            $list.Description = $listDef.Description
            $list.Update()
            $Context.ExecuteQuery()
        }
        $contentTypes = $list.ContentTypes
        $context.Load($contentTypes)
        $context.Load($list.RootFolder)
        $context.Load($list.ContentTypes)
        $context.ExecuteQuery()
        $ctNames =  $contentTypes | Select -ExpandProperty Name
        foreach($contentTypeDef in $listDef.ContentTypes.ContentType)
        {
            if($contentTypeDef.Name -eq "Item" -or $contentTypeDef.Name -eq "Document" -or $contentTypeDef.Name -eq "Folder")
            {
                continue; #causes issues trying to re add them.
            }
            #Dev Note: The list ct id is not part of the web cts, so we need to look for the parent id only.
            $parentContentTypeId = Get-ParentContentTypeId $contentTypeDef.ID
            $contentType = $contentTypes | Where-Object { 
                $_.ID.StringValue.Length -ge $parentContentTypeId.Length -and
                $parentContentTypeId -eq $_.ID.StringValue.SubString(0,$parentContentTypeId.Length)
            }
            if($contentType -eq $null -or $contentType.Count -eq 0)
            {
				Write-host "Looking for parent Content Type $parentContentTypeId ($($contentTypeDef.Name))"
                try {
                    $contentType = $context.Site.RootWeb.ContentTypes.GetById($parentContentTypeId)
                    $context.Load($contentType)
                    $Context.ExecuteQuery()
                }
                catch
                {
                    Write-Error "Could not load parent Content Type $parentContentTypeId ($($contentTypeDef.Name)), ensure it exists in the root web"
                    continue;
                }
                if($contentType -eq $null -or $contentType.Id -eq $null)
                {
                    throw "Could not find content type $($contentTypeDef.Name)"
                    continue;
                }
                #look for CT in list, if it exists ignore.
                if($ctNames.Contains($contentTypeDef.Name))
                {
                    write-host "Content Type $($contentType.Name) already exists in list $($list.Title)"
                    continue
                }
                write-host "Adding CT $($contentTypeDef.Name) to list $($listDef.Title)"

                $list.ContentTypesEnabled = $true
                $listContentType = $list.ContentTypes.AddExistingContentType($contentType) 
                $list.Update()
                $context.Load($list)
                $context.Load($list.RootFolder)
                $context.Load($listContentType)
                $context.ExecuteQuery()
                #DevNote: You must commit these changes and reload before trying to reorder.
                $ctOrdered = New-Object System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]
                $ctOrdered.Add($listContentType.Id)
                foreach($ct in $list.ContentTypes)
                {
                    if($ct -eq $listContentType -or $ct.Name -eq "Folder")
                    {
                        continue;
                    }
                    $ctOrdered.Add($ct.Id)
                }
                $list.RootFolder.UniqueContentTypeOrder = $ctOrdered
                $list.Update()
                $context.Load($list)
                $context.ExecuteQuery()
            }
        }
    }
}

#if null or empty return the default.
function GetWithDefault($value, $default){
    if($value -eq $null)
    {
        return $default
    }
    if([string]::IsNullOrEmpty($value))
    {
        return $default
    }
    return $value
}
#Uses a SiteFields.xml file to deploy fields if they dont exist
Function Install-SiteFields($context, $web, $location)
{
    $path = [System.IO.Path]::Combine($location, "SiteFields")
	$xmlFilePath = [System.IO.Path]::Combine($path, "SiteFields.xml")
    if((Test-Path $path) -eq $false -or (Test-Path $xmlFilePath) -eq $false)
    {
        write-host "No Site Fields to deploy"
        return
    }

    # Handle Lookup fields
    $listManifestPath = [System.IO.Path]::Combine($location, "Lists")
	$listXmlFilePath = [System.IO.Path]::Combine($listManifestPath, "Lists.xml")
    $lists = @{}
    if((Test-Path $listManifestPath) -eq $false -or (Test-Path $listXmlFilePath) -eq $false)
    {
        write-warning "No List manifest could be found, Choice columns may not be accurate"
    }
    else {
        #load the list details
        [xml]$listsXML = Get-Content $listXmlFilePath
        foreach($listDef in $listsXML.xml.Lists.List)
        {
            $list = $web.Lists.GetByTitle($listDef.Title)
            $context.Load($list)
            if($listDef.Id -ne $null)
            {
                $lists.Add($listDef.Id, $list);
            }
        }
        if($lists.Count -gt 0) {
            $context.ExecuteQuery();
        }
    }
	
    [xml]$fieldsXML = Get-Content $xmlFilePath

    $context.Load($web.Fields)
    $context.ExecuteQuery()

    $fieldsXML.xml.Fields.Field | ForEach-Object {
    if($_ -eq $null)
    {
        return
    }
    #Configure core properties belonging to all column types
    $fieldXML = '<Field Type="' + $_.Type + '"
    Name="' + $_.Name + '"
    ID="' + $_.ID + '"
    Description="' + $_.Description + '"
    DisplayName="' + $_.DisplayName + '"
    StaticName="' + $_.StaticName + '"
    Group="' + $_.Group + '"
    Hidden="' + (GetWithDefault $_.Hidden "FALSE") + '"
    Required="' + $_.Required + '"
    Sealed="' + (GetWithDefault $_.Sealed, "FALSE") + '"'
    
    #Configure optional properties belonging to specific column types - you may need to add some extra properties here if present in your XML file
    if ($_.ShowInDisplayForm) { $fieldXML = $fieldXML + "`n" + 'ShowInDisplayForm="' + $_.ShowInDisplayForm + '"'}
    if ($_.ShowInEditForm) { $fieldXML = $fieldXML + "`n" + 'ShowInEditForm="' + $_.ShowInEditForm + '"'}
    if ($_.ShowInListSettings) { $fieldXML = $fieldXML + "`n" + 'ShowInListSettings="' + $_.ShowInListSettings + '"'}
    if ($_.ShowInNewForm) { $fieldXML = $fieldXML + "`n" + 'ShowInNewForm="' + $_.ShowInNewForm + '"'}
        
    if ($_.EnforceUniqueValues) { $fieldXML = $fieldXML + "`n" + 'EnforceUniqueValues="' + $_.EnforceUniqueValues + '"'}
    if ($_.Indexed) { $fieldXML = $fieldXML + "`n" + 'Indexed="' + $_.Indexed + '"'}
    if ($_.Format) { $fieldXML = $fieldXML + "`n" + 'Format="' + $_.Format + '"'}
    if ($_.MaxLength) { $fieldXML = $fieldXML + "`n" + 'MaxLength="' + $_.MaxLength + '"' }
    if ($_.FillInChoice) { $fieldXML = $fieldXML + "`n" + 'FillInChoice="' + $_.FillInChoice + '"' }
    if ($_.NumLines) { $fieldXML = $fieldXML + "`n" + 'NumLines="' + $_.NumLines + '"' }
    if ($_.RichText) { $fieldXML = $fieldXML + "`n" + 'RichText="' + $_.RichText + '"' }
    if ($_.RichTextMode) { $fieldXML = $fieldXML + "`n" + 'RichTextMode="' + $_.RichTextMode + '"' }
    if ($_.IsolateStyles) { $fieldXML = $fieldXML + "`n" + 'IsolateStyles="' + $_.IsolateStyles + '"' }
    if ($_.AppendOnly) { $fieldXML = $fieldXML + "`n" + 'AppendOnly="' + $_.AppendOnly + '"' }
    if ($_.Sortable) { $fieldXML = $fieldXML + "`n" + 'Sortable="' + $_.Sortable + '"' }
    if ($_.RestrictedMode) { $fieldXML = $fieldXML + "`n" + 'RestrictedMode="' + $_.RestrictedMode + '"' }
    if ($_.UnlimitedLengthInDocumentLibrary) { $fieldXML = $fieldXML + "`n" + 'UnlimitedLengthInDocumentLibrary="' + $_.UnlimitedLengthInDocumentLibrary + '"' }
    if ($_.CanToggleHidden) { $fieldXML = $fieldXML + "`n" + 'CanToggleHidden="' + $_.CanToggleHidden + '"' }
    if ($_.List) { $fieldXML = $fieldXML + "`n" + 'List="' + $_.List + '"' }
    if ($_.ShowField) { $fieldXML = $fieldXML + "`n" + 'ShowField="' + $_.ShowField + '"' }
    if ($_.UserSelectionMode) { $fieldXML = $fieldXML + "`n" + 'UserSelectionMode="' + $_.UserSelectionMode + '"' }
    if ($_.UserSelectionScope) { $fieldXML = $fieldXML + "`n" + 'UserSelectionScope="' + $_.UserSelectionScope + '"' }
    if ($_.BaseType) { $fieldXML = $fieldXML + "`n" + 'BaseType="' + $_.BaseType + '"' }
    if ($_.Mult) { $fieldXML = $fieldXML + "`n" + 'Mult="' + $_.Mult + '"' }
    if ($_.ReadOnly) { $fieldXML = $fieldXML + "`n" + 'ReadOnly="' + $_.ReadOnly + '"' }
    if ($_.FieldRef) { $fieldXML = $fieldXML + "`n" + 'FieldRef="' + $_.FieldRef + '"' }    

    #Update Lookup Field Id's
    if ($_.Type -eq "Lookup") {
        #This will replace the List="[Guid]" with the current lists id as these change in new environments.
        foreach($mapping in $lists.Keys){
            $fieldXML = $fieldXML -ireplace $mapping, $lists[$mapping].Id
        }
        $fieldXML = $fieldXML + " WebId=`"" + $web.Id  + "`""
    }

    $fieldXML = $fieldXML + ">"
    #Create choices if choice column
    if ($_.Type -eq "Choice") {
        $fieldXML = $fieldXML + "`n<CHOICES>"
        $_.Choices.Choice | ForEach-Object {
           $fieldXML = $fieldXML + "`n<CHOICE>" + $_ + "</CHOICE>"
        }
        $fieldXML = $fieldXML + "`n</CHOICES>"
    }
    
    #Set Default value, if specified  
    if ($_.Default) { $fieldXML = $fieldXML + "`n<Default>" + $_.Default + "</Default>" }
    
    #End XML tag specified for this field
    $fieldXML = $fieldXML + "</Field>"
    
    $iterator = $_
    $field = $web.Fields | Where-Object { $_.ID -icontains $iterator.ID } 
    
    if($field -eq $null)
    {
        $web.Fields.AddFieldAsXml($fieldXML.Replace("&","&amp;"), $false, [Microsoft.SharePoint.Client.AddFieldOptions]::AddToAllContentTypes) | out-null
        write-host "Creating site column" $_.DisplayName "on" $web.Url
    }
    else {
        $field.SchemaXml = $fieldXML.Replace("&","&amp;")
        $field.Update()
        write-host "Updating site column" $_.DisplayName "on" $web.Url
    }
    $context.ExecuteQuery()
    }

}
# Creates content types and field links if they don't exist
#This method may generate errors if you haven't aquired all fields in the Build process.
Function Install-ContentTypes($context, $web, $location)
{
    
    $path = [System.IO.Path]::Combine($location, "ContentTypes")
	$xmlFilePath = [System.IO.Path]::Combine($path, "ContentTypes.xml")
    if((Test-Path $path) -eq $false -or (Test-Path $xmlFilePath) -eq $false)
    {
        write-host "No Site Fields to deploy"
        return
    }
    [xml]$ctsXML = Get-Content $xmlFilePath
    $context.Load($web.Fields)
    $context.Load($web.AvailableContentTypes)
    $context.ExecuteQuery()


    $ctsXML.xml.ContentTypes.ContentType | ForEach-Object {
    
    if($_ -eq $null)
    {
        return
    }
    $iterator = $_
    write-host ("Processing Content Type : " + $_.Name)
    $spContentType = $web.AvailableContentTypes | Where {$_.Name -eq $iterator.Name}
    $requiresUpdate = $false
    if($spContentType -eq $null)
    {
        #Create Content Type object inheriting from parent
        [Microsoft.SharePoint.Client.ContentTypeCreationInformation] $spContentTypeCI = new-object Microsoft.SharePoint.Client.ContentTypeCreationInformation
        $spContentTypeCI.Id = $_.ID
        $spContentTypeCI.Name = $_.Name
        #$spContentType.ParentContentType = $parentContentType
        $spContentTypeCI.Group = $_.Group
        $spContentType = $web.ContentTypes.Add($spContentTypeCI)
        $spContentType.Description = $_.Description
        $spContentType.Group = $_.Group
        $requiresUpdate = $true
    }

    #Set Content Type description and group

    $context.Load($spContentType.FieldLinks)
    $context.Load($spContentType.Fields)
    $context.ExecuteQuery()
    
    write-host "Adding field links"
    foreach($fieldDef in $_.Fields.Field) {
        if($fieldDef.DisplayName.contains(":")) 
        {
            #these fields must be skipped (The choice additional fields automatically added).
            continue;
        }
        $doesFieldExist = $false
        foreach($fld in $spContentType.Fields)
        {
            if($fld.get_InternalName() -eq $fieldDef.Name){
               $doesFieldExist = $true
               break
            }
        }
        if($doesFieldExist -eq $false)
            {
                #Get all the columns associated with content type
                $fieldRefCollection = $spContentType.FieldLinks;
                
                #Add a new column to the content type
                $fieldName = $fieldDef.Name
                $field = $web.Fields | Where-Object { $_.InternalName -eq $fieldName }
                if($field -eq $null)
                {
                    Write-Error "The field $fieldName is not found in the Site Fields.  This means the field was not aquired through the build process. The group is $($fieldDef.Group)." -ErrorAction:Continue
                    continue;
                }
                
                Write-Host "Adding site column " $fieldDef.Name "to the content type" -ForegroundColor Cyan
                $fieldReferenceLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
                $fieldReferenceLink.Field = $field;
                $fieldLink = $spContentType.FieldLinks.Add($fieldReferenceLink);
                $fieldLink.Required =  ($fieldDef.Required -ieq "TRUE")
                $fieldLink.Hidden =  ($fieldDef.Hidden -ieq "TRUE")

                $requiresUpdate = $true
                Write-Host "Field added successfully" -ForegroundColor Green
            }
       }
    
    #Create Content Type on the site and update Content Type object
    if($requiresUpdate)
    {
        $spContentType.Update($true)
        $context.ExecuteQuery()
        write-host "Content type" $fieldDef.Name "has been create/updated" -ForegroundColor Green
    }
    else {
        write-host "Content type" $fieldDef.Name "exists and has no updates"
    }
  }
}

#Creates page layouts.  Uses a PageLayouts.xml to get metadata.
#layout is uploaded and metadata set. $subFolderPath is the sub directory in the master page library and should the be project name.
Function Install-PageLayouts($context, $web, $location, $subFolderPath, $csomVersion)
{
  $path = [System.IO.Path]::Combine($location, "PageLayouts")
  $xmlFilePath = [System.IO.Path]::Combine($path, "PageLayouts.xml")
    if((Test-Path $path) -eq $false -or (Test-Path $xmlFilePath) -eq $false)
    {
        write-host "No Site Fields to deploy"
        return
    }

  $masterPageUrl = Combine-Url $web.ServerRelativeUrl "_catalogs" "masterpage"
  $list = $context.Site.GetCatalog([Microsoft.SharePoint.Client.ListTemplateType]::MasterPageCatalog)
  $context.Load($list)
  $context.Load($list.RootFolder)

  $context.ExecuteQuery();
  $publishingEnabled = $list.EnableModeration
  $versioningEnabled = $list.EnableVersioning
	
  [xml]$settingsFile = Get-Content $xmlFilePath
  $files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false} # -and $_.Name -eq "FileName.aspx"
  foreach($file in $files)
  {
    if($file.Name -eq "PageLayouts.xml")
    {
      continue
    }
    Write-Host "Uploading Page Layout $($file.Name)"
    $properties = $settingsFile.xml.Layouts.Layout | where-object { $_.FileName -eq $file.Name}
    #$folderUrl =   ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
    $folderUrl =  $subFolderPath 
    $fileUrl = Combine-Url $folderUrl $file.Name
    if([string]::IsNullOrEmpty($folderUrl) -eq $false)
    {
        Ensure-Folder -Web $web -ParentFolder $list.RootFolder $folderUrl | Out-Null
    }
    #$fullFileUrl = Combine-Url $folderUrl $fileUrl
    $properties.Attributes.RemoveNamedItem("FileName") #FileName is added in in the build process and is not part of the sharepoint schema.
    $properties.Attributes.RemoveNamedItem("_VirusStatus")
    $fullFileUrl = Combine-Url $masterPageUrl $fileUrl
    if($csomVersion -eq "15")
    {
        $tempFileName = New-TempFile $file.FullName
        Convert-FileTo2013 $tempFileName
        UploadFile $context $web (get-item $tempFileName) $fullFileUrl $properties $versioningEnabled $publishingEnabled | out-null
        Remove-Item $tempFileName -Force
    }
    else 
    {
        UploadFile $context $web $file $fullFileUrl $properties $versioningEnabled $publishingEnabled | out-null
    }
  }
}

Function Install-ClientApp($context, $web, $rootWeb, $location, $manifest, $csomVersion)
{
    $xmlFilePath = [System.IO.Path]::Combine($location, "manifest.xml")
    if((Test-Path $xmlFilePath) -eq $false)
    {
        Write-Host "No manifest.xml file found at the location $location"
        return 
    }
    write-host "Deploying Client App - $($manifest.xml.ClientApp.Name)"
    Install-WebParts $context $rootWeb $location $csomVersion
    Install-StyleLibrary $context $rootWeb $location
    Install-Lists $context $web $location
    Install-SiteFields $context $rootWeb $location
    Install-ContentTypes $context $rootWeb $location
    Install-ConfigureLists $context $web $location
    Install-ListItems $context $web $location
    Install-DisplayTemplates $context $rootWeb $location
}
Function Install-ClientApps($context, $web, $location, $csomVersion)
{
	$path = [System.IO.Path]::Combine($location, "Client Apps")
    if((Test-Path $path) -eq $false)
    {
        Write-Host "No Client Apps"
        return 
    }
    
    $webPartFolders = Get-ChildItem $path -Directory
    $rootWeb = $context.Site.RootWeb
    $context.Load($rootWeb)
    $context.ExecuteQuery()
    foreach($webPartPath in $webPartFolders)
    {
        $xmlFilePath = [System.IO.Path]::Combine($webPartPath.FullName, "manifest.xml")
        if((Test-Path $xmlFilePath) -eq $false)
        {
            Write-warning "No manifest.xml file for client app found at the location $xmlFilePath"
            continue
        }
        [xml]$manifest = get-content $xmlFilePath
        Install-ClientApp $context $web $rootWeb $webPartPath.FullName $manifest $csomVersion
    }
}
Function Activate-WebFeature($context, $web, [guid]$featureId, $name)
{
    #note: This may only work in SP Online.  To verify.
    $features = $web.Features
    $context.Load($features)
    $context.ExecuteQuery()
    $feature = $features | Where-Object { $_.DefinitionId -eq $featureId}
    if($feature -eq $null -or $feature.Length -eq 0)
    {
        Write-Host "Adding feature $name"
        $features.Add($featureId, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
        try {
            $context.ExecuteQuery()
        }
        catch
        {
            Write-Warning "The feature $name with id $featureId could not be activated.  This may be an issue with the CSOM Version (Client or Server). "
            Write-Error $_            
        }
    }
    else {
        Write-Host "The feature $name is already activated"
    }
}
Function Activate-SiteFeature($context, [guid]$featureId, $name)
{
    #note: This may only work in SP Online.  To verify.
    $features = $context.Site.Features
    $context.Load($features)
    $context.ExecuteQuery()
    $feature = $features | Where-Object { $_.DefinitionId -eq $featureId}
    if($feature -eq $null -or $feature.Length -eq 0)
    {
        Write-Host "Adding feature $name"
        $features.Add($featureId, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
        try {
            $context.ExecuteQuery()
        }
        catch
        {
            Write-Warning "The feature $name with id $featureId could not be activated.  This may be an issue with the CSOM Version (Client or Server). "
            Write-Error $_            
        }
    }
    else {
        Write-Host "The feature $name is already activated"
    }
}
Function Install-Webs ($context, $web, $location)
{
  $websPath = [IO.Path]::Combine($location, "Webs")
  if(Test-Path $websPath)
  {
      #$rootWebPath = [IO.Path]::Combine($websPath, "Root")
      #Install-Web $context $web $rootWebPath
      $webPaths = Get-ChildItem $websPath  | Where-Object { $_.PSIsContainer -eq $true}
      foreach($webPath in $webPaths)
      {
         Install-Web $context $web $webPath.FullName
      }

  }
}

Function Install-Web ($context, $parentWeb, $location)
{
    $webManifestPath = [IO.Path]::Combine($location, "manifest.xml")
    [xml]$webManifest = Get-Content $webManifestPath
    #Build the structure of the web based on the manifest
    if($webManifest -eq $null)
    {
      write-host "No web manifest.  Skipping Web deploy"
      return
    }
    $web = $parentWeb
    $isRootWeb = $location.EndsWith("Root")
    if($isRootWeb -eq $false)
    {
        #Look to see if web exists
        $webs = $parentWeb.Webs
        $context.Load($webs)
        $context.ExecuteQuery()
        $existing = $webs | Where-Object {$_.Title -ieq $webManifest.xml.Web.Title}
        if($existing -ne $null -and $existing.Length -ge 1)
        {
            Write-Host "Web $($webManifest.xml.Web.Title) already exists"
            $web = $existing
        }
        else {
            $webProperties = $webManifest.xml.Web
            $title = $webProperties.Title
            $webUrl = $webProperties.ServerRelativeUrl.TrimStart("/")
            $template = $webProperties.WebTemplate
            $configuration = $webProperties.Configuration
            Write-Host "Creating Web $webUrl" #TODO
            $createWeb = New-Object Microsoft.SharePoint.Client.WebCreationInformation
            $createWeb.Url = $webUrl
            $createWeb.UseSamePermissionsAsParentSite = $true
            $createWeb.Title = $title
            $createWeb.WebTemplate = "$($template)#$($configuration)"
            $web = $parentWeb.Webs.Add($createWeb)
            $context.Load($web)
            Write-Host "Creating new web at $webUrl" -f Green -NoNewline
            $context.ExecuteQuery()
            Write-Host "...Created" -f Green
        }
    }
    Install-Lists $context $web $location
    Install-ListItems $context $web $location
}
Function Install-Masterpages($context, $web, $location)
{
	$path = [System.IO.Path]::Combine($location, "Masterpages")
	
    if((Test-Path $path) -eq $false)
    {
      write-host "No Master pages to deploy"
      return
    }    

	$list = $context.Site.GetCatalog([Microsoft.SharePoint.Client.ListTemplateType]::MasterPageCatalog)
	$context.Load($list)
    $context.Load($list.RootFolder)

	$context.ExecuteQuery();
	$publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
	
	$files = Get-ChildItem $path -Recurse | Where-Object { $_.PSIsContainer -eq $false}
	foreach($file in $files)
	{
		Write-Host "Uploading Master page $($file.Name)"
		$folderUrl = ($file.FullName -ireplace [regex]::escape($path),"") -ireplace "\\","/" -ireplace ("/" + [regex]::escape($file.Name)), ""
        $fileUrl = ($folderUrl + "/" + $file.Name)
        if([string]::IsNullOrEmpty($folderUrl) -eq $false)
        {
		    Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl | out-null
        }
        $fullFileUrl = Combine-Url $list.RootFolder.ServerRelativeUrl $fileUrl
        UploadFile $context $web $file $fullFileUrl $null $versioningEnabled $publishingEnabled | out-null
        PublishFile $context $web $fullFileUrl $versioningEnabled $publishingEnabled
	}
}
function ApplyMetadataToListItem([Microsoft.SharePoint.Client.ClientContext]$context,[Microsoft.SharePoint.Client.List]$list, [Microsoft.SharePoint.Client.ListItem] $listItem, $properties)
{
    foreach($fieldValue in $properties)
        {
            $name = $fieldValue.name
            if($name -eq "Attachments" )
            {
                continue
            }
            [Array]$field = $list.Fields | Where-Object { $_.InternalName -eq $name }
            if($field.Length -eq 0)
            {
                Write-Warning "Field does not exist for : $($list.Name) - $name"
                continue
            }
            if($field.IsReadOnly) {
                Write-Warning "Skipping readonly field"
                continue
            }
            $value = $fieldValue.'#cdata-section' | ConvertFrom-Json
            if([string]::IsNullOrEmpty($value))
            {
                continue
            }
            switch ($field[0].TypeAsString)
            {
                "File" {}
                "DateTime" {
                    $value = Get-Date ($value.value) #.ToString()
                    $listItem[$name] = $value
                }
                "URL" {
                    $urlValue = New-Object Microsoft.SharePoint.Client.FieldUrlValue 
                    $urlValue.Url = $value.Url
                    $urlValue.Description = $value.Description
                    $listItem[$name] =  [Microsoft.SharePoint.Client.FieldUrlValue]$urlValue
                }
                "Lookup" { 
                    $lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                    $lookupValue = $value.Split(":")[0]
                    $listItem[$name] = $lookupValue
                }
                default {
                    $listItem[$name] = $value
                }
            }
            
        }
}
Function Install-ListItems($context, $web, $location)
{
    #Dev Notes: If the destination environment contains required fields like Title, and we aren't exporting that field value
    #then this will error.  Attempted to find required fields and set a default value, but for some reason
    #the fields from the list did not show as required even though it did in SP.

  $path = [System.IO.Path]::Combine($location, "Lists")
  $xmlFilePath = [System.IO.Path]::Combine($path, "ListItems.xml")
  if((Test-Path $path) -eq $false -or (Test-Path $xmlFilePath) -eq $false)
  {
    Write-Host "No list items to import"
    return
  }

  [xml]$manifest = Get-Content $xmlFilePath

  foreach($listDef in $manifest.xml.Lists.List)
  {
    $listName = $listDef.name 
    $totalItems = $listDef.ListItem.Length
    Write-Host "Importing $totalItems list items to $listName"
    $targetWeb = $web
    if(Is-True $listDef.rootWeb)
    {
        $targetWeb = $rootWeb
    }

    $list = $targetWeb.Lists.GetByTitle($listName)
    #get the first item.  If there is data in the list we won't import.
    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = "<View><Query></Query><RowLimit>1</RowLimit></View>"
    $listItems = $list.GetItems($query)
    $fields = $list.Fields
    $context.Load($list)    
    $context.Load($list.RootFolder)    
    $context.Load($listItems)
    $context.Load($fields)
    $context.ExecuteQuery()
    $isLibrary = $list.BaseType -eq "1"
    $publishingEnabled = $list.EnableModeration
	$versioningEnabled = $list.EnableVersioning
    $filePath = [System.IO.Path]::Combine($path, "Files\$listName") 
    if($listItems.Count -ge 1)
    {
        Write-Warning "The list $listName already has content, items will not be imported"
        continue
    }
    foreach($listItemDef in $listDef.ListItem)
    {
        if($isLibrary)
        {
            $fileRef = ($listItemDef.FieldValue | Where-Object { $_.name -eq "FileLeafRef" }).'#cdata-section'
			$fileRef = $fileRef.Replace("`"","") #CData sections contain literal quotes if its a string value.
            $filename = [IO.Path]::Combine($filePath, $fileRef)
            $file = Get-ChildItem $filename 
            if($file -eq $null -or $file.Exists -eq $false)
            {
                throw "Could not locate file for library at $filename"
            }
            
            $fullFileUrl = Combine-Url $list.RootFolder.ServerRelativeUrl $fileRef
            if($versioningEnabled -and (FileExists $context $web $fullFileUrl))
            {
                CheckOutFile $context $web $fullFileUrl 
            }
            [Microsoft.SharePoint.Client.FileCreationInformation]$fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
            $fileInfo.Content = [IO.File]::ReadAllBytes($filename)
            $fileInfo.Overwrite = $true
            $fileInfo.Url = $fullFileUrl
            $uploadedFile = $list.RootFolder.Files.Add($fileInfo)
            $listItem = $uploadedFile.ListItemAllFields

            ApplyMetadataToListItem $context $list $listItem $listItemDef.FieldValue

            $listItem.Update()
            $context.ExecuteQuery()
            PublishFile $context $web $fullFileUrl $versioningEnabled $publishingEnabled
        }
        else {
            $itemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $listItem = $list.AddItem($itemInfo)
            ApplyMetadataToListItem $context $list $listItem $listItemDef.FieldValue
            $listItem.Update()
            $context.ExecuteQuery()
        }
    }
    $Context.ExecuteQuery()
  }
}
Export-ModuleMember -Function 'Install-*'
Export-ModuleMember -Function 'Activate-*'
