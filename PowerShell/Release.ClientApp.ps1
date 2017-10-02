<#
.SYNOPSIS
	Deploys a module Client Application to SharePoint (2013/16/Online) using the Client Side Object Model.

.DESCRIPTION
	A Client Application in this context is a defined package of Scripts, Images, Content Types, Lists, List Items, Documents, Web Parts, Display Templates, Layouts, Master Pages.
	These are modular components that have a manifest.xml file that describes the functionality and dependencies.  
	This script pushes the package to {SiteCollectionUrl}/Style Library/Client Apps/[App Name] and will potentially deploy Content Types, Columns, 
	Lists or other items required by the Client App.
	This script will also load the Common and Release module from the Modules directory.

.PARAMETER siteUrl
	a URL to a SharePoint Site Collection or Web.

.PARAMETER spOnline
	When $true this parameter will utilise the Online credential system.  Set this to $fale for Active Directory authentication.

.PARAMETER username
	The login name to use.  Must have permissions to create and publish to the Style Library and Master Page libraries.

.PARAMETER password
	The password for the account.  If you do not provide this it will be requested.

.PARAMETER clientApps
	A full path to a Client App or an Array of paths. eg: "C:\Repos\sp-clientapps\Client Apps\Common","C:\Repos\sp-clientapps\Client Apps\Facebook"

.PARAMETER steps
	Use $null to deploy all steps, or specify a single or an array of steps to run.
	Availble steps are: "Web Parts","Style Library","Display Templates","Site Fields","Lists","List Items","Page Layouts", "Content Types"

.PARAMETER csomVersion
	This is either 15 or 16 depending on your version of SharePoint.  Use 15 for SP2013, 16 for SP2016 and SP Online.  
	This parameter determines which CSOM assemblies are loaded.  Note that once one version of the assemblies have been loaded
	into a PowerShell session, changing this parameter will have no effect.  If you need to utilise multiple CSOM versions you must close 
	and re-open your PowerShell session.

.EXAMPLE
	Deploy the Common and Facebook Client Apps to a Site Collection in SharePoint Online.
	Release.ClientApp.ps1 -siteUrl https://mysharepoint.sharepoint.com -username "myuser@mysharepoint.onmicrosoft.com" -clientApps "C:\Repos\sp-clientapps\Client Apps\Common","C:\Repos\sp-clientapps\Client Apps\Facebook"

.EXAMPLE 

	Deploy the Style Library assets for the Facebook Client App to a Site Collection in SharePoint Online.
	Release.ClientApp.ps1 -siteUrl https://mysharepoint.sharepoint.com -username "myuser@mysharepoint.onmicrosoft.com" -steps "Style Library" -clientApps "C:\Repos\sp-clientapps\Client Apps\Facebook"

.NOTES
	These client apps require the presence of the Style Library which is provisioned in a Publishing Site.  
	You could create this library manually for non-publishing sites but this scenario has not been tested.
	The Client App "Common" is a special Client App that should be deployed first.  Other Client Apps will utilise libraries or other assets from here.
	The CSOM versions can be downloaded from:
	15 -  https://www.nuget.org/packages/Microsoft.SharePoint2013.CSOM/
	16 - https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/
    The assemblies must be extracted and placed in the CSOM/15 or CSOM/16 folder.
#>
param(
    $siteUrl = $($(read-host "Site Collection Url")),
    $spOnline = $($true),
    $username = $($(Read-Host "Username")),
    $password = $($null),
    $clientApps = $($(Read-Host "Full Path (or Array of paths) to Client App(s)")), 
    $steps = $($null), #"Web Parts","Style Library","Display Templates","Site Fields","Lists","List Items","Page Layouts", "Content Types"
    $csomVersion = "16" #15 = SP2013, 16=SP2016/Online
)
Add-Type -LiteralPath ($PSScriptRoot + "\CSOM\$csomVersion\Microsoft.SharePoint.Client.dll") -PassThru | out-null
Import-Module "$PSScriptRoot\Modules\Common.Module.psm1" -Force -DisableNameChecking 
Import-Module "$PSScriptRoot\Modules\Release.Module.psm1" -Force -DisableNameChecking 

function Deploy([ScriptBlock] $scriptBlock, $name) {
    try {
        if ($steps -ne $null -and $steps.Contains($name) -eq $false) {
            Write-Host "Skipping step '$name' as its not provided in the steps parameter"
            return
        }
        write-host "Deploying $name..."
        & $scriptBlock
    }
    catch {
        write-error ("Error building $name " + $_ + " : " + $_.ScriptStackTrace)
    }
}
try {
    $context = Login-Web $siteUrl $username $password -spOnline:$spOnline
    if($context -eq $null)
    {
        return
    }
    $web = $context.Web
    $rootWeb = $context.Site.RootWeb
    $context.Load($rootWeb)
    $context.ExecuteQuery()
    foreach ($clientApp in $clientApps) {
        if ($clientApp -ne $null) {
            if ($clientApp.EndsWith(".xml")) {
                $xmlFilePath = $clientApp;
            }
            else {
                $xmlFilePath = [System.IO.Path]::Combine($clientApp, "manifest.xml")
            }
            if ((Test-Path $xmlFilePath) -eq $null) {
                throw "No manifest exists for client app at $clientApp"
            }
            [xml]$manifest = get-content $xmlFilePath
            $location = ([System.IO.Path]::GetDirectoryName($xmlFilePath))
            
            if($steps -eq $null) {
                Install-ClientApp $context $web $rootWeb $location $manifest #This command runs all steps
            }
            else {
                #note: This process should match Install-ClientApp in Release.Module.psm1
                write-host "Deploying Client App - $($manifest.xml.ClientApp.Name)"
                Deploy { Install-WebParts $context $rootWeb $location } "Web Parts"
                Deploy { Install-StyleLibrary $context $rootWeb $location } "Style Library"
                Deploy { Install-SiteFields $context $rootWeb $location } "Site Fields"
                Deploy { Install-ContentTypes $context $rootWeb $location } "Content Types"
                Deploy { Install-Lists $context $web $location } "Lists"
                Deploy { Install-ListItems $context $web $location } "List Items"
                Deploy { Install-DisplayTemplates $context $rootWeb $location } "Display Templates"
            }
        }
    }
}
finally {
    if ($context -ne $null) {
        $context.Dispose()
        $context = $null
    }
}
"Script Complete"
