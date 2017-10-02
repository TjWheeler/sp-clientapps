# sp-clientapps
Client Applications for SharePoint On-Prem and Online

## About Client Apps
A Client Application is a defined package that may contain any of the following: Scripts, Images, Content Types, Lists, List Items, Documents, Web Parts, Display Templates, Layouts, Master Pages.
The deployment of the client applications is automated through PowerShell. 
Client Applications are deployed to {SiteCollectionUrl}/Style Library/Client Apps/[App Name].

## Conents
* PowerShell - This folder contains the release scripts to deploy the Client Apps.
* Client Apps/Facebook - A configurable JavaScript based Facebook web part that 


## The Facebook Web Part

### Configuration Requirements

In order to configure the Facebook Client App Web Part, you must:
* Sign Up as a Facebook App Developer (https://developers.facebook.com/)
* Register a new App 
* Get your Access Token from https://developers.facebook.com/tools/access_token/.  This is used in the Web Part configuration.

For further reading on the Facebook app process visit https://developers.facebook.com/docs/apps/register

### Example: Deploying the Facebook Web Part 

Deploy the Common and Facebook Client Apps to a Site Collection in SharePoint Online.
```
	.\PowerShell\Release.ClientApp.ps1 -siteUrl https://mysharepoint.sharepoint.com -username "myuser@mysharepoint.onmicrosoft.com" -clientApps "C:\Repos\sp-clientapps\Client Apps\Common","C:\Repos\sp-clientapps\Client Apps\Facebook"
```
This example assumes the repository is located at C:\Repos\sp-clientapps.  If it is in a different location you will need to update the path.
