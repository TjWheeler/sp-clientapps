# sp-clientapps
Client Applications for SharePoint On-Prem and Online

## About Client Apps
A Client Application is a defined package that may contain any of the following: Scripts, Images, Content Types, Lists, List Items, Documents, Web Parts, Display Templates, Layouts, Master Pages.
The deployment of the client applications is automated through PowerShell. 
Client Applications are deployed to {SiteCollectionUrl}/Style Library/Client Apps/[App Name].
The Client Apps are designed to be deployed to the Style Library which is available in SharePoint Publishing Sites.  

## Conents
* PowerShell - This folder contains the release scripts to deploy the Client Apps.
* Client Apps/Facebook - A configurable JavaScript based Facebook web part that 


## The Facebook Web Part

See the [Facebook Blog Post](http://blog.timwheeler.io/facebook-webpart/) for more detailed instructions.

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

### Adding the Web Part to your Page

* Edit or Create a Publishing Page
* Click Add Web Part on desired Web Part zone
* From the Web Part Categories, find and click on 'Client Apps'
* Select the Client App and click Add
* While still in Edit Mode, the app will display a Gears Icon.  Click this to show the Client App Configuration.
* Make sure you click the Update button after entering your configuration values.  Save the page to show the web part.

### Troubleshooting

If nothing displays, check the console for errors by pressing the F12 key in IE/Edge/Chrome.
Ensure you are deploying to a publishing site, or create a Style Library manually.
Make sure you deploy the 'Common' client app before installing anything else.
