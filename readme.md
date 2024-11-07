Inspiration:

https://javascript.plainenglish.io/onedrive-integration-with-react-step-by-step-guide-c068bb8e3fb8

App Only Authentication: (works but not for OneDrive)
https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http


User auth access only: (Nov 2023)
https://github.com/abraunegg/onedrive/discussions/2548

Hack appraoch using user auth:
https://stackoverflow.com/questions/54643731/microsoft-onedrive-create-folder-using-api-key-without-login


some other links:

https://techcommunity.microsoft.com/t5/microsoft-sharepoint-blog/develop-applications-that-use-sites-selected-permissions-for-spo/ba-p/3790476





1 -Install Powershell:
brew install powershell/tap/powershell



2- Install SharePoint Online Management Shell:

Install-Module -Name Microsoft.Online.SharePoint.PowerShell  -Repository PSGallery -Force
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
Connect-AzureAD

3- Connect to SharePoint:

Connect-SPOService -Credential $creds -Url https://simnovaoffice-admin.sharepoint.com/ -ModernAuth $true -AuthenticationUrl https://login.microsoftonline.com/organizations

4- List Groups

Get-SPOSiteGroup -Site https://simnovaoffice.sharepoint.com/sites/onedrivepoc

Install-Module -Name Az -Repository PSGallery -Force