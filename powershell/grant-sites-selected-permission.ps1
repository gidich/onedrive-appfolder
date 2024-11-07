# PowerShell Script to grant an AAD App access to the "General" folder of a Microsoft Teams channel
param (
    [Parameter(Mandatory=$true)]
    [string]$TeamsTeamUrl,

    [Parameter(Mandatory=$true)]
    [string]$ChannelName,

    [Parameter(Mandatory=$true)]
    [string]$AADApplicationId
)

function Write-ErrorMessage {
    param (
        [string]$ErrorMessage
    )
    Write-Host "ERROR: $ErrorMessage" -ForegroundColor Red
    exit 1
}

try {
    
    # Extracting the tenant from the Teams URL
    if ($TeamsTeamUrl -notmatch "https://(.+?).sharepoint.com/sites/(.+)") {
        Write-ErrorMessage -ErrorMessage "Invalid Teams team URL. Ensure the URL is in the correct format."
    }
    
    $tenantName = $matches[1]
    $siteName = $matches[2]
    $AdminUrl = "https://$tenantName-admin.sharepoint.com"
    
    # Authenticating to SharePoint Online using modern authentication
    try {
      Connect-SPOService -Url $AdminUrl -ModernAuth $true -AuthenticationUrl https://login.microsoftonline.com/organizations
    } catch {
      Write-ErrorMessage -ErrorMessage "Failed to authenticate. Please verify your credentials."
    }
    
    # Retrieve SharePoint site information
    $site = Get-SPOSite -Identity "https://$tenantName.sharepoint.com/sites/$siteName"
    if (-not $site) {
      Write-ErrorMessage -ErrorMessage "Failed to retrieve the site. Check that the provided URL is correct."
    }

    # getting the OneDrive Tenant ID
    # Construct the OpenID configuration URL
    $oidcConfigUrl = "https://login.microsoftonline.com/$tenantName.onmicrosoft.com/.well-known/openid-configuration"

    # Get the OpenID configuration
    $oidcConfig = Invoke-RestMethod -Method Get -Uri $oidcConfigUrl

    # Extract the tenant ID from the issuer URL
    $tenantId = $oidcConfig.issuer.Split('/')[3].TrimEnd('/')

    # Output the tenant ID
    Write-Host "ONEDRIVE_TENANT_ID: $tenantId"

    # Connecting to Microsoft Graph with interactive authentication to ensure consent is granted
    try {
      Connect-MgGraph -Scopes "Sites.FullControl.All" -TenantId "$tenantId" -Environment "Global" -NoWelcome
    } catch {
      Write-ErrorMessage -ErrorMessage "Failed to authenticate to Microsoft Graph. Please verify your credentials."
    }

    try {
      # Getting the Drive ID of the team's document library
      $drive = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/sites/${tenantName}.sharepoint.com:/sites/${siteName}:/drive" 
      $teamDriveId = $drive.id
      Write-Host "ONEDRIVE_DRIVE_ID: $teamDriveId"

      # Getting the Folder ID for the ChannelName
      $rootItems =  Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/drives/${teamDriveId}/root/children"
      #for each item in $rootitems.value hashtable, inspect the rootItem.Value hastable and find the one with Name matching $ChannelName and return the ID of this one item
      $channelDirectory = $rootItems.value | Where-Object { $_.name -eq $ChannelName } | Select-Object -First 1
      $channelDirectoryId = $channelDirectory.id
      Write-Host "ONEDRIVE_ROOT_FOLDER_ID: $channelDirectoryId"
     
      # Get The SiteID
      $foundSite = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/sites/${tenantName}.sharepoint.com:/sites/${siteName}:"
      $siteId = $foundSite.id
      Write-Host "SITE ID: $siteId"
     
      # Granting AAD Application access to the specified Team's Site
      $permissionBodySite = @{ roles = @("write"); grantedToIdentities = @( @{ application = @{ id = $AADApplicationId;  displayName = "aadAppOnly" } } ) } | ConvertTo-Json -Depth 3
      
      $permissionResult = Invoke-MgGraphRequest -Method POST "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -Body $permissionBodySite -OutputType Json -SkipHttpErrorCheck
      Write-Host "Permssion Result: $permissionResult"
    } catch {
        Write-ErrorMessage -ErrorMessage $_.Exception.Message
    }

} catch {
  Write-ErrorMessage -ErrorMessage $_.Exception.Message
}

# Sample usage:
# .\grant-sites-selected-permission.ps1 -TeamsTeamUrl "https://tenantname.sharepoint.com/sites/TeamName" -ChannelName "General" -AADApplicationId "YOUR AAD APPLICATION ID" 