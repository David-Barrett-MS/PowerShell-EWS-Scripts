 
# Test-ManagementActivityAPI.ps1
#
# By David Barrett, Microsoft Ltd. 2018-2021. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

<#
.SYNOPSIS
Interact with the Office 365 Management API.

.DESCRIPTION
This script allows you to interact with the Office 365 Management API.  You must register your application in Azure (you can use -RegisterAzureApplication and this script to do so) AND grant admin consent prior to use.

.EXAMPLE
.\Test-ManagementActivityAPI.ps1 -Start -AppId "afc98c8f-1faf-4d38-b0d5-e421070134c7" -TenantId "fc69f6a8-90cd-4047-977d-0c768925b8ec" -AppSecretKey "xx"

.EXAMPLE
.\Test-ManagementActivityAPI.ps1 -ListContent -ListContentDate $([DateTime]::Now.AddDays(-1)) -AppId "afc98c8f-1faf-4d38-b0d5-e421070134c7" -TenantId "fc69f6a8-90cd-4047-977d-0c768925b8ec" -AppSecretKey "xx"

This will list all content that was made available to the Office 365 Management API yesterday (24 hour period assuming UTC).  The AppId, TenantId and AppSecretKey are obtained by registering an application in Azure.

.EXAMPLE
.\Test-ManagementActivityAPI.ps1 -RetrieveContent -SaveContentPath "c:\Temp\API Data" -ListContentDate $([DateTime]::Now.AddDays(-1)) -AppId "afc98c8f-1faf-4d38-b0d5-e421070134c7" -TenantId "fc69f6a8-90cd-4047-977d-0c768925b8ec" -AppSecretKey "xx"

This will retrieve all content that was made available to the Office 365 Management API yesterday (24 hour period assuming UTC).  The AppId, TenantId and AppSecretKey are obtained by registering an application in Azure.

.EXAMPLE
.\Test-ManagementActivityAPI.ps1 -RegisterAzureApplication -AzureApplicationName "ManagementAPIData" -Verbose

The above will create an application called ManagementAPIData in the specified tenant, and configure permissions needed to be able to read all data.  It will display the secret key and the application Id that you'll need to take a note of.
This uses the AzureAD module, which will be automatically installed if necessary (and you have permissions to install).  You'll be prompted to log in when you run the script (tenant information is automatically retrieved).
A tenant administrator will need to grant the permissions to the application once it has been created (this is done via the Azure Portal).
If you know your tenant Id, or want to specify the tenant (in case the Azure registered application is multi-tenant), then that can be specified and the script will attempt to connect directly to that (prompting for credentials).

#>

param (
    [Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId = "",

    [Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD)")]
    [ValidateNotNullOrEmpty()]
    [string]$AppSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="Authentication certificate (certificate must include the private key as this is used to identify the application as registered in Azure)")]
    [ValidateNotNullOrEmpty()]
    [string]$AppAuthCertificate = "",

    [Parameter(Mandatory=$False,HelpMessage="Redirect URI for the application")]
    [ValidateNotNullOrEmpty()]
    [string]$AppRedirectURI = "http://localhost/TestManagementActivityAPI",

    [Parameter(Mandatory=$False,HelpMessage="Tenant Id")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="Publisher Id (this is the tenant Id of the publisher - if specified, the publisher's quota will be used)")]
    [ValidateNotNullOrEmpty()]
    [string]$PublisherId = "",

    [Parameter(Mandatory=$False,HelpMessage="Start subscription.  If ContentType not specified, will attempt to enable all.")]
    [switch]$Start,

    [Parameter(Mandatory=$False,HelpMessage="Webhook address (URL to which audit logs will be sent).  Note that webhooks are no longer recommended.")]
    [string]$WebhookAddress = "",

    [Parameter(Mandatory=$False,HelpMessage="Which audit logs do we want to retrieve?  Default is general audit logs.  Can be left blank when starting subscriptions to enable collection of all types.")]
    [ValidateNotNullOrEmpty()]
    [string]$ContentType = "",

    [Parameter(Mandatory=$False,HelpMessage="Stop subscription")]
    [switch]$Stop,

    [Parameter(Mandatory=$False,HelpMessage="List current subscriptions")]
    [switch]$List,

    [Parameter(Mandatory=$False,HelpMessage="List available content")]
    [switch]$ListContent,

    [Parameter(Mandatory=$False,HelpMessage="Retrieve available content (implies -ListContent, but retrieves the content as well as the location of the content)")]
    [switch]$RetrieveContent,

    [Parameter(Mandatory=$False,HelpMessage="If this is specified, content will be saved to this path (each content blob will be a separate text file)")]
    [string]$SaveContentPath,

    [Parameter(Mandatory=$False,HelpMessage="Maximum number of jobs that can run at one time (in parallel) to download content.")]
    [int]$MaxRetrieveContentJobs = 32,

    [Parameter(Mandatory=$False,HelpMessage="Date for which to retrieve content (if specified, will retrieve the 24 hours of data for this date).  Do not specify start and end time.")]
    $ListContentDate,

    [Parameter(Mandatory=$False,HelpMessage="If specified, this is the start of the time range that will be requested.  Do not use with -ListContentDate.")]
    $ListContentStartTime,

    [Parameter(Mandatory=$False,HelpMessage="If specified, this is the start of the time range that will be requested.  Do not use with -ListContentDate.")]
    $ListContentEndTime,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the script attempts to register an application in Azure using the given parameters (and with permission to access Management API logs)")]
    [switch]$RegisterAzureApplication,

    [Parameter(Mandatory=$False,HelpMessage="Name of the application to register in Azure (required when -RegisterAzureApplication specified)")]
    [ValidateNotNullOrEmpty()]
    [string]$AzureApplicationName = "",

    [Parameter(Mandatory=$False,HelpMessage="Permissions that the application will require (these are all application permissions as this script authenticates as application)")]
    [ValidateNotNullOrEmpty()]
    $AzureApplicationRequiredPermissions = @("ActivityFeed.Read", "ActivityFeed.ReadDlp", "ServiceHealth.Read"),

    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file")]	
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="HTTP trace file - all HTTP request and responses will be logged to this file")]	
    [string]$DebugPath = ""
)
$script:ScriptVersion = "1.1.5"

# We work out the root Uri for our requests based on the tenant Id
$rootUri = "https://manage.office.com/api/v1.0/$tenantId/activity/feed"

$AvailableContentTypes = @("Audit.AzureActiveDirectory", "Audit.Exchange", "Audit.SharePoint", "Audit.General", "DLP.All")


########################################################
#
# Function definitions
#
########################################################

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
    LogToFile $Details
}

$script:LastError = $Error[0]
Function ErrorReported($Context)
{
    # Check for any error, and return the result ($true means a new error has been detected)

    # We check for errors using $Error variable, as try...catch isn't reliable when remoting
    if ([String]::IsNullOrEmpty($Error[0])) { return $false }

    # We have an error, have we already reported it?
    if ($Error[0] -eq $script:LastError) { return $false }

    # New error, so log it and return $true
    $script:LastError = $Error[0]
    if ($Context)
    {
        Log "Error ($Context): $($Error[0])" Red
    }
    else
    {
        Log "Error: $($Error[0])" Red
    }
    return $true
}

Function ReportError($Context)
{
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}

function LoadLibraries
{
    param (
        [parameter(Position=0,Mandatory=$true)][bool]$searchProgramFiles,
        [parameter(Position=1,Mandatory=$true)][array]$dllNames
    )
    # Attempt to find and load the specified libraries

    foreach ($dllName in $dllNames)
    {
        # First check if the dll is in current directory
        $dll = $null
        try
        {
            $dll = Get-ChildItem $dllName
        }
        catch {}

        if ($searchProgramFiles)
        {
            if ($dll -eq $null)
            {
	            $dll = Get-ChildItem -Recurse "C:\Program Files (x86)" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            if (!$dll)
	            {
		            $dll = Get-ChildItem -Recurse "C:\Program Files" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            }
            }
        }

        if ($dll -eq $null)
        {
            Log "Unable to load locate $dll" Red
            return $false
        }
        else
        {
            try
            {
		        LogVerbose ([string]::Format("Loading {2} v{0} found at: {1}", $dll.VersionInfo.FileVersion, $dll.VersionInfo.FileName, $dllName))
		        Add-Type -Path $dll.VersionInfo.FileName
            }
            catch
            {
                return $false
            }
        }
    }
    return $true
}

function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!script:msalApiLoaded)
    {
        $msalLocation = @()
        $script:msalApiLoaded = $(LoadLibraries -searchProgramFiles $false -dllNames @("Microsoft.Identity.Client.dll") -dllLocations ([ref]$msalLocation))
        if (!$script:msalApiLoaded)
        {
            Log "Failed to load MSAL.  Cannot continue with certificate authentication." Red
            exit
        }
    }   

    $cca1 = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca2 = $cca1.WithCertificate($OAuthCertificate)
    $cca3 = $cca2.WithTenantId($OAuthTenantId)
    $cca = $cca3.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://manage.office.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
}

function GetTokenWithKey
{
    # We have secret key, so use that to request access token

    $Body = @{
      "grant_type"    = "client_credentials";
      "client_id"     = "$AppId";
      "client_secret" = "$AppSecretKey";
      "scope"         = "https://manage.office.com/.default"
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token: $Error" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}

# Get our OAuth token
function GetAccessToken
{
    # Obtain OAuth token for accessing mailbox

    $script:LastError = $Error[0]
    if (![String]::IsNullOrEmpty($AppSecretKey))
    {
        # We are authenticating using secret key
        GetTokenWithKey
        LogVerbose "OAuth call (secret key) complete"
        return $script:oAuthAccessToken
    }

    if ([String]::IsNullOrEmpty($AppAuthCertificate))
    {
        Log "Either certificate or secret key must be supplied" Red
        Exit
    }

    # We are using certificate authentication - we currently need ADAL for this
    GetTokenWithCertificate

    # Check we've got an access token
    LogVerbose "OAuth call complete"
    return $script:oAuthAccessToken
}

function GetValidAccessToken
{
    # Check if access token needs renewing, and if so renew it.  We renew only if token has expired (before this, ADAL will simply return the same one)
    if ($script:oauthToken -ne $null)
    {
        # We already have a token
        if ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1))
        {
            # Token still valid, so return that
            return $script:oAuthAccessToken
        }

        # Token needs renewing
    }


    # We either don't have an access token, or it has expired
    Log("Obtaining OAuth access token")
    $script:accessToken = GetAccessToken
    if ( $script:oauthTokenAcquireTime -and $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -le [DateTime]::UtcNow )
    {
        Log "Access token is expired" Red
        exit
    }
    Log "OAuth token renewed, will expire at $($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in))" Green
    return $script:authenticationResult.Result.AccessToken
}

function RetriableInvoke-WebRequest
{
    param (
        [parameter(Mandatory=$true)][string]$Uri,
        [parameter(Mandatory=$false)]$Headers = $null,
        [parameter(Mandatory=$false)][string]$ContentType,
        [parameter(Mandatory=$false)][string]$Body,
        [parameter(Mandatory=$false)][string]$Method = "Get"
    )

    # Error trapped Invoke-WebRequest

    $retries = 0
    $result = $null

    # Ensure we have a valid access token
    if ($Headers -ne $null)
    {
        if ($Headers.ContainsKey("Authorization"))
        {
            $Headers.Remove("Authorization")
        }
        $Headers.Add("Authorization", "Bearer $(GetValidAccessToken)")
    }
    else
    {
        $Headers = @{"Authorization" = "Bearer $(GetValidAccessToken)"}
    }

    do
    {
        if ($retries -gt 0)
        {
            LogVerbose "Retry attempt $retries"
            if ($retries -gt 1)
            {
                # For retries after the second, we add a delay before retrying.
                Log "Waiting $(($retries-1) * 30) seconds before retrying" Yellow
                Start-Sleep -Seconds (($retries-1) * 30)
            }
        }
        try
        {
            if ( $Method -eq "Post" )
            {
                $result = Invoke-WebRequest -Uri $Uri -Headers $Headers -Method $Method -ContentType $ContentType -Body $Body
            }
            else
            {
                $result = Invoke-WebRequest -Uri $Uri -Headers $Headers -Method $Method
            }
        }
        catch [System.Net.WebException] {
            ReportError "Invoke-WebRequest"
            if ($error[0].Exception.ToString().Contains("(400) Bad Request"))
            {
                $result = $error[0].ErrorDetails.Message
            }
        }
        catch {
            Write-Host "Failed" Red
            exit
        }
        $retries++
    } until ( ($result -ne $null) -or ($retries -gt 3) )

    return $result

}

$script:requestIndex = 1 # We trace by dumping each request and response to a new file
Function GetWithTrace()
{
    param (
        [parameter(Position=0,Mandatory=$true)][string]$requestUrl,
        [parameter(Position=1,Mandatory=$false)]$headers = $null
    )

    if ( [String]::IsNullOrEmpty($DebugPath) )
    {
        return $(RetriableInvoke-WebRequest -Uri $requestUrl -Headers $headers -Method Get)
    }

    $traceFilename = $DebugPath
    if (!$traceFilename.EndsWith("\")) { $traceFilename = "$traceFilename\" }
    $traceFilename = "$traceFilename$($script:requestIndex)"
    LogVerbose "Tracing request $($script:requestIndex) to: $traceFileName"
    $script:requestIndex++

    "GET $requestUrl" |  Out-File "$traceFilename.request"
    $headers | Format-Table -HideTableHeaders -Wrap | Out-File "$traceFilename.request" -Append
    $data = RetriableInvoke-WebRequest -Uri $requestUrl -Headers $headers -Method Get
    if ($data.RawContent)
    {
        $data.RawContent | Out-File "$traceFilename.response"
    }
    else
    {
        $data | Out-File "$traceFilename.response"
    }
    return $data
}

# POST a REST request and receive response
Function PostRest
{
    param (
        [parameter(Position=0,Mandatory=$true)][string]$requestUrl,
        [parameter(Position=1,Mandatory=$true)]$request
    )

    return RetriableInvoke-WebRequest -Uri $requestUrl -ContentType "application/json" -Method Post -Body $request
}

# GET a REST response
Function GetRest
{
    param (
        [parameter(Position=0,Mandatory=$true)][string]$requestUrl
    )

    $script:nextPageUri = ""
    LogVerbose "REST query: $requestUrl"
    $thisPage = GetWithTrace -requestUrl $requestUrl
    if ($thisPage.Headers.NextPageUri -ne $null)
    {
        $script:nextPageUri = $thisPage.Headers.NextPageUri
        LogVerbose "NextPageUri: $($script:nextPageUri)"
    }
        
    return $thisPage.Content
}

# Create a secret key that can be used for an Azure application
Function CreateSecretKey
{
    $aes = New-Object System.Security.Cryptography.AesManaged
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::Zeros
    $aes.BlockSize = 128
    $aes.KeySize = 256
    $aes.GenerateKey()
    return [System.Convert]::ToBase64String($aes.Key)
}

# Create a certificate used to authenticate with Azure.  
Function CreateAuthCertificate
{
    # Not implemented
    makecert -r -pe -n "CN=MyCompanyName MyAppName Cert" -b 03/15/2015 -e 03/15/2017 -ss my -len 2048
}

# Create application key that can be used for an Azure application
Function CreateAppKey([DateTime] $ValidFromDate, [double] $DurationInYears, [string]$SecretKey)
{
    $key = New-Object Microsoft.Open.AzureAD.Model.PasswordCredential
    $key.StartDate = $ValidFromDate
    $key.EndDate = $ValidFromDate.AddYears($DurationInYears) 
    $key.Value = $SecretKey
    $key.KeyId = (New-Guid).ToString()
    return $key
}

Function CreatePermissionSet([string] $ResourceDisplayName, [string[]]$Permissions, [string[]]$PermissionTypes)
{
    # Get information about the resource (which will include available permissions)
    $resourceSP = $null
    $resourceSP = Get-AzureADServicePrincipal -Filter "DisplayName eq '$resourceDisplayName'"
    if (!$resourceSP)
    {
        Log "Failed to locate resource API: $resourceDisplayName" Red
        return $false
    }

    # Create a RequiredResourceAccess object (this will be used to specify the permissions that our application needs)
    $requiredResourceAccess = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
    $requiredResourceAccess.ResourceAppId = $resourceSP.AppId
    $requiredResourceAccess.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]

    # $Permissions contains our list of permissions, while $PermissionTypes defines what type we need ("Scope" is for delegated permissions, while "Role" is for application permissions)
    $i = 0
    while ($i -lt $Permissions.Length)
    {
        # Find the matching permission from our resource (if we can't find it, it isn't a valid permission)
        $permissionFound = $false
        foreach ($resourcePermission in $resourceSP.OAuth2Permissions)
        {
            if ($resourcePermission.Value -eq $Permissions[$i])
            {
                # This is the permission that our application needs
                $permissionFound = $true
                $resourceAccess = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
                
                if ($PermissionTypes.Length -eq 1)
                {
                    $resourceAccess.Type = $PermissionTypes[0]
                    LogVerbose "Adding permission: $($resourcePermission.Value); Scope: $($PermissionTypes[0])"
                }
                else
                {
                    $resourceAccess.Type = $PermissionTypes[$i]
                    LogVerbose "Adding permission: $($resourcePermission.Value); Scope: $($PermissionTypes[$i])"
                }
                $resourceAccess.Id = $resourcePermission.Id # This is the Id of the permission we are requesting, as read from Azure
                $requiredResourceAccess.ResourceAccess.Add($resourceAccess)
                break
            }
        }
        if (!$permissionFound)
        {
            Log "$ResourceDisplayName does not expose the permission: $($Permissions[$i])" Yellow
        }
        $i++
    }

    return $requiredResourceAccess
}

# Register Azure application with correct permissions
Function RegisterAzureApplication
{
    # Check we have Azure AD module available
    $azureAD = Get-Module -Name "AzureAD"
    if (!$azureAD)
    {
        LogVerbose "AzureAD module not available, attempting to install"
        try
        {
            Install-Module "AzureAD"
        }
        catch
        {
            Log "Failed to install AzureAD module, which is required to register Azure applications" Red
            Exit
        }
    }

    LogVerbose "Attempting to register application in Azure"

    # Connect to Azure AD and obtain tenant information
    $tenant = $null
    if (![String]::IsNullOrEmpty($TenantId))
    {
        # If TenantId has been specified, we always connect to Azure AD (to ensure we have the right tenant)
        Connect-AzureAD -TenantId $TenantId | out-null
    }
    else
    {
        # When no tenant is specified, we only connect to Azure AD if we don't already have tenant information (i.e. are not logged on)
        try
        {
            $tenant = Get-AzureADTenantDetail
        } catch {}
        if ($tenant -eq $null)
        {
            Connect-AzureAD | out-null
        }
    }
    if ($tenant -eq $null)
    {
        try
        {
            $tenant = Get-AzureADTenantDetail
        } catch {}
        if ($tenant -eq $null)
        {
            Log "Failed to connect to Azure tenant" Red
            return
        }
    }
    
    # If we get here, then we have successfully logged onto a tenant
    $tenantName =  ($tenant.VerifiedDomains | Where { $_._Default -eq $True }).Name
    LogVerbose "Will register application in tenant: $tenantName"
    if ([String]::IsNullOrEmpty($TenantId))
    {
        $TenantId = $tenant.ObjectId
        Log "Tenant id: $TenantId"
    }

    # Create the Azure application
    LogVerbose "Creating the Azure application: $AzureApplicationName"

    if ([String]::IsNullOrEmpty($AppSecretKey))
    {
        # No secret key specified, so we create one
        $AppSecretKey = CreateSecretKey
        Log "Application secret key generated: $AppSecretKey`r`n"
    }
    
    $appRegKey = CreateAppKey -ValidFromDate $([DateTime]::Now) -DurationInYears 2 -SecretKey $AppSecretKey
    $azureApplication = New-AzureADApplication -DisplayName "$AzureApplicationName" -HomePage "https://localhost/$AzureApplicationName" -IdentifierUris "https://$tenantName/$AzureApplicationName" -PasswordCredentials $appRegKey -PublicClient $False -ReplyUrls @("$AppRedirectURI")
    if (!$azureApplication)
    {
        Log "Failed to create Azure application" Red
        return
    }
    New-AzureADServicePrincipal -AppId $($azureApplication.AppId) -Tags {WindowsAzureActiveDirectoryIntegratedApp} | out-null # This is required to make the application visible in the App Registrations (v1) blade in Azure AD
    Log "Azure application created; Id: $($azureApplication.AppId)"
    
    # Add Required Resources Access (from application to the Management API)
    LogVerbose "Getting access from '$AzureApplicationName' to 'Office 365 Management APIs'"
    $requiredPermissions = CreatePermissionSet -ResourceDisplayName "Office 365 Management APIs" -Permissions $AzureApplicationRequiredPermissions -PermissionTypes @("Role")
    if ($requiredPermissions -eq $false)
    {
        Log "Unable to build permission list for application.  Registration has been successful, but no permissions assigned." Red
        return
    }

    LogVerbose "Setting requested permissions on Azure application"
    $requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]
    $requiredResourcesAccess.Add($requiredPermissions)
    Set-AzureADApplication -ObjectId $azureApplication.ObjectId -RequiredResourceAccess $requiredResourcesAccess | out-null
    Log "Application created and permissions have been set - please grant permissions via the Azure Portal.  NOTE THE SECRET KEY, as it cannot be recovered if lost (though a new one can be created)."
}


########################################################
#
# Main script
#
########################################################

########################################################
#
# Register Azure Application
#
########################################################

if ($RegisterAzureApplication)
{
    RegisterAzureApplication
    Exit
}

########################################################
#
# Configure auth (obtain our token) and PublisherId
#
########################################################

$script:accessToken = GetAccessToken

if ([String]::IsNullOrEmpty($script:accessToken))
{
    Log "Failed to acquire valid access token" Red
    Exit
}

if ([String]::IsNullOrEmpty($PublisherId))
{
    # If we don't have a publisher Id, then we assume that the publisher is the tenant
    $PublisherId = $TenantId
}


########################################################
#
# Start subscription
#
########################################################

function SubscribeContentType([string]$SubscriptionContentType)
{
    $request = @{
        address = $WebhookAddress
        authId = "O365ActivityAPINotification"
        expiration = ""
    }
    [string]$json = ConvertTo-JSON -InputObject $request
    LogVerbose $("Webhook details: " + $json.Replace("`r`n", ""))
    PostRest "$rootUri/subscriptions/start?contentType=$SubscriptionContentType&PublisherIdentifier=$PublisherId" $json
}

if ($Start)
{
    if (![String]::IsNullOrEmpty($ContentType))
    {
        # We have a specified Content Type to subscribe to
        SubscribeContentType $ContentType
    }
    else
    {
        # No Content Type specified, so subscribe to all that we know about
        foreach ($ct in $AvailableContentTypes)
        {
            LogVerbose "Enabling Content-Type $ct"
            SubscribeContentType $ct
        }
    }
    Exit
}


########################################################
#
# Stop subscription
#
########################################################

if ($Stop)
{
    PostRest "$rootUri/subscriptions/stop?contentType=$ContentType&PublisherIdentifier=$PublisherId" ""
    Exit
}


########################################################
#
# List current subscriptions
#
########################################################

if ($List)
{
    GetRest "$rootUri/subscriptions/list?PublisherIdentifier=$PublisherId"
    Exit
}


########################################################
#
# List available content
#
########################################################


$script:contentUrls = @()

function ListContent($listContentType)
    {
    $script:nextPageUri = ""
    if ($ListContent -or $RetrieveContent)
    {
        if ($ListContentDate)
        {
            # We have a specified date to retrieve, so work out start and end date
            $startDate = $null
            if ($ListContentDate.GetType().Name.Equals("DateTime"))
            {
                # If date is supplied as DateTime, we just use the supplied value.  String conversion can cause DateFormat issues (e.g. UK->US date issues)
                $startDate = $ListContentDate
            }
            else
            {
                $startDate = [DateTime]::Parse($ListContentDate.ToString())
            }

            if ($startDate)
            {
                $startDate = [DateTime]::new($startDate.Year, $startDate.Month, $startDate.Day, 0, 0, 0)
                $endDate = $startDate.AddDays(1)
                $startDateStr = $startDate.ToString("yyyy-MM-ddTHH:mm:ss")
                $endDateStr = [String]::Format("{0:yyyy-MM-ddTHH:mm:ss}", $endDate)
                $script:nextPageUri = "$rootUri/subscriptions/content?contentType=$listContentType&PublisherIdentifier=$PublisherId&startTime=$startDateStr&endTime=$endDateStr"
            }
            else
            {
                Log "Failed to parse ListContentDate: $ListContentDate"
                Exit
            }
            Log "Listing content date range: Start = $startDate   End = $endDate" Green
        }
        else
        {
            if ($ListContentStartTime)
            {
                if (!$ListContentEndTime)
                {
                    Log "Must specify ListContentEndTime with ListContentStartTime" Red
                    exit
                }
                $startDate = [DateTime]::Parse($ListContentStartTime.ToString())
                $endDate = [DateTime]::Parse($ListContentEndTime.ToString())
                $startDateStr = $startDate.ToString("yyyy-MM-ddTHH:mm:ss")
                $endDateStr = [String]::Format("{0:yyyy-MM-ddTHH:mm:ss}", $endDate)
                $script:nextPageUri = "$rootUri/subscriptions/content?contentType=$listContentType&PublisherIdentifier=$PublisherId&startTime=$startDateStr&endTime=$endDateStr"
                Log "Listing content date range: Start = $startDate   End = $endDate" Green
            }
            else
            {
                # Retrieve the data for the current 24 hour period
                LogVerbose "Listing content from last 24 hours"
                $script:nextPageUri = "$rootUri/subscriptions/content?contentType=$listContentType&PublisherIdentifier=$PublisherId"
            }
        }

        # Retrieve the content links
        while ( ![String]::IsNullOrEmpty($script:nextPageUri) )
        {
            $sanityCheck = $script:nextPageUri
            $content = GetRest $script:nextPageUri
                
            if ($content -ne $null)
            {
                if (!$RetrieveContent)
                {
                    $content
                }
                else
                {
                    $jsonContent = ConvertFrom-JSON $content
                    Log "List content: $($jsonContent.Count) content blob(s) available for download" Green
                    if ($jsonContent.Count -gt 0)
                    {
                        foreach ($contentBlob in $jsonContent)
                        {
                            if ($contentBlob)
                            {
                                if ($contentBlob.ContentUri)
                                {
                                    $script:contentUrls += $contentBlob.ContentUri
                                    LogVerbose "Content Url added to retrieve list: $($contentBlob.ContentUri)"
                                }
                            }
                        }
                    }
                }
                $content = $null
            }
            else
            {
                # Failed to retrieve any content
                if ($sanityCheck -eq $script:nextPageUri)
                {
                    $script:nextPageUri = $null
                    Log "Unexpected failure in nextPageUri processing" Red
                }
            }
        }
    }
    LogVerbose "Finished ListContent for $listContentType"
}

if ([String]::IsNullOrEmpty($ContentType))
{
    # If we don't have a ContentType specified, we retrieve all of them
    foreach ($ContentType in $AvailableContentTypes)
    {
        ListContent $ContentType
    }
}
else
{
    ListContent $ContentType
}

########################################################
#
# Retrieve available content
#
########################################################

# Define the download function in a script block, as it is run as a job (which also means that it can't access any variables from the main session)

$downloadFunction = {
    function DownloadContentBlob([string]$contentUrl, [string]$auth, [string]$savePath)
    {
        # Download the specified content blob and save to file
        $auditData = ""
        $auditResponse = Invoke-WebRequest -Uri $contentUrl -Headers @{"Authorization" = $auth} -Method Get
        $auditData = $auditResponse.Content

        if ($auditData.Length -gt 0)
        {
            if (![String]::IsNullOrEmpty($savePath))
            {
                # Save this data

                # ContentUri will be of format https://manage.office.com/api/v1.0/fc69f6a8-90cd-4047-977d-0c768925b8ec/activity/feed/audit/20190205113446710142142$20190205134333251104202$audit_exchange$Audit_Exchange
                # We use the last part as the filename
                $outputFileName = $contentUrl.Substring($contentUrl.LastIndexOf("/")+1)
                [string]$outputFile = $savePath
                if (![String]::IsNullOrEmpty($outputFile) -and !$outputFile.EndsWith("\")) { $outputFile = "$outputFile\" }
                $outputFile = "$outputFile$outputFileName"
            
                if ($(Test-Path "$outputFile.txt"))
                {
                    # Output file already exists

                    # We perform a sanity check here to ensure that the blob we have already retrieved is the same data
                    $existingBlob = [IO.File]::ReadAllText("$outputFile.txt")
                    if (!$existingBlob.Equals($($auditData | out-string)))
                    {
                        # This data is different - which is unexpected, so we'll save it as an additional file
                        Write-Host "Content blob is different to the one already retrieved, but should be the same: $outputFile.txt" -ForegroundColor Red
                        $i = 1
                        while ($(Test-Path "$outputFile.$i.txt"))
                            { $i++ }
                        $outputFile = "$outputFile.$i"
                    }
                    else
                    {
                        Write-Host "Data already retrieved: $outputFile.txt"
                        $outputFile = ""
                    }
                }
                if (![String]::IsNullOrEmpty($outputFile))
                {
                    Write-Host "Saving data blob to: $outputFile.txt"
                    $auditData | Out-File -Filepath "$outputFile.txt" -NoClobber
                }
            }
        }
        else
        {
            Write-Host "No data returned from $contentUrl" -ForegroundColor Red
        }
    }
}

if ($RetrieveContent -and $contentUrls.Length -gt 0)
{    
    $script:contentRetrieved = 0
    $totalContentCount = $contentUrls.Length

    $activeJobs = New-Object System.Collections.ArrayList

    if ($ListContentDate)
    {
        $progressActivity = "Retrieving content made available between $startDate and $endDate"
    }
    else
    {
        $progressActivity = "Retrieving content from last 24 hours"
    }

    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0
    foreach ($contentUrl in $script:contentUrls)
    {
        $auth = "Bearer $(GetValidAccessToken)"
        LogVerbose "Downloading content blob: $contentUrl"
        $downloadJob = Start-Job -Scriptblock { param($url, $auth, $savePath) DownloadContentBlob $url $auth $savePath } -ArgumentList ($contentUrl, $auth, $SaveContentPath) -InitializationScript $downloadFunction
        $activeJobs.Add($downloadJob) | out-null
        while ($activeJobs.Count -ge $MaxRetrieveContentJobs)
        {
            # We have enough jobs running, so we need to wait until at least one has completed
            # We go through all of them and remove any that have finished
            for ($i=$activeJobs.Count-1; $i -ge 0; $i--)
            {
                if ($activeJobs[$i].State -ne "Running")
                {
                    Receive-Job $activeJobs[$i]
                    LogVerbose "Final job state: $($activeJobs[$i].State)"
                    $activeJobs.RemoveAt($i)
                }
            }
            if ($activeJobs.Count -ge $MaxRetrieveContentJobs)
            {
                # We didn't remove any jobs, so we'll sleep and try again
                Start-Sleep -Milliseconds 100
            }
        }

        $percentComplete = ($script:contentRetrieved/$totalContentCount)*100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete
    }

    # We now need to wait for all jobs to complete
    for ($i=$activeJobs.Count-1; $i -ge 0; $i--)
    {
        while ($activeJobs[$i].State -eq "Running")
        {
            LogVerbose "Waiting for job $i to finish"
            $activeJobs[$i] | Wait-Job | out-null
        }
        Receive-Job $activeJobs[$i]
        LogVerbose "Final job state: $($activeJobs[$i].State)"
        $script:contentRetrieved++
        $percentComplete = ($script:contentRetrieved/$totalContentCount)*100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete
    }
    Write-Progress -Activity $progressActivity -Status "100% complete" -Completed
}

Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) finished" Green