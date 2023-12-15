#
# Remove-DuplicateItems.ps1
#
# By David Barrett, Microsoft Ltd. 2017-2023. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

param (
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox will be accessed (instead of the main mailbox).")]
    [switch]$Archive,
		
    [Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox message root folder is assume.d")]
    [string]$FolderPath,

    [Parameter(Mandatory=$False,HelpMessage="Folder to which any duplicates will be moved.  If not specified, duplicate items are soft deleted (will go to Deleted Items folder).")]
    [string]$DuplicatesTargetFolder,

    [Parameter(Mandatory=$False,HelpMessage="When specified, any subfolders will be processed also.")]
    [switch]$RecurseFolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, duplicates will be matched anywhere within the mailbox (instead of just within the current folder).")]
    [switch]$MatchEntireMailbox,
	
    [Parameter(Mandatory=$False,HelpMessage="Only items created before the given date will be matched (though they can match items outside the date range).")]
    [string]$CreatedBefore,
    
    [Parameter(Mandatory=$False,HelpMessage="Only items created after the given date will be matched (though they can match items outside the date range).")]
    [string]$CreatedAfter,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder.")]
    [switch]$PublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="When speciifed, duplicate items will be hard deleted (normally they are only soft deleted).")]
    [switch]$HardDelete,

    [Parameter(Mandatory=$False,HelpMessage="When speciifed, the total number of duplicates found will be sent to the pipeline.")]
    [switch]$ReturnDuplicateCount,
    
#>** EWS/OAUTH PARAMETERS START **#
    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS.")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
    [Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for Office 365)")]
    [switch]$OAuth,

    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
    [string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id (application must be registered in the same tenant being accessed).")]
    [string]$OAuthTenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$OAuthSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.  Certificate auth requires MSAL libraries to be available.")]
    $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="If set, OAuth tokens will be stored in global variables for access in other scripts/console.  These global variable will be checked by later scripts using delegate auth to prevent additional log-in prompts.")]	
    [switch]$GlobalTokenStorage,

    [Parameter(Mandatory=$False,HelpMessage="For debugging purposes.")]
    [switch]$OAuthDebug,

    [Parameter(Mandatory=$False,HelpMessage="A value greater than 0 enables token debugging (specify total number of token renewals to debug).")]	
    $DebugTokenRenewal = 0,

    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox.")]
    [switch]$Impersonate,
	
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used).")]	
    [string]$EwsUrl,

    [Parameter(Mandatory=$False,HelpMessage="If specified, requests are directed to Office 365 endpoint (this overrides -EwsUrl).")]
    [switch]$Office365,
	
    [Parameter(Mandatory=$False,HelpMessage="If specified, only TLS 1.2 connections will be negotiated.")]
    [switch]$ForceTLS12,
	
    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed).")]	
    [string]$EWSManagedApiPath = "",
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate).")]	
    [switch]$IgnoreSSLCertificate,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing AutoDiscover.")]	
    [switch]$AllowInsecureRedirection,

    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file.")]	
    [string]$TraceFile,
#>** EWS/OAUTH PARAMETERS END **#

#>** LOGGING PARAMETERS START **#
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified.")]	
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="Enable verbose log file.  Verbose logging is written to the log whether -Verbose is enabled or not.")]	
    [switch]$VerboseLogFile,

    [Parameter(Mandatory=$False,HelpMessage="Enable debug log file.  Debug logging is written to the log whether -Debug is enabled or not.")]	
    [switch]$DebugLogFile,

    [Parameter(Mandatory=$False,HelpMessage="If selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled).")]
    [switch]$FastFileLogging,
#>** LOGGING PARAMETERS END **#
	
    [Parameter(Mandatory=$False,HelpMessage="Do not apply any changes, just report what would be updated")]	
    [switch]$WhatIf

)
$script:ScriptVersion = "1.2.0"
$script:debug = $false
$script:debugMaxItems = 3

# Define our functions

#>** LOGGING FUNCTIONS START **#
$scriptStartTime = [DateTime]::Now

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

Function UpdateDetailsWithCallingMethod([string]$Details)
{
    # Update the log message with details of the function that logged it
    $timeInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())"
    $callingFunction = (Get-PSCallStack)[2].Command # The function we are interested in will always be frame 2 on the stack
    if (![String]::IsNullOrEmpty($callingFunction))
    {
        return "$timeInfo [$callingFunction] $Details"
    }
    return "$timeInfo $Details"
}

Function LogToFile([string]$logInfo)
{
    if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    
    if ($FastFileLogging)
    {
        # Writing the log file using a FileStream (that we keep open) is significantly faster than using out-file (which opens, writes, then closes the file each time it is called)
        $fastFileLogError = $Error[0]
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            Write-Verbose "Opening/creating log file: $LogFile"
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
                $script:logFileStream = $null
            }
        }
        if ($script:logFileStream)
        {
            if (!$script:logFileStreamWriter)
            {
                $script:logFileStreamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            }
            $script:logFileStreamWriter.WriteLine($logInfo)
            $script:logFileStreamWriter.Flush()
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
            }
            else
            {
                return
            }
        }
    }

	$logInfo | Out-File $LogFile -Append
}

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Host $Details -ForegroundColor $Colour
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
    if ( !$VerboseLogFile -and !$DebugLogFile -and ($VerbosePreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
    if (!$DebugLogFile -and ($DebugPreference -eq "SilentlyContinue") ) { return }
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
        Log "ERROR ($Context): $($Error[0])" Red
    }
    else
    {
        $log = UpdateDetailsWithCallingMethod("ERROR: $($Error[0])")
        Log $log Red
    }
    return $true
}

Function ReportError($Context)
{
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}
#>** LOGGING FUNCTIONS END **#

#>** EWS/OAUTH FUNCTIONS START **#
function LoadLibraries()
{
    param (
        [bool]$searchProgramFiles,
        $dllNames,
        [ref]$dllLocations = @()
    )
    # Attempt to find and load the specified libraries

    foreach ($dllName in $dllNames)
    {
        # First check if the dll is in current directory
        LogDebug "Searching for DLL: $dllName"
        $dll = $null
        $dll = Get-ChildItem $dllName -ErrorAction Ignore

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
                if ($dllLocations)
                {
                    $dllLocations.value += $dll.VersionInfo.FileName
                    ReportError
                }
            }
            catch
            {
                ReportError "LoadLibraries"
                return $false
            }
        }
    }
    return $true
}

function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!$script:msalApiLoaded)
    {
        $msalLocation = @()
        $script:msalApiLoaded = $(LoadLibraries -searchProgramFiles $false -dllNames @("Microsoft.Identity.Client.dll") -dllLocations ([ref]$msalLocation))
        if (!$script:msalApiLoaded)
        {
            Log "Failed to load MSAL.  Cannot continue with certificate authentication." Red
            exit
        }
    }   

    $cca = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca = $cca.WithCertificate($OAuthCertificate)
    $cca = $cca.WithTenantId($OAuthTenantId)
    $cca = $cca.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    LogVerbose "Requesting token using certificate auth"
    $script:oauthToken = $acquire.ExecuteAsync().Result
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    $script:Impersonate = $true
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$OAuthRedirectUri&response_mode=query&scope=openid%20profile%20email%20offline_access%20https://outlook.office365.com/.default"
    Write-Host "Please complete log-in via the web browser, and then copy the redirect URL (including auth code) to the clipboard to continue" -ForegroundColor Green
    Set-Clipboard -Value "Waiting for auth code"
    Start-Process $authUrl

    do
    {
        $authcode = Get-Clipboard
        Start-Sleep -Milliseconds 250
    } while ($authCode -eq "Waiting for auth code")

    $codeStart = $authcode.IndexOf("?code=")
    if ($codeStart -gt 0)
    {
        $authcode = $authcode.Substring($codeStart+6)
        $codeEnd = $authcode.IndexOf("&session_state=")
        if ($codeEnd -gt 0)
        {
            $authcode = $authcode.Substring(0, $codeEnd)
        }
        Write-Verbose "Using auth code: $authcode"
    }
    else
    {
        throw "Failed to obtain Auth code from clipboard"
    }

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$authcode;redirect_uri=$OAuthRedirectUri}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
        return
    }
    catch {}

    throw "Failed to obtain OAuth token"
}

function RenewOAuthToken
{
    # Renew the delegate token (original token was obtained by auth code, but we can now renew using the access token)
    if (!$script:oAuthToken)
    {
        # We don't have a token, so we can't renew
        GetTokenViaCode
        return
    }

    $body = @{grant_type="refresh_token";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;refresh_token=$script:oauthToken.refresh_token}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to renew OAuth token (auth code grant)" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}

function GetTokenWithKey
{
    $Body = @{
      "grant_type"    = "client_credentials";
      "client_id"     = "$OAuthClientId";
      "scope"         = "https://outlook.office365.com/.default"
    }

    if ($script:oAuthToken -ne $null)
    {
        # If we have a refresh token, add that to our request body and change grant type
        if (![String]::IsNullOrEmpty($script:oAuthToken.refresh_token))
        {
            $Body.Add("refresh_token", $script:oAuthToken.refresh_token)
            $Body["grant_type"] = "refresh_token"
        }
    }
    if ($Body["grant_type"] -eq "client_credentials")
    {
        # To obtain our first access token we need to use the secret key
        $Body.Add("client_secret", $OAuthSecretKey)
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Log "Failed to obtain OAuth token: $Error" Red
        exit # Failed to obtain a token
    }
    $script:Impersonate = $true
}

function JWTToPSObject
{
    param([Parameter(Mandatory=$true)][string]$token)

    $tokenheader = $token.Split(".")[0].Replace('-', '+').Replace('_', '/')
    while ($tokenheader.Length % 4) { $tokenheader = "$tokenheader=" }    
    $tokenHeaderObject = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    while ($tokenPayload.Length % 4) { $tokenPayload = "$tokenPayload=" }
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)
    $tokenObject = $tokenArray | ConvertFrom-Json
    return $tokenObject
}

function LogOAuthTokenInfo
{
    if ($global:OAuthAccessToken -eq $null)
    {
        Log "No OAuth token obtained." Red
        return
    }

    $idToken = $null
    if (-not [String]::IsNullOrEmpty($global:OAuthAccessToken.id_token))
    {
        $idToken = $global:OAuthAccessToken.id_token
    }
    elseif (-not [String]::IsNullOrEmpty($global:OAuthAccessToken.IdToken))
    {
        $idToken = $global:OAuthAccessToken.IdToken
    }

    if ([String]::IsNullOrEmpty($idToken))
    {
        Log "OAuth ID token not present" Yellow
    }
    else
    {
        $global:idTokenDecoded = JWTToPSObject($idToken)
        Log "OAuth ID Token (`$idTokenDecoded):" Yellow
        Log $global:idTokenDecoded Yellow
    }

    if (-not [String]::IsNullOrEmpty($global:OAuthAccessToken))
    {
        $global:accessTokenDecoded = JWTToPSObject($global:OAuthAccessToken)
        Log "OAuth Access Token (`$accessTokenDecoded):" Yellow
        Log $global:accessTokenDecoded Yellow
    }
    else
    {
        Log "OAuth access token not present" Red
    }
}

function GetOAuthCredentials
{
    # Obtain OAuth token for accessing mailbox
    param (
        [switch]$RenewToken
    )
    $exchangeCredentials = $null

    if ($script:oauthToken -ne $null)
    {
        # We already have a token
        if ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1))
        {
            # Token still valid, so return that
            $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
            return $exchangeCredentials
        }
        # Token needs renewing
    }

    if (![String]::IsNullOrEmpty($OAuthSecretKey))
    {
        GetTokenWithKey
    }
    elseif ($OAuthCertificate -ne $null)
    {
        GetTokenWithCertificate
    }
    else
    {
        if ($RenewToken)
        {
            RenewOAuthToken
        }
        else
        {
            if ($GlobalTokenStorage -and $script:oauthToken -eq $null)
            {
                # Check if we have token variable set globally
                if ($global:oAuthPersistAppId -eq $OAuthClientId)
                {
                    $script:oAuthToken = $global:oAuthPersistToken
                    $script:oauthTokenAcquireTime = $global:oAuthPersistTokenAcquireTime
                }
                RenewOAuthToken
            }
            else
            {
                GetTokenViaCode
            }
        }
    }

    if ($GlobalTokenStorage -or $OAuthDebug)
    {
        # Store the OAuth in a global variable for later access
        $global:oAuthPersistToken = $script:oAuthToken
        $global:oAuthPersistAppId = $OAuthClientId
        $global:oAuthPersistTokenAcquireTime = $script:oauthTokenAcquireTime
    } 

    if ($OAuthDebug)
    {
        LogVerbose "`$oAuthPersistToken contains token response"
        $global:OAuthAccessToken = $script:oAuthAccessToken
        LogVerbose "`$OAuthAccessToken: `r`n$($global:OAuthAccessToken)"
        LogOAuthTokenInfo
    }

   

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

$script:oAuthDebugStop = $false
$script:oAuthDebugStopCount = 0
function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $script:services -eq $null ) { return }

    
    if ($DebugTokenRenewal -gt 0 -and $script:oauthToken)
    {
        # When debugging tokens, we stop after on every other EWS call and wait for the token to expire
        if ($script:oAuthDebugStop)
        {
            # Wait until token expires (we do this after every call when debugging OAuth)
            # Access tokens can't be revoked, but a policy can be assigned to reduce lifetime to 10 minutes: https://learn.microsoft.com/en-us/graph/api/resources/tokenlifetimepolicy?view=graph-rest-1.0
            if ($OAuthCertificate -ne $null)
            {
                $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
            }
            else
            {
                $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
            }
            $timeUntilExpiry = $tokenExpire.Subtract([DateTime]::UtcNow).TotalSeconds
            if ($timeUntilExpiry -gt 0)
            {
                Write-Host "Waiting until token has expired: $tokenExpire (UTC)" -ForegroundColor Cyan
                Start-Sleep -Seconds $tokenExpire.Subtract([DateTime]::UtcNow).TotalSeconds
            }
            Write-Host "Token expired, continuing..." -ForegroundColor Cyan
            $oAuthDebugStop = $false
            $script:oAuthDebugStopCount++
        }
        else
        {
            if ($DebugTokenRenewal-$script:oAuthDebugStopCount -gt 0)
            {
                $script:oAuthDebugStop = $true
            }
        }
    }
    
    if ($OAuthCertificate -ne $null)
    {
        if ( [DateTime]::UtcNow -lt $script:oauthToken.ExpiresOn.UtcDateTime) { return }
    }
    elseif ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    LogVerbose("[ApplyEWSOAuthCredentials] OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($exchangeCredentials -eq $null) { return }

    if ($OAuthCertificate -ne $null)
    {
        $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
        if ( [DateTime]::UtcNow -ge $tokenExpire)
        {
            Log "[ApplyEWSOAuthCredentials] OAuth Token renewal failed (certificate auth)"
            exit # We no longer have access to the mailbox, so we stop here
        }
    }
    else
    {
        if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -lt [DateTime]::UtcNow )
        { 
            Log "[ApplyEWSOAuthCredentials] OAuth Token renewal failed"
            exit # We no longer have access to the mailbox, so we stop here
        }
        $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
    }

    Log "[ApplyEWSOAuthCredentials] OAuth token successfully renewed; new expiry: $tokenExpire"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = $exchangeCredentials
        }
        LogVerbose "[ApplyEWSOAuthCredentials] Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
    }
}

Function LoadEWSManagedAPI
{
	# Find and load the managed API
    $ewsApiLocation = @()
    $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
    ReportError "LoadEWSManagedAPI"

    if (!$ewsApiLoaded)
    {
        # Failed to load the EWS API, so try to install it from Nuget
        $ewsapi = Find-Package "Exchange.WebServices.Managed.Api"
        if ($ewsapi.Entities.Name.Equals("Microsoft"))
        {
	        # We have found EWS API package, so install as current user (confirm with user first)
	        Write-Host "EWS Managed API is not installed, but is available from Nuget.  Install now for current user (required for this script to continue)? (Y/n)" -ForegroundColor Yellow
	        $response = Read-Host
	        if ( $response.ToLower().Equals("y") )
	        {
		        Install-Package $ewsapi -Scope CurrentUser -Force
                $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
                ReportError "LoadEWSManagedAPI"
	        }
        }
    }

    if ($ewsApiLoaded)
    {
        if ($ewsApiLocation[0])
        {
            Log "Using EWS Managed API found at: $($ewsApiLocation[0])" Gray
            $script:EWSManagedApiPath = $ewsApiLocation[0]
        }
        else
        {
            Write-Host "Failed to read EWS API location: $ewsApiLocation"
            Exit
        }
    }

    return $ewsApiLoaded
}

Function CurrentUserPrimarySmtpAddress()
{
    # Attempt to retrieve the current user's primary SMTP address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $result = $searcher.FindOne()

    if ($result -ne $null)
    {
        $mail = $result.Properties["mail"]
        LogDebug "Current user's SMTP address is: $mail"
        return $mail
    }
    return $null
}

Function TrustAllCerts()
{
    # Implement call-back to override certificate handling (and accept all)

    $TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
            public class TrustAll : System.Net.ICertificatePolicy {
                public TrustAll()
                {
                }
                public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                    System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                    System.Net.WebRequest req, int problem)
                {
                    return true;
                }
            }
        }
'@ 

    Add-Type -TypeDefinition $TASource -ReferencedAssemblies "System.DLL"

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=[Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll]::new()
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

Function CreateTraceListener($service)
{
    # Create trace listener to capture EWS conversation (useful for debugging)

    if ([String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        Log "Managed API path missing; unable to create tracer" Red
        Exit
    }

    if ($script:Tracer -eq $null)
    {
        $traceFileForCode = ""

        if (![String]::IsNullOrEmpty($TraceFile))
        {
            Log "Tracing to: $TraceFile"
            $traceFileForCode = $traceFile.Replace("\", "\\")
        }

        $TraceListenerClass = @"
		    using System;
		    using System.Text;
		    using System.IO;
		    using System.Threading;
		    using Microsoft.Exchange.WebServices.Data;
		
		    public class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
		    {
			    private StreamWriter _traceStream = null;
                private string _lastResponse = String.Empty;

			    public EWSTracer()
			    {
"@
    if (![String]::IsNullOrEmpty(($traceFileForCode)))
    {
        $TraceListenerClass = 
@"
$TraceListenerClass
				    try
				    {
					    _traceStream = File.AppendText("$traceFileForCode");
				    }
				    catch { }
"@
    }

        $TraceListenerClass = 
@"
$TraceListenerClass			        }

			    ~EWSTracer()
			    {
                    Close();
			    }

                public void Close()
			    {
				    try
				    {
					    _traceStream.Flush();
					    _traceStream.Close();
				    }
				    catch { }
			    }


			    public void Trace(string traceType, string traceMessage)
			    {
                    if ( traceType.Equals("EwsResponse") )
                        _lastResponse = traceMessage;

                    if ( traceType.Equals("EwsRequest") )
                        _lastResponse = String.Empty;

				    if (_traceStream == null)
					    return;

					try
					{
						_traceStream.WriteLine(traceMessage);
						_traceStream.Flush();
					}
					catch { }
			    }

                public string LastResponse
                {
                    get { return _lastResponse; }
                }
		    }
"@

        Add-Type -TypeDefinition $TraceListenerClass -ReferencedAssemblies $EWSManagedApiPath
        $script:Tracer=[EWSTracer]::new()

        # Attach the trace listener to the Exchange service
        $service.TraceListener = $script:Tracer
    }
}

function CreateService($smtpAddress, $impersonatedAddress = "")
{
    # Creates and returns an ExchangeService object to be used to access mailboxes

    # First of all check to see if we have a service object for this mailbox already
    if ($script:services -eq $null)
    {
        $script:services = @{}
    }
    if ($script:services.ContainsKey($smtpAddress))
    {
        return $script:services[$smtpAddress]
    }

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)

    # Do we need to use OAuth?
    if ($Office365) { $OAuth = $true }
    if ($OAuth)
    {
        $exchangeService.Credentials = GetOAuthCredentials
        if ($exchangeService.Credentials -eq $null)
        {
            # OAuth failed
            return $null
        }
    }
    else
    {
        # Set credentials if specified, or use logged on user.
        if ($Credentials -ne $Null)
        {
            LogVerbose "Applying given credentials: $($Credentials.UserName)"
            $exchangeService.Credentials = $Credentials.GetNetworkCredential()
        }
        else
        {
	        LogVerbose "Using default credentials"
            $exchangeService.UseDefaultCredentials = $true
        }
    }

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl -or $Office365)
    {
        if ($Office365) { $EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx" }
    	$exchangeService.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
		    LogVerbose "Performing autodiscover for $smtpAddress"
		    if ( $AllowInsecureRedirection )
		    {
			    $exchangeService.AutodiscoverUrl($smtpAddress, {$True})
		    }
		    else
		    {
			    $exchangeService.AutodiscoverUrl($smtpAddress)
		    }
		    if ([string]::IsNullOrEmpty($exchangeService.Url))
		    {
			    Log "$smtpAddress : autodiscover failed" Red
			    return $Null
		    }
		    LogVerbose "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
            Log "$smtpAddress : error occurred during autodiscover: $($Error[0])" Red
            return $null
    	}
    }
 
    if ([String]::IsNullOrEmpty($impersonatedAddress))
    {
        $impersonatedAddress = $smtpAddress
    }
    $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $impersonatedAddress)
	}

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    if (![String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        CreateTraceListener $exchangeService
        if ($script:Tracer)
        {
            $exchangeService.TraceListener = $script:Tracer
            $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
            $exchangeService.TraceEnabled = $True
        }
        else
        {
            Log "Failed to create EWS trace listener.  Throttling back-off time won't be detected." Yellow
        }
    }

    $script:services.Add($smtpAddress, $exchangeService)
    LogVerbose "Currently caching $($script:services.Count) ExchangeService objects" $true
    return $exchangeService
}

#>** EWS/OAUTH FUNCTIONS END **#

function GetFolderPath($Folder)
{
    # Return the full path for the given folder

    # We cache our folder lookups for this script
    if (!$script:folderCache)
    {
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive
        # We use a .Net Dictionary object instead
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
    $parentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $Folder.Id, $propset)
    $folderPath = $Folder.DisplayName
    $parentFolderId = $Folder.Id
    while ($parentFolder.ParentFolderId -ne $parentFolderId)
    {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
        {
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
        }
        else
        {
            $parentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $parentFolder.ParentFolderId, $propset)
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function IsDuplicateAppointment($item)
{
    # Test for duplicate appointment
    $isDupe = $False
    if ($script:icaluids.ContainsKey($item.ICalUid))
    {
        # Duplicate ICalUid exists
        LogDebug "Matched on iCalUid: $($item.ICalUid)"
        return $True
    }
    else
    {
        $script:icaluids.Add($item.ICalUid, $item.Id.UniqueId)

        $subject_cmp = $item.Subject
        if ([String]::IsNullOrEmpty($subject_cmp))
        {
            $subject_cmp = "[No Subject]" # If the subject is blank, we need to give it an arbitrary value to prevent checks failing
        }
        if ($script:calsubjects.ContainsKey($subject_cmp))
        {
            # Duplicate subject exists, so we now check the start and end date to confirm if this is a duplicate
            $dupSubjects = $script:calsubjects[$subject_cmp]
            LogDebug "$($dupSubjects.Count) matching appointment subjects: $subject_cmp"
            foreach ($dupSubject in $dupSubjects)
            {
                if (($dupSubject.Start -eq $item.Start) -and ($dupSubject.End -eq $item.End))
                {
                    # Same subject, start, and end date, so this is a duplicate
                    LogVerbose "Duplicate appointment found: $subject_cmp"
                    return $true
                }
                else
                {
                    LogDebug "Start: $($dupSubject.Start) and $($item.Start)    End: $($dupSubject.End) and $($item.End)"
                }
            }
            # Add this item to the list of items with the same subject (as it is not a duplicate)
            $script:calsubjects[$subject_cmp] += $item
        }
        else
        {
            # Add this to our subject list
            LogDebug "New appointment subject: $subject_cmp"
            $script:calsubjects.Add($subject_cmp, @($item))
        }
    }
    return $false
}

function IsDuplicateContact($item)
{
    # Test for duplicate contact
    $item.Load([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    if (![String]::IsNullOrEmpty(($item.DisplayName)))
    {
        if (!$script:displayNames.ContainsKey($item.DisplayName))
        {
            # No duplicate contact display name found, so we add this one to our list
            LogDebug "New contact DisplayName: $($item.DisplayName)"
            $script:displayNames.Add($item.DisplayName, @($item))
            return $false
        }
    }
    else
    {
        # If display name is empty, we do not count this as a duplicate (we ignore it)
        return $false
    }

    # We have another contact with same display name, so we now need to check other fields to confirm match

    $possibleMatches = $script:displayNames[$item.DisplayName]
    LogDebug "$($possibleMatches.Count) matching contact names: $($item.DisplayName)"

    foreach ($possibleMatch in $possibleMatches)
    {
        $match = $true
        if ($item.EmailAddress1 -ne $possibleMatch.EmailAddress1) { $match = $false }
        if ($item.EmailAddress2 -ne $possibleMatch.EmailAddress2) { $match = $false }
        if ($item.EmailAddress3 -ne $possibleMatch.EmailAddress3) { $match = $false }
        if ($item.ImAddress1 -ne $possibleMatch.ImAddress1) { $match = $false }
        if ($item.ImAddress2 -ne $possibleMatch.ImAddress2) { $match = $false }
        if ($item.ImAddress3 -ne $possibleMatch.ImAddress3) { $match = $false }
        if ($item.BusinessPhone -ne $possibleMatch.BusinessPhone) { $match = $false }
        if ($item.BusinessPhone2 -ne $possibleMatch.BusinessPhone2) { $match = $false }
        if ($item.CompanyName -ne $possibleMatch.CompanyName) { $match = $false }
        if ($item.HomePhone -ne $possibleMatch.HomePhone) { $match = $false }
        if ($item.HomePhone2 -ne $possibleMatch.HomePhone2) { $match = $false }
        if ($item.MobilePhone -ne $possibleMatch.MobilePhone) { $match = $false }
        if ($item.Birthday -ne $possibleMatch.Birthday) { $match = $false }

        if ($match)
        {
            LogVerbose "Duplicate contact found: $($item.DisplayName)"
            return $true
        }
    }

    # This isn't a duplicate, so we want to add it to our list of possible duplicates with the same display name
    $script:displayNames[$item.DisplayName] += $item
    return $false
}

function IsDuplicateEmail($item)
{
    # Test for duplicate email
    $isDupe = $False

    if (![String]::IsNullOrEmpty(($item.InternetMessageId)))
    {
        if ($script:imsgids.ContainsKey($item.InternetMessageId))
        {
            # Duplicate Internet Message Id exists
            LogDebug "Matched on InternetMessageId: $($item.InternetMessageId)"
            return $True
        }
        $script:imsgids.Add($item.InternetMessageId, $item.Id.UniqueId)
    }

    $subject_cmp = $item.Subject
    if ([String]::IsNullOrEmpty($subject_cmp))
    {
        $subject_cmp = "[No Subject]" # If the subject is blank, we need to give it an arbitrary value to prevent checks failing
    }
    if ($script:msgsubjects.ContainsKey($subject_cmp))
    {
        # Duplicate subject exists, so we now check the start and end date to confirm if this is a duplicate
        $dupSubjects = $script:msgsubjects[$subject_cmp]
        foreach ($dupSubject in $dupSubjects)
        {
            if ($item.ItemClass -eq $dupSubject.ItemClass)
            {
                if ($item.IsFromMe)
                {
                    # This is a sent item
                    if (($dupSubject.DateTimeSent -eq $item.DateTimeSent) -and ($dupSubject.IsFromMe))
                    {
                        # Same subject and sent date, so this is a duplicate
                        LogVerbose "Duplicate sent email found: $subject_cmp"
                        return $true
                    }
                }
                else
                {
                    # This is a received item
                    if (($dupSubject.DateTimeReceived -eq $item.DateTimeReceived) -and (!$dupSubject.IsFromMe))
                    {
                        # Same subject and received date, so this is a duplicate
                        LogVerbose "Duplicate received email found: $subject_cmp"
                        return $true
                    }
                }
            }
        }
        # Add this item to the list of items with the same subject (as it is not a duplicate)
        $script:msgsubjects[$subject_cmp] += $item
    }
    else
    {
        # Add this to our subject list
        $script:msgsubjects.Add($subject_cmp, @($item))
    }
    return $false
}

function IsDuplicate($item)
{
    # Test if item is duplicate (the check we do depends upon the item type)

    if ($script:createdBeforeDate -ne $null -and $item.DateTimeCreated -ge $script:createdBeforeDate)
    {
        LogVerbose "Item is outside date range, so will not be considered a duplicate (therefore skipping checks)"
        return $false
    }
    if ($script:createdAfterDate -ne $null -and $item.DateTimeCreated -le $script:createdAfterDate)
    {
        LogVerbose "Item is outside date range, so will not be considered a duplicate (therefore skipping checks)"
        return $false
    }

    if ($item.ItemClass.StartsWith("IPM.Note") -or $item.ItemClass.StartsWith("IPM.Schedule.MeetingRequest"))
    {
        return IsDuplicateEmail($item)
    }
    if ($item.ItemClass.Equals("IPM.Appointment"))
    {
        return IsDuplicateAppointment($item)
    }
    if ($item.ItemClass.Equals("IPM.Contact"))
    {
        return IsDuplicateContact($item)
    }
    LogVerbose "Unsupported item type being ignored: $($item.ItemClass)"
    return $false
}

function SearchForDuplicates($folder)
{
    # Search the folder for duplicate appointments
    # We read all the items in the folder, and build a list of all the duplicates

    # First of all, check if we are recursing and process any subfolders first
    if ($RecurseFolders)
    {
		if ($folder.ChildFolderCount -gt 0)
		{
			# Deal with any subfolders first
            LogVerbose "Processing subfolders of $($folder.DisplayName)"
			$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
			$FolderView.PropertySet = $script:requiredFolderProperties
			$findFolderResults = $folder.FindFolders($FolderView)
			ForEach ($subfolder in $findFolderResults.Folders)
			{
                if ($subFolder.ExtendedProperties[0].Value -ne 2) # Ignore search folders
                {
                    SearchForDuplicates $subfolder
                }
			}
		}
    }

    if (!$MatchEntireMailbox -or ($script:calsubjects -eq $Null))
    {
        # Clear/initialise the duplicate tracking lists (we are only checking for duplicates within a folder)
        $script:calsubjects = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
        $script:msgsubjects = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
        $script:icaluids = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
        $script:imsgids = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
        $script:displayNames = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }
    $dupeCount = 0

    # Progress Reporting
    $itemCount = $folder.TotalCount
    $processedCount = 0
    $activity = "Reading items in folder $($folder.DisplayName)"

    $offset = 0
    $moreItems = $true
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, 0)
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid)
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId)
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived)
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent)
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsFromMe)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)

    $view.PropertySet = $propset

    Write-Progress -Activity $activity -Status "0 of $itemCount items read" -PercentComplete 0

    while ($moreItems)
    {
        $results = $Folder.FindItems($view)
        $moreItems = $results.MoreAvailable
        $view.Offset = $results.NextPageOffset
        foreach ($item in $results)
        {
            LogVerbose "Processing: $($item.Subject)"
            If (IsDuplicate($item))
            {
                if ($script:debug)
                {
                    if ($script:debugMaxItems-- -gt 0)
                    {
                        $script:duplicateItems += $item
                    }
                }
                else
                {
                    $script:duplicateItems += $item
                }
                $dupeCount++
            }
            $processedCount++
            if ($processedCount % 100 -eq 0)
            {
                Write-Progress -Activity $activity -Status "$processedCount of $itemCount item(s) read" -PercentComplete (($processedCount/$itemCount)*100)
            }
        }
    }
    Write-Progress -Activity $activity -Completed

    if ($dupeCount -eq 0)
    {
        Log "No duplicate items found in folder $($folder.Displayname)" Green
        return
    }
    Log "$dupeCount duplicate(s) found in folder $($folder.Displayname)" White

}

$script:itemRetryCount = @{}
Function RemoveProcessedItemsFromList()
{
    # Process the results of a batch move/copy and remove any items that were successfully moved from our list of items to move
    param (
        $requestedItems,
        $results,
        $suppressErrors = $false,
        $Items
    )

    if ($results -ne $null)
    {
        $failed = 0
        for ($i = 0; $i -lt $requestedItems.Count; $i++)
        {
            if ($results[$i].ErrorCode -eq "NoError")
            {
                LogVerbose "Item successfully processed: $($requestedItems[$i])"
                [void]$Items.Remove($requestedItems[$i])
            }
            else
            {
                $permanentErrors = @("ErrorMoveCopyFailed", "ErrorInvalidOperation", "ErrorItemNotFound", "ErrorToFolderNotFound")
                if ( $permanentErrors.Contains($results[$i].ErrorCode.ToString()) )
                {
                    # This is a permanent error, so we remove the item from the list
                    [void]$Items.Remove($requestedItems[$i])
                    if (!$suppressErrors)
                    {
                        if ([String]::IsNullOrEmpty($results[$i].MessageText))
                        {
                            Log "Permanent error $($results[$i].ErrorCode) reported for item: $($requestedItems[$i].UniqueId)" Red
                        }
                        else
                        {
                            Log "Permanent error $($results[$i].ErrorCode) ($($results[$i].MessageText)) reported for item: $($requestedItems[$i].UniqueId)" Red
                        }
                    }
                }
                else
                {
                    # This is most likely a temporary error, so we don't remove the item from the list
                    $retryCount = 0
                    if ( $script:itemRetryCount.ContainsKey($requestedItems[$i].UniqueId) )
                        { $retryCount = $script:itemRetryCount[$requestedItems[$i].UniqueId] }
                    $retryCount++
                    if ($retryCount -lt 4)
                    {
                        LogVerbose "Error $($results[$i].ErrorCode) ($($results[$i].MessageText)) reported for item (attempt $retryCount): $($requestedItems[$i].UniqueId)"
                        $script:itemRetryCount[$requestedItems[$i].UniqueId] = $retryCount
                    }
                    else
                    {
                        # We got an error 3 times in a row, so we'll admit defeat
                        [void]$Items.Remove($requestedItems[$i])
                        if (!$suppressErrors)
                        {
                            if ([String]::IsNullOrEmpty($results[$i].MessageText))
                            {
                                Log "Permanent error $($results[$i].ErrorCode) reported for item: $($requestedItems[$i].UniqueId)" Red
                            }
                            else
                            {
                                Log "Permanent error $($results[$i].ErrorCode) ($($results[$i].MessageText)) reported for item: $($requestedItems[$i].UniqueId)" Red
                            }
                        }
                    }
                }
                $failed++
            } 
        }
    }
    else
    {
        Log "No results returned - assuming all items processed" White
        for ($i = 0; $i -lt $requestedItems.Count; $i++)
        {
            [void]$Items.Remove($requestedItems[$i])
        }

    }
    if ( ($failed -gt 0) -and !$suppressErrors )
    {
        Log "$failed items reported error during batch request (if throttled, some failures are expected, and will be retried)" Yellow
    }
    else
    {
        LogVerbose "All batch items processed successfully"
    }
}

Function ThrottledBatchMove()
{
    # Send request to move/copy items, allowing for throttling (which in this case is likely to manifest as time-out errors)
    param (
        $ItemsToMove,
        $TargetFolderId
    )

    $consecutive401Errors = 0

	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $totalItems = $ItemsToMove.Count
    Write-Progress -Activity "Moving items" -Status "0% complete" -PercentComplete 0

    while ( !$finished )
    {
	    $script:moveIds = [Activator]::CreateInstance($genericItemIdList)

        LogVerbose "Current batch size is $($script:currentBatchSize)"
        
        for ([int]$i=0; $i -lt $script:currentBatchSize; $i++)
        {
            if ($ItemsToMove[$i] -ne $null)
            {
                if ($moveIds.Contains($ItemsToMove[$i]))
                {
                    LogVerbose "Item already in batch: $ItemsToMove[$i]"
                    $ItemsToMove.Remove($ItemsToMove[$i])
                    if ($i -gt 0) { $i-- }
                }
                else
                {
                    $moveIds.Add($ItemsToMove[$i])
                    LogVerbose "Added to batch: $($ItemsToMove[$i])"
                }
            }
            else
            {
                LogVerbose "Ignored null source item (index $i)"
            }
            if ($i -ge $ItemsToMove.Count)
                { break }
        }

        $results = $null
        try
        {
            LogVerbose "Sending batch request to move $($moveIds.Count) items ($($ItemsToMove.Count) remaining)"
			$results = $script:service.MoveItems( $moveIds, $TargetFolderId, $false)
            LogVerbose "Batch request completed"
        }
        catch
        {
            if ( Throttled )
            {
                # We've been throttled, which should now have expired (the Throttled function waits as necessary), so can now retry
            }
            else
            {                
                if ($Error[0].Exception)
                {
                    if ($Error[0].Exception.Message.Contains("(401) Unauthorized."))
                    {
                        # This is most likely an issue with the OAuth token.
                        $consecutive401Errors++
                        if ( ($consecutive401Errors -lt 2) -and $OAuth)
                        {
                            Log "Access denied response - checking OAuth token"
                            ApplyEWSOauthCredentials
                        }
                        else
                        {
                            Log "Consecutive access denied errors encountered - stopping processing" Red
                            Exit
                        }
                    }
                    elseif ($Error[0].Exception.InnerException.ToString().Contains("The operation has timed out"))
                    {
                        # We've probably been throttled, so we'll reduce the batch size and try again
                        if ($script:currentBatchSize -gt 50)
                        {
                            LogVerbose "Timeout error received"
                            DecreaseBatchSize
                        }
                        else
                        {
                            $finished = $true
                        }
                    }
                    else
                    {
                        LogVerbose "ERROR ON MOVE: $($Error[0].Exception.Message)"
                    }
                }
                else
                {
                    $finished = $true # Unknown error, so we finish processing as we don't know the best way to handle it
                    $lastResponse = [String]::Empty
                    try
                    {
                        if ($script:Tracer)
                        {
                            $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
                        }
                    } catch {}
                    if (![String]::IsNullOrEmpty($lastResponse))
                    {
                        $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
                        $responseXml = [xml]$lastResponse
	                    if ($responseXml.Trace.Envelope.Body.Fault.detail.ResponseCode.Value -eq "ErrorNoRespondingCASInDestinationSite")
                        {
                            # We get this error if the destination CAS (request was proxied) hasn't returned any data within the timeout (usually 60 seconds)
                            # Reducing the batch size should help here, and we want to reduce it quite aggressively
                            if ($script:currentBatchSize -gt 50)
                            {
                                LogVerbose "ErrorNoRespondingCASInDestinationSite error received"
                                DecreaseBatchSize 0.7
                                $finished = $false
                            }
                        }
                        else
                        {
                            ReportError "ThrottledBatchMove"
                        }
                    }
                    else
                    {
                        ReportError "ThrottledBatchMove"
                    }
                }
            }
        }
        ApplyEWSOauthCredentials

        RemoveProcessedItemsFromList $moveIds $results $false $ItemsToMove

        $percentComplete = ( ($totalItems - $ItemsToMove.Count) / $totalItems ) * 100
        Write-Progress -Activity "Moving items" -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToMove.Count -eq 0)
        {
            $finished = $True
            Write-Progress -Activity "Moving items" -Status "100% complete" -Completed
        }
    }
}


Function ThrottledBatchDelete()
{
    # Send request to delete items, allowing for throttling (which in this case is likely to manifest as time-out errors)
    param (
        $ItemsToDelete,
        $BatchSize = 500,
        $SuppressNotFoundErrors = $false
    )

    if ($script:MaxBatchSize -gt 0)
    {
        # If we've had to reduce the batch size previously, we'll start with the last size that was successful
        $BatchSize = $script:MaxBatchSize
    }

    if ($ItemsToDelete.Count -lt 1)
    {
        return
    }

    $progressActivity = "Deleting items"
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $totalItems = $ItemsToDelete.Count
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0
    $consecutiveErrors = 0

    while ( !$finished )
    {
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        
        for ([int]$i=0; $i -lt $BatchSize; $i++)
        {
            if ($ItemsToDelete[$i] -ne $null)
            {
                $deleteIds.Add($ItemsToDelete[$i])
            }
            if ($i -ge $ItemsToDelete.Count)
                { break }
        }

        $results = $null
        try
        {
            LogVerbose "Sending batch request to delete $($deleteIds.Count) items ($($ItemsToDelete.Count) remaining)"
			$results = $script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
            $consecutiveErrors = 0 # Reset the consecutive error count, as if we reach this point then this request succeeded with no error
        }
        catch
        {
            # We reduce the batch size if we encounter an error (sometimes throttling does not return a throttled response, this can happen if the EWS request is proxied, and the proxied request times out)
            if ($BatchSize -gt 50)
            {
                $BatchSize = [int]($BatchSize * 0.8)
                $script:MaxBatchSize = $BatchSize
                LogVerbose "Batch size reduced to $BatchSize"
            }
            else
            {
                # If we've already reached a batch size of 50 or less, we set it to 10 (this is the minimum we reduce to)
                if ($BatchSize -ne 10)
                {
                    $BatchSize = 10
                    LogVerbose "Batch size set to 10"
                }
            }
            if ( -not (Throttled) )
            {
                $consecutiveErrors++
                try
                {
                    Log "Unexpected error: $($Error[0].Exception.InnerException.ToString())" Red
                }
                catch
                {
                    Log "Unexpected error: $($Error[1])" Red
                }
                $finished = ($consecutiveErrors -gt 9) # If we have 10 errors in a row, we stop processing
            }
            ApplyEWSOauthCredentials
        }

        RemoveProcessedItemsFromList $deleteIds $results $SuppressNotFoundErrors $ItemsToDelete

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
}

Function ProcessWhatIf($processType)
{
	ForEach ($dupe in $script:duplicateItems)
	{
        if ([String]::IsNullOrEmpty($dupe.Subject))
        {
            Log "Would $($processType): [No Subject]" Gray
        }
        else
        {
            Log "Would $($processType): $($dupe.Subject)" Gray
        }
    }
}

Function BatchDeleteDuplicates()
{
    # Delete all identified duplicates

    if ( $WhatIf )
    {
	    ProcessWhatIf("delete")
        return
    }

    Log "Deleting $($script:duplicateItems.Count) items"
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    $batchDeleteIds = [Activator]::CreateInstance($genericItemIdList)

    foreach ($item in $script:duplicateItems)
    {
        $batchDeleteIds.Add($item.Id)
    }
    ThrottledBatchDelete $batchDeleteIds

    if ($batchDeleteIds.Count -eq 0)
    {
        Log "All items successfully deleted" Green
    }
    else
    {
        Log "$($batchDeleteIds.Count) items were not deleted" Yellow
    }
}

Function BatchMoveDuplicates()
{
    # Move all the duplicates to the specified folder

    if ( $WhatIf )
    {
	    ProcessWhatIf("move")
        return
    }

    Log "Moving $($script:duplicateItems.Count) items"
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    $batchMoveIds = [Activator]::CreateInstance($genericItemIdList)

    foreach ($item in $script:duplicateItems)
    {
        $batchMoveIds.Add($item.Id)
    }
    ThrottledBatchMove $batchMoveIds $script:moveToFolder.Id

    if ($batchMoveIds.Count -eq 0)
    {
        Log "All items successfully moved" Green
    }
    else
    {
        Log "$($batchMoveIds.Count) items were not deleted" Yellow
    }
}


Function DecreaseBatchSize()
{
    param (
        $DecreaseMultiplier = 0.8
    )

    $script:currentBatchSize = [int]($script:currentBatchSize * $DecreaseMultiplier)
    if ($script:currentBatchSize -lt 50) { $script:currentBatchSize = 50 }
    LogVerbose "Retrying with smaller batch size of $($script:currentBatchSize)"
}

Function Throttled()
{
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning

    if ([String]::IsNullOrEmpty($script:Tracer.LastResponse))
    {
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds")
    {
        # We are throttled, and the server has told us how long to back off for
        Log "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
        Start-Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text"
        Log "Throttling budget should now be reset, resuming operations" Gray
        return $true
    }
    return $false
}

function ThrottledFolderBind()
{
    param (
        [Microsoft.Exchange.WebServices.Data.FolderId]$folderId,
        $propset = $null,
        $exchangeService = $null)

    if ($folderId -eq $null)
    {
        Log "[ThrottledFolderBind]Empty folder Id passed to ThrottledFolderBind" Red
        return $null
    }

    LogVerbose "[ThrottledFolderBind]Attempting to bind to folder $folderId"
    $folder = $null
    if ($exchangeService -eq $null)
    {
        Log "[ThrottledFolderBind]No ExchangeService object set" Red
        return $null
    }
    if ($propset -eq $null)
    {
        $propset = $script:requiredFolderProperties
    }

    try
    {

        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
        if (!($folder -eq $null))
        {
            LogVerbose "[ThrottledFolderBind]Successfully bound to folder $folderId"
        }
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
            if (!($folder -eq $null))
            {
                LogVerbose "[ThrottledFolderBind]Successfully bound to folder $folderId"
            }
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    ReportError "ThrottledFolderBind"
    LogVerbose "FAILED to bind to folder $folderId"
    return $null
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        $RootFolder,
        [string]$FolderPath,
        [bool]$Create = $false
    )	
	
    if ( $RootFolder -eq $null )
    {
        LogVerbose "[GetFolder]GetFolder called with null root folder"
        return $null
    }

    LogVerbose "[GetFolder]Locating folder: $FolderPath"
    if ($FolderPath.ToLower().StartsWith("wellknownfoldername"))
    {
        # Well known folder, so bind to it directly
        $wkf = $FolderPath.SubString(20)
        if ( $wkf.Contains("\") )
        {
            $wkf = $wkf.Substring( 0, $wkf.IndexOf("\") )
        }
        LogVerbose "[GetFolder]Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $Mailbox )
        $RootFolder = ThrottledFolderBind $folderId $null $script:service
        if ($RootFolder -ne $null)
        {
            $FolderPath = $FolderPath.Substring(20+$wkf.Length)
            LogVerbose "[GetFolder]Remainder of path to match: $FolderPath"
        }
    }

	$Folder = $RootFolder
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
                $FolderResults = $Null
                try
                {
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                }
                catch {}
                if ($FolderResults -eq $Null)
                {
                    if (Throttled)
                    {
                        try
                        {
				            $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                        }
                        catch {}
                    }
                }
                if ($FolderResults -eq $null)
                {
                    return $null
                }

				if ($FolderResults.TotalCount -gt 1)
				{
					# We have more than one folder returned... We shouldn't ever get this, as it means we have duplicate folders
					$Folder = $null
					Log "[GetFolder]Duplicate folders ($($PathElements[$i])) found in path $FolderPath" Red
					break
				}
                elseif ( $FolderResults.TotalCount -eq 0 )
                {
                    if ($Create)
                    {
                        # Folder not found, so attempt to create it
					    $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($RootFolder.Service)
					    $subfolder.DisplayName = $PathElements[$i]
                        try
                        {
					        $subfolder.Save($Folder.Id)
                            LogVerbose "[GetFolder]Created folder $($PathElements[$i])"
                        }
                        catch
                        {
					        # Failed to create the subfolder
					        $Folder = $null
					        Log "[GetFolder]Failed to create folder $($PathElements[$i]) in path $FolderPath" Red
					        break
                        }
                        $Folder = $subfolder
                    }
                    else
                    {
					    # Folder doesn't exist
					    $Folder = $null
					    Log "[GetFolder]Folder $($PathElements[$i]) doesn't exist in path $FolderPath" Red
					    break
                    }
                }
                else
                {
				    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id $null $RootFolder.Service
                }
			}
		}
	}
	
	$Folder
}

function GetFolderPath($Folder)
{
    # Return the full path for the given folder

    # We cache our folder lookups for this script
    if (!$script:folderCache)
    {
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive
        # We use a .Net Dictionary object instead
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $parentFolder = ThrottledFolderBind $Folder.Id $script:requiredFolderProperties $Folder.Service
    $folderPath = $Folder.DisplayName
    $parentFolderId = $Folder.Id
    while ($parentFolder.ParentFolderId -ne $parentFolderId)
    {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
        {
            try
            {
                $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
            }
            catch {}
        }
        else
        {
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $script:requiredFolderProperties $Folder.Service
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function ParseDate($date, $description)
{
    # Parse the date, failing if it is specified but invalid

    if ([String]::IsNullOrEmpty($date)) { return $Null }

    try
    {
        $date = [DateTime]::Parse($date)
        return $date
    }
    catch
    {
        Log "Invalid -$description parameter: $date" Red
        exit
    }
}

function ProcessMailbox()
{
    # Process the mailbox
    Log "Processing mailbox $Mailbox" Gray
	$script:service = CreateService( $Mailbox )
	if ($script:service -eq $Null)
	{
		Log "Failed to create ExchangeService" Red
	}
	
    $Folder = $Null
    if ([String]::IsNullOrEmpty($FolderPath))
    {
        $FolderPath = "WellKnownFolderName.MsgFolderRoot"
    }

    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox )
    $rootFolder = ThrottledFolderBind $folderId $null $script:service

	$Folder = GetFolder $rootFolder $FolderPath
	if (!$Folder)
	{
		Log "Failed to find folder $FolderPath" Red
		return
	}

    # Ensure any date parameters are valid (and parse into DateTime if so)
    $script:createdBeforeDate = ParseDate $CreatedBefore "CreatedBefore"
    $script:createdAfterDate = ParseDate $CreatedAfter "CreatedAfter"

    $script:duplicateItems = @()
	SearchForDuplicates $Folder
    Log "$($script:duplicateItems.Count) duplicate items have been found" Green
    if ($ReturnDuplicateCount)
    {
        $script:duplicateItems.Count
    }

    if ($script:duplicateItems.Count -gt 0)
    {
        if ($DuplicatesTargetFolder)
        {
            # We have a target folder specified, so we move all the duplicate items into that
            $script:moveToFolder = GetFolder $rootFolder $DuplicatesTargetFolder
            if ($script:moveToFolder -ne $null)
            {
                BatchMoveDuplicates
            }
            else
            {
                Log "Failed to find target folder for duplicates: $DuplicatesTargetFolder" Red
            }
        }
        else
        {
            BatchDeleteDuplicates
        }
    }
}


# The following is the main script


if ( [string]::IsNullOrEmpty($Mailbox) )
{
    $Mailbox = CurrentUserPrimarySmtpAddress
    if ( [string]::IsNullOrEmpty($Mailbox) )
    {
	    Write-Host "Mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red
	    Exit
    }
    else
    {
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $Mailbox)) -ForegroundColor Green
    }
}

# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
    TrustAllCerts
}
 
# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}
$script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PR_FOLDER_TYPE)


Write-Host ""

$script:currentBatchSize = 100

# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
	$csv = Import-CSV $Mailbox
	foreach ($entry in $csv)
	{
		$Mailbox = $entry.PrimarySmtpAddress
		if ( [string]::IsNullOrEmpty($Mailbox) -eq $False )
		{
			ProcessMailbox
		}
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox
}
