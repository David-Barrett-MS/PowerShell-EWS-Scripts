#
# Update-Folders.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2021. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

param (
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox).")]
    [switch]$Archive,
		
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, public folders will be processed.")]
    [switch]$PublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="If specified, folder or list of folders to process.  If omitted, root message folder is assumed.")]
    $FolderPaths = "",

    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be excluded.")]
    [ValidateNotNullOrEmpty()]
    $ExcludedFolderPaths,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is set, only those folders that specifically match the exclusion list will be excluded (subfolders will still be processed).")]
    [ValidateNotNullOrEmpty()]
    [switch]$DoNotExcludeSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is set, search folders will also be processed (by default they are excluded).")]
    [ValidateNotNullOrEmpty()]
    [switch]$IncludeSearchFolders,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the script attempts to create the folders specified in FolderPaths parameter if they do not already exist.")]
    [switch]$CreateFolderPath,

    [Parameter(Mandatory=$False,HelpMessage="Changes the display name of the folder(s).")]
    $NewDisplayName,
    
    [Parameter(Mandatory=$False,HelpMessage="Any characters defined here will be removed from any folder names. Use an array of characters, e.g. @('.').")]
    $RemoveCharactersFromDisplayName,
    
    [Parameter(Mandatory=$False,HelpMessage="Sets the class of the folder(s) to that specified (e.g. IPF.Note).")]
    $FolderClass,
    
    [Parameter(Mandatory=$False,HelpMessage="If specified, any folders that do not have an item class defined (i.e. it is empty) will have the item class set.  If -FolderClass is not specified, the class is set to the same as the parent folder.")]
    [switch]$RepairFolderClass,
    
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, subfolders will also be processed.")]
    [switch]$ProcessSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="Adds the given properties (must be supplied as hash table @{}) to the folder(s).")]
    $AddFolderProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the given properties from the folder(s).")]
    $DeleteFolderProperties,

    [Parameter(Mandatory=$False,HelpMessage="Only processes folders created after this date.")]
    [datetime]$CreatedAfter,
	
    [Parameter(Mandatory=$False,HelpMessage="Only processes folders created before this date.")]
    [datetime]$CreatedBefore,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only folders that have values in the given properties will be updated.")]
    $PropertiesMustExist,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only folders that match the given values in the given properties will be updated.  Properties must be supplied as a Dictionary @{""propId"" = ""value""}.")]
    $PropertiesMustMatch,
    
    [Parameter(Mandatory=$False,HelpMessage="Hides the folder(s) (PidTagAttributeHidden is set to true).")]
    [switch]$Hide,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the folder(s).")]
    [switch]$Delete,
    
    [Parameter(Mandatory=$False,HelpMessage="Purges (empties) the folder(s).  This parameter is required if you want to delete folders that have messages in them.")]
    [switch]$Purge,
    
    [Parameter(Mandatory=$False,HelpMessage="Purges (empties) the folder(s).  This switch will force a hard-delete of any items found in the folder (otherwise soft-delete is used).  Can only be used with -Purge (both switches are required).")]
    [switch]$HardPurge,
    
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

    [Parameter(Mandatory=$False,HelpMessage="User-Agent header that will be set on ExchangeService and AutodiscoverService objects.")]	
    [string]$UserAgent = "https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts",

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

    [Parameter(Mandatory=$False,HelpMessage="If specified, changes will be made to the mailbox (but actions that would be taken will be logged).")]	
    [switch]$WhatIf
)
$script:ScriptVersion = "1.0.8"

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
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Verbose $Details
    if ( !$VerboseLogFile -and !$DebugLogFile -and ($VerbosePreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
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

# These functions are common for all my EWS scripts and are injected as part of the build/publish process.  Changes should be made to EWSOAuth.ps1 code snippet, not the script being run.
# EWS/OAuth library version: 1.0.5

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
            if ($null -eq $dll)
            {
                Log "$dllName not found in current directory - searching Program Files folders" Yellow
	            $dll = Get-ChildItem -Recurse "C:\Program Files (x86)" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            if (!$dll)
	            {
		            $dll = Get-ChildItem -Recurse "C:\Program Files" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            }
            }
        }

        if ($null -eq $dll)
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
    if ($null -eq $acquire)
    {
        Log "Failed to create token acquisition object" Red
        exit
    }
    LogVerbose "Requesting token using certificate auth"
    
    try
    {
        $execCall = $acquire.ExecuteAsync()
        $script:oauthToken = $execCall.Result
    }
    catch
    {
        Log "Failed to obtain OAuth token: $Error" Red
        exit # Failed to obtain a token
    }

    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    if ($null -ne $script:oAuthAccessToken)
    {
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
        $script:Impersonate = $true
        return
    }

    # If we get here, we don't have a token so can't continue
    if ($null -ne $execCall.Exception)
    {
        $global:CertException = $execCall.Exception
        Log "Failed to obtain OAuth token: $($global:CertException.Message)" Red
        Log "Full exception available in `$CertException"
    }
    else {
        Log "Failed to obtain OAuth token (no error thrown)" Red
    }
    exit
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$OAuthRedirectUri&response_mode=query&prompt=select_account&scope=openid%20profile%20email%20offline_access%20https://outlook.office365.com/.default"
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
        LogVerbose "Using auth code: $authcode"
        Write-Host "Auth code acquired, attempting to obtain access token" -ForegroundColor Green
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

    if ($null -ne $script:oAuthToken)
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
    #$tokenHeaderObject = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    while ($tokenPayload.Length % 4) { $tokenPayload = "$tokenPayload=" }
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)
    $tokenObject = $tokenArray | ConvertFrom-Json
    return $tokenObject
}

function LogOAuthTokenInfo
{
    if ($null -eq $global:OAuthAccessToken)
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

    if ($null -ne $script:oauthToken)
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
    elseif ($null -ne $OAuthCertificate)
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
            if ($GlobalTokenStorage -and $null -eq $script:oauthToken)
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
    if ( $null -eq $script:services ) { return }

    
    if ($DebugTokenRenewal -gt 0 -and $script:oauthToken)
    {
        # When debugging tokens, we stop after on every other EWS call and wait for the token to expire
        if ($script:oAuthDebugStop)
        {
            # Wait until token expires (we do this after every call when debugging OAuth)
            # Access tokens can't be revoked, but a policy can be assigned to reduce lifetime to 10 minutes: https://learn.microsoft.com/en-us/graph/api/resources/tokenlifetimepolicy?view=graph-rest-1.0
            if ( $null -ne $OAuthCertificate)
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
    
    if ($null -ne $OAuthCertificate)
    {
        if ( [DateTime]::UtcNow -lt $script:oauthToken.ExpiresOn.UtcDateTime) { return }
    }
    elseif ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    LogVerbose("OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($null -eq $exchangeCredentials) { return }

    if ($null -ne $OAuthCertificate)
    {
        $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
        if ( [DateTime]::UtcNow -ge $tokenExpire)
        {
            Log "OAuth Token renewal failed (certificate auth)"
            exit # We no longer have access to the mailbox, so we stop here
        }
    }
    else
    {
        if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -lt [DateTime]::UtcNow )
        { 
            Log "OAuth Token renewal failed"
            exit # We no longer have access to the mailbox, so we stop here
        }
        $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
    }

    Log "OAuth token successfully renewed; new expiry: $tokenExpire"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = $exchangeCredentials
        }
        LogVerbose "Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
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
        # We can only set any EWS properties once the API is loaded

        $script:PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6698, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) # PR_REPLICA_LIST is required to perform AutoDiscover requests for public folders
        $script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    }

    return $ewsApiLoaded
}

Function CurrentUserPrimarySmtpAddress()
{
    # Attempt to retrieve the current user's primary SMTP address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $result = $searcher.FindOne()

    if ($null -ne $result)
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

Function SetClientRequestId($exchangeService)
{
    # Apply unique client-request-id to the EWS service object

    $exchangeService.ClientRequestId = (New-Guid).ToString()
}

Function CreateTraceListener($exchangeService)
{
    # Create trace listener to capture EWS conversation (useful for debugging)

    if ([String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        Log "Managed API path missing; unable to create tracer" Red
        Exit
    }

    if ($null -eq $script:Tracer)
    {
        $TraceListenerClass = @"
            using System;
            using System.IO;
            using Microsoft.Exchange.WebServices.Data;

            public class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
            {
                private StreamWriter _traceStream = null;
                private string _lastResponse = String.Empty;
                private string _traceFileFullPath = "No trace file configured";

                public EWSTracer(string traceFileName = "" )
                {
                    if (!String.IsNullOrEmpty(traceFileName))
                    {
                        try
                        {
                            _traceStream = File.AppendText(traceFileName);
                            FileInfo fi = new FileInfo(traceFileName);
                            _traceFileFullPath = fi.Directory.FullName + "\\" + fi.Name;
                        }
                        catch { }
                    }
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
                    else if ( traceType.Equals("EwsRequest") )
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

                public string TraceFileFullPath
                {
                    get { return _traceFileFullPath; }
                }
            }
"@

        if ("EWSTracer" -as [type]) {} else {
            Add-Type -TypeDefinition $TraceListenerClass -ReferencedAssemblies $EWSManagedApiPath
        }
        $script:Tracer=[EWSTracer]::new($TraceFile)

        # Attach the trace listener to the Exchange service
        $exchangeService.TraceListener = $script:Tracer
        if (![String]::IsNullOrEmpty($TraceFile))
        {
            Log "Tracing to: $($script:Tracer.TraceFileFullPath)"
        }
    }
}

function CreateService($smtpAddress, $impersonatedAddress = "")
{
    # Creates and returns an ExchangeService object to be used to access mailboxes

    # First of all check to see if we have a service object for this mailbox already
    if ($null -eq $script:services)
    {
        $script:services = @{}
    }
    if ($script:services.ContainsKey($smtpAddress))
    {
        return $script:services[$smtpAddress]
    }

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    $exchangeService.UserAgent = $UserAgent

    # Do we need to use OAuth?
    if ($Office365) { $OAuth = $true }
    if ($OAuth)
    {
        $exchangeService.Credentials = GetOAuthCredentials
        if ($null -eq $exchangeService.Credentials)
        {
            # OAuth failed
            return $null
        }
    }
    else
    {
        # Set credentials if specified, or use logged on user.
        if ($null -ne $Credentials)
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

    SetClientRequestId $exchangeService
    $exchangeService.ReturnClientRequestId = $true

    $script:services.Add($smtpAddress, $exchangeService)
    LogVerbose "Currently caching $($script:services.Count) ExchangeService objects" $true
    return $exchangeService
}

Function CreateAutoDiscoverService($exchangeService)
{
    $autodiscover = new-object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService(new-object Uri("https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"), ExchangeVersion.Exchange2016)
    $autodiscover.Credentials = $exchangeService.Credentials
    $autodiscover.TraceListener = $exchangeService.TraceListener
    $autodiscover.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $autodiscover.TraceEnabled = $true
    $autodiscover.UserAgent = $UserAgent
    return $autodiscover;
}

function SetArchiveMailboxHeaders($ExchangeService, $PrimarySmtpAddress)
{
    
}

function SetPublicFolderHeirarchyHeaders($ExchangeService, $AutodiscoverAddress)
{
    # Sets the X-PublicFolderMailbox and X-AnchorMailbox properties for a request to public folders

    if (!$OAuth -and !$Office365)
    {
        LogDebug "Public folder headers not implemented for on-premises Exchange"
        return $null        
    }

    if ($null -eq $script:publicFolderHeirarchyHeaders)
    {
        # We keep a cache of known folders to avoid unnecessary AutoDiscover
        $script:publicFolderHeirarchyHeaders = New-Object 'System.Collections.Generic.Dictionary[String,String[]]'
    }

    $xAnchorMailbox = ""
    $xPublicFolderMailbox = ""
    if ($script:publicFolderHeirarchyHeaders.ContainsKey($AutodiscoverAddress))
    {
        $xAnchorMailbox = $script:publicFolderHeirarchyHeaders[$AutodiscoverAddress][0]
        $xPublicFolderMailbox = $script:publicFolderHeirarchyHeaders[$AutodiscoverAddress][1]
    }    
    else {
        $autoDiscoverService = CreateAutoDiscoverService $ExchangeService
        [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[]]$userSettingsRequired = @([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer)
        $userSettings = $autoDiscoverService.GetUserSettings($AutodiscoverAddress, $userSettingsRequired)
    
        if ($null -eq $userSettings)
        {
            LogVerbose "Failed to obtain Autodiscover public folder settings"
            return
        }
    
        $xAnchorMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation]
        if ([String]::IsNullOrEmpty($xAnchorMailbox))
        {
            LogVerbose "PublicFolderInformation not present in Autodiscover response"
            return
        }
        LogVerbose "Public folder heirarchy X-AnchorMailbox set to $xAnchorMailbox"
    
        # Now we need to retrieve the X-PublicFolderMailbox value
        $userSettings = $autoDiscoverService.GetUserSettings($xAnchorMailbox, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer)
        $xPublicFolderMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer]
        if ([String]::IsNullOrEmpty($xPublicFolderMailbox))
        {
            LogVerbose "PublicFolderInformation not present in Autodiscover response fpr $xAnchorMailbox"
            return
        }
        LogVerbose "Public folder heirarchy X-PublicFolderMailbox set to $xPublicFolderMailbox"

        $script:publicFolderHeirarchyHeaders.Add($AutodiscoverAddress, @($xAnchorMailbox, $xPublicFolderMailbox))
    }

    if ($ExchangeService.HttpHeaders.ContainsKey("X-PublicFolderMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-PublicFolderMailbox") | out-null
    }
    if ($ExchangeService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-AnchorMailbox") | out-null
    }
    $ExchangeService.HttpHeaders.Add("X-PublicFolderMailbox", $xPublicFolderMailbox) | out-null
    $ExchangeService.HttpHeaders.Add("X-AnchorMailbox", $xAnchorMailbox) | out-null
}

Function SetPublicFolderContentHeaders($ExchangeService, $PublicFolder)
{
    # Sets the X-PublicFolderMailbox and X-AnchorMailbox properties for a content request to public folders

    if (!$OAuth -and !$Office365)
    {
        LogDebug "Public folder headers not implemented for on-premises Exchange"
        return $null        
    }

    if ($null -eq $script:publicFolderContentHeaders)
    {
        # We keep a cache of known folders to avoid unnecessary AutoDiscover
        $script:publicFolderContentHeaders = New-Object 'System.Collections.Generic.Dictionary[String,String]'
    }

    $xPublicFolderMailbox = ""
    if ($script:publicFolderContentHeaders.ContainsKey($PublicFolder.FolderId.UniqueId))
    {
        $xPublicFolderMailbox = $script:publicFolderContentHeaders[$PublicFolder.FolderId.UniqueId]
    }
    else
    {
        # We need to perform an AutoDiscover request to obtain the correct header value        
        $replicaGuid = ""
        foreach ($extendedProperty in $PublicFolder.ExtendedProperties)
        {
            if ($extendedProperty.PropertyDefinition -eq $script:PR_REPLICA_LIST)
            {
                $replicaGuid =[System.Text.Encoding]::ASCII.GetString($extendedProperty.Value, 0, 36)
                break
            }
        }
        if ([String]::IsNullOrEmpty($replicaGuid))
        {
            LogVerbose "Public folder PR_REPLICA_LIST not present"
            return
        }

        # Work out the AutoDiscover address from the replica GUID and domain
        if ($null -eq $Mailbox -and $null -ne $SourceMailbox)
        {
            $Mailbox = $SourceMailbox
        }
        $domainStart = $Mailbox.IndexOf("@")
        if ($domainStart -lt 0)
        {
            Log "Invalid mailbox: $Mailbox" Red
            return
        }
        $autoDiscoverAddress = "$replicaGuid$($Mailbox.Substring($domainStart))"
        LogVerbose "AutoDiscover address for $autoDiscoverAddress to access public folder $($PublicFolder.DisplayName)"
        
        $autoDiscoverService = CreateAutoDiscoverService $ExchangeService
        if ($null -eq $autoDiscoverService)
        {
            LogVerbose "Failed to create AutoDiscover service"
            return
        }

        [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[]]$userSettingsRequired = @([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress)
        $userSettings = $autoDiscoverService.GetUserSettings($autodiscoverAddress, $userSettingsRequired)

        if ($null -eq $userSettings)
        {
            LogVerbose "Failed to obtain user settings for $autoDiscoverAddress"
            return
        }

        $xPublicFolderMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress]
        if ([String]::IsNullOrEmpty($xPublicFolderMailbox))
        {
            LogVerbose "Failed to obtain AutoDiscoverSMTPAddress for $autoDiscoverAddress"
            return
        }

        $script:publicFolderContentHeaders.Add($PublicFolder.FolderId.UniqueId, $xPublicFolderMailbox)
        LogVerbose "Caching X-PublicFolderMailbox for folder $($PublicFolder.DisplayName): $xPublicFolderMailbox"
    }

    # Both X-PublicFolderMailbox and X-AnchorMailbox are required for public folder content requests, but they have the same value
    if ($ExchangeService.HttpHeaders.ContainsKey("X-PublicFolderMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-PublicFolderMailbox") | out-null
    }
    if ($ExchangeService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-AnchorMailbox") | out-null
    }
    $ExchangeService.HttpHeaders.Add("X-PublicFolderMailbox", $xPublicFolderMailbox) | out-null
    $ExchangeService.HttpHeaders.Add("X-AnchorMailbox", $xPublicFolderMailbox) | out-null
    LogVerbose "Set X-PublicFolderMailbox and X-AnchorMailbox to $xPublicFolderMailbox"
}

Function Throttled()
{
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning
    if ( $null -eq $script:Tracer -or [String]::IsNullOrEmpty($script:Tracer.LastResponse))
    {
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds")
    {
        # We are throttled, and the server has told us how long to back off for
        $backOffMilliseconds = [int]::Parse($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text")
        $resumeTime = [DateTime]::Now.AddMilliseconds($backOffMilliseconds)
        Log "Back off for $backOffMilliseconds milliseconds (will resume at $($resumeTime.ToLongTimeString()))" Yellow
        Start-Sleep -Milliseconds $backOffMilliseconds
        Log "Resuming operations" Gray
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

    if ($null -eq $folderId)
    {
        Log "FolderId missing" Red
        return $null
    }

    LogVerbose "Attempting to bind to folder $folderId"
    $folder = $null
    if ($null -eq $exchangeService)
    {
        $exchangeService = $script:service
        if ($null -eq $exchangeService)
        {
            Log "No ExchangeService object set" Red
            return $null
        }
    }

    if ($null -eq $script:requiredFolderProperties)
    {
        # If scripts require a custom property set, this variable should be set before calling this function.  If it's missing at this point, we just retrieve the most useful folder properties
        $script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PR_FOLDER_TYPE)
        $script:requiredFolderProperties.Add($script:PR_REPLICA_LIST) # Required for public folder Autodiscover
    }    
    if ($null -eq $propset)
    {
        $propset = $script:requiredFolderProperties
    }

    try
    {
        ApplyEWSOauthCredentials
        SetClientRequestId $exchangeService
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
        if (!($null -eq $folder))
        {
            LogVerbose "Successfully bound to folder $folderId"
        }
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
            ApplyEWSOauthCredentials
            SetClientRequestId $exchangeService
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
            if (!($null -eq $folder))
            {
                LogVerbose "Successfully bound to folder $folderId"
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

#>** EWS/OAUTH FUNCTIONS END **#

Function LogFolderAction($Folder, [String]$Details, [ConsoleColor]$Colour)
{
    $folderPath = GetFolderPath $Folder
    if ($WhatIf)
    {
        Log "[WHATIF]$($folderPath): $Details" $Colour
    }
    else
    {
        Log "$($folderPath): $Details" $Colour
    }
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
        # Back off for the time given by the server
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

    LogVerbose "Attempting to bind to folder $folderId"
    $folder = $null
    if ($null -eq $exchangeService)
    {
        $exchangeService = $script:service
    }
    if ($null -eq $propset)
    {
        $propset = $script:requiredFolderProperties
    }

    try
    {
        if ($null -eq $propset)
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
        }
        else
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
        }
        if (!($null -eq $folder))
        {
            LogVerbose "Successfully bound to folder $folderId"
        }
        Start-Sleep -Milliseconds $script:currentThrottlingDelay
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
            if ($null -eq $propset)
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
            }
            else
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
            }
            if (!($null -eq $folder))
            {
                LogVerbose "Successfully bound to folder $folderId"
            }
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    LogVerbose "FAILED to bind to folder $folderId"
    return $null
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        $RootFolder,
        [string]$FolderPath,
        [bool]$Create
    )	
	
    if ( ($null -eq $RootFolder) -and (!$FolderPath.ToLower().StartsWith("wellknownfoldername")) )
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
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $RootFolder = ThrottledFolderBind $folderId $null $RootFolder.Service
        if ($null -ne $RootFolder)
        {
            $FolderPath = $FolderPath.Substring(20+$wkf.Length)
            if ( ![String]::IsNullOrEmpty($FolderPath) )
            {
                LogVerbose "[GetFolder]Remainder of path to match: $FolderPath"
            }
            else
            {
                return $RootFolder
            }
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
                    Start-Sleep -Milliseconds $script:currentThrottlingDelay
                }
                catch {}
                if ($null -eq $FolderResults)
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
                if ($null -eq $FolderResults)
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
    LogVerbose "Retrieving folder path for $($Folder.DisplayName)"

    # We cache our folder lookups for this script
    if (!$script:folderCache)
    {
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive
        # We use a .Net Dictionary object instead
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)

    if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
    {
        $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
    }
    else
    {
        $parentFolder = ThrottledFolderBind $Folder.Id $propset
    }
    $folderPath = $Folder.DisplayName
    $parentFolderId = $Folder.Id
    while ($null -ne $parentFolder -and $parentFolder.ParentFolderId -ne $parentFolderId )
    {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
        {
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
        }
        else
        {
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset
            if ($null -eq $parentFolder)
            {
                $script:LastError = $Error[0]
                return $folderPath
            }
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

Function EWSPropertyType($MAPIPropertyType)
{
    # Return the EWS property type for the given MAPI Property value

    switch ([Convert]::ToInt32($MAPIPropertyType,16))
    {
        0x0    { return $Null }
        0x1    { return $Null }
        0x2    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short }
        0x1002 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ShortArray }
        0x3    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer }
        0x1003 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::IntegerArray }
        0x4    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Float }
        0x1004 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::FloatArray }
        0x5    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Double }
        0x1005 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::DoubleArray }
        0x6    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Currency }
        0x1006 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CurrencyArray }
        0x7    { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ApplicationTime }
        0x1007 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ApplicationTimeArray }
        0x0A   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Error }
        0x0B   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean }
        0x0D   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Object }
        0x100D { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ObjectArray }
        0x14   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long }
        0x1014 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::LongArray }
        0x1E   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String }
        0x101E { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray }
        0x1F   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String }
        0x101F { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray }
        0x40   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime }
        0x1040 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTimeArray }
        0x48   { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CLSID }
        0x1048 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CLSIDArray }
        0x102  { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary }
        0x1102 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::BinaryArray }
    }
    Write-Verbose "Couldn't match MAPI property type"
    return $Null
}

Function GenerateEWSProp($PropertyDefinition)
{
    # Parse the string representation of the property and return as ExtendedPropertyDefinition
    $propdef = $null

    if ($PropertyDefinition.Contains("/"))
    {
        # Property definition will be one of these:
        # {GUID}/name/mapiType - named property
        # {GUID]/id/mapiType   - MAPI property (shouldn't be used when accessing named properties)
        # DefaultExtendedPropertySet/name/mapiType

        $propElements = $PropDef -Split "/"
        if ($propElements.Length -eq 2)
        {
            # We expect three elements, but if there are two it most likely means that the MAPI property Id includes the Mapi type
            if ($propElements[1].Length -eq 8)
            {
                $propElements += $propElements[1].Substring(4)
                $propElements[1] = [Convert]::ToInt32($propElements[1].Substring(0,4),16)
            }
        }

        if ( $propElements[0].StartsWith("{") )
        {
            # GUID based property definition
            try
            {
                $guid = New-Object Guid($propElements[0])
                $propType = EWSPropertyType($propElements[2])
                $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($guid, $propElements[1], $propType)
            }
            catch {}
        }
        else
        {
            # Test DefaultExtendedPropertySet definition
            try
            {
                $propSet = [Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::$($propElements[0])
                $propType = EWSPropertyType($propElements[2])
                $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propSet, $propElements[1], $propType)
            }
            catch {}
        }
    }
    else
    {
        # Assume MAPI property (e.g. 0x30070040, which is PR_CREATION_TIME)
        if ($PropertyDefinition.ToLower().StartsWith("0x"))
        {
            $PropertyDefinition = $PropertyDefinition.SubString(2)
        }
        try
        {
            $propId = [Convert]::ToInt32($PropertyDefinition.SubString(0,4),16)
            $propType = EWSPropertyType($PropertyDefinition.SubString(5))
            $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
        }
        catch {}
    }

    if ($null -ne $propdef)
    {
        LogVerbose "Property $PropertyDefinition successfully parsed"
    }
    else
    {
        Log "Failed to parse $PropertyDefinition"
    }
    return $propdef
}

Function GenerateEWSPropList($PropertyDefinitions)
{
    # Convert the given property definitions into EWS definitions
    $props = @()

    foreach ($PropDef in $PropertyDefinitions)
    {
        $EWSPropDef = GenerateEWSProp($PropDef)
        if ($null -ne $EWSPropDef)
        {
            $props += $EWSPropDef
        }
        else
        {
            Log "Failed to parse (or convert) property $PropDef" Red
        }
    }
    return $props
}

Function CreatePropLists()
{
    # Process each of the parameters that can contain properties and convert them to EWS property lists

    if ($DeleteItemProperties)
    {
        Write-Verbose "Building list of properties to delete"
        $script:deleteItemPropsEws = GenerateEWSPropList($DeleteItemProperties)
    }

    if ($PropertiesMustExist)
    {
        Write-Verbose "Building list of properties that must exist"
        $script:propertiesMustExistEws = GenerateEWSPropList($PropertiesMustExist)
    }

    if ($PropertiesMustMatch)
    {
        Write-Verbose "Building list of properties that must match"
        $script:propertiesMustMatchEws = @{}

        foreach ($PropDef in $PropertiesMustMatch.Keys)
        {
            $EWSPropDef = GenerateEWSProp($PropDef)
            if ($null -ne $EWSPropDef)
            {
                $script:propertiesMustMatchEws.Add($EWSPropDef, $PropertiesMustMatch[$PropDef])
                $script:requiredFolderProperties.Add($EWSPropDef)          
            }

        }
    }
}

function FolderHasRequiredProperties($folder)
{
    # Check that this folder matches any property requirements
    if ($null -eq $script:propertiesMustExistEws)
    {
        # No properties to test
        return $true
    }

    LogVerbose "Checking for required properties on folder $($folder.DisplayName)"
    foreach ($requiredProperty in $script:propertiesMustExistEws)
    {
        # Check the folder has this property
        $propExists = $false
        if (![String]::IsNullOrEmpty(($requiredProperty.PropertySetId)))
        {
            LogVerbose "Checking for $($requiredProperty.PropertySetId)"
            foreach ($itemProperty in $item.ExtendedProperties)
            {
                if ($requiredProperty.PropertySetId -eq $itemProperty.PropertyDefinition.PropertySetId)
                {
                    # Same property set, check the value
                    if (![String]::IsNullOrEmpty(($requiredProperty.Id)))
                    {
                        if ($requiredProperty.Id -eq $itemProperty.PropertyDefinition.Id)
                        {
                            $propExists = $true
                            break
                        }
                    }
                    elseif (![String]::IsNullOrEmpty(($requiredProperty.Name)))
                    {
                        if ($requiredProperty.Name -eq $itemProperty.PropertyDefinition.Name)
                        {
                            $propExists = $true
                            break
                        }
                    }
                }
            }
        }
        elseif ($null -ne $requiredProperty.Tag)
        {
            LogVerbose "Checking for $($requiredProperty.Tag)"
            foreach ($itemProperty in $item.ExtendedProperties)
            {
                if ($requiredProperty.Tag -eq $itemProperty.PropertyDefinition.Tag)
                {
                    $propExists = $true
                    break
                }
            }               
        }
        if (!$propExists)
        {
            Write-Verbose "$requiredProperty does not exist, ignoring item"
            return $false
        }
    }
    return $true
}

function FolderPropertiesMatchRequirements($folder)
{
    if ($null -ne $script:propertiesMustMatchEws)
    {
        foreach ($requiredProperty in $script:propertiesMustMatchEws.Keys)
        {
            # Check the item has this property
            $propMatches = $false

            foreach ($folderProperty in $folder.ExtendedProperties)
            {
                if (![String]::IsNullOrEmpty(($requiredProperty.PropertySetId)))
                {
                    if ($requiredProperty.PropertySetId -eq $folderProperty.PropertyDefinition.PropertySetId)
                    {
                        # Same property set, check the value
                        if (![String]::IsNullOrEmpty(($requiredProperty.Id)))
                        {
                            if ($requiredProperty.Id -eq $folderProperty.PropertyDefinition.Id)
                            {
                                Write-Host "Checking prop" -ForegroundColor Green
                                $propMatches = ($folderProperty.Value -eq $script:propertiesMustMatchEws[$requiredProperty])
                                break
                            }
                        }
                        elseif (![String]::IsNullOrEmpty(($requiredProperty.Name)))
                        {
                            if ($requiredProperty.Name -eq $folderProperty.PropertyDefinition.Name)
                            {
                                Write-Host "Checking named Prop" -ForegroundColor Green
                                $propMatches = ($folderProperty.Value -eq $script:propertiesMustMatchEws[$requiredProperty])
                                break
                            }
                        }
                    }
                }
                elseif ($null -ne $requiredProperty.Tag -and $requiredProperty.Tag -eq $folderProperty.PropertyDefinition.Tag)
                {
                    # Check MAPI extended property
                    LogVerbose "Checking MAPI Prop $($folderProperty.PropertyDefinition.Tag) ($($folderProperty.Value)) against $($script:propertiesMustMatchEws[$requiredProperty])"
                    $areEqual = Compare-Object $folderProperty.Value $script:propertiesMustMatchEws[$requiredProperty] -SyncWindow 0
                    $propMatches = $null -eq $areEqual
                    break
                }
            }

            if (!$propMatches)
            {
                Write-Verbose "$requiredProperty does not match, ignoring folder"
                return $false
            }
        }
    }
    return $true    
}

Function HexStringToByteArray {
    param(
        [String] $HexString
    )

    $Bytes = [byte[]]::new($HexString.Length / 2)

    For($i=0; $i -lt $HexString.Length; $i+=2){
        try {
            $Bytes[$i/2] = [convert]::ToByte($HexString.Substring($i, 2), 16)    
        }
        catch {
            return $null # If we fail to convert any single byte, then we didn't have a byte array
        }       
    }

    return $Bytes
}

Function AddFolderProperties($folder)
{
    # Add the specified properties to the folder

    # First of all ensure we have some properties to add...
    if ($null -eq $AddFolderProperties) { return }

    # We need to convert the properties to EWS extended properties
    if ($null -eq $script:addFolderPropsEws)
    {
        Write-Verbose "Building list of properties to add"
        $script:addFolderPropsEws = @{}
        foreach ($addProperty in $AddFolderProperties.Keys)
        {
            $value = $AddFolderProperties[$addProperty]
            if ($addProperty.ToLower().StartsWith("0x"))
            {
                $addProperty = $addProperty.SubString(2)
            }
            $propId = [Convert]::ToInt32($addProperty.SubString(0,4),16)
            $propType = EWSPropertyType($addProperty.SubString(5))

            $propdef = $Null
            try
            {
                $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
            }
            catch {}
            if ($null -ne $propdef)
            {
                if ($propType -eq [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
                {
                    [byte[]]$bytes = $null
                    if ($value.GetType().Name.Equals("Byte[]"))
                    {
                        $bytes = $value
                    }
                    else
                    {
                        # Parameter is binary but not byte array - test if it is a hex string
                        $bytes = $null
                        $bytes = HexStringToByteArray $value
                        if ($null -eq $bytes)
                        {
                            # Not a hex string, so try to convert it to a byte array
                            $bytes = [System.Convert]::FromBase64String($value)
                        }
                    }
                    $script:addFolderPropsEws.Add($propdef, $bytes)
                }
                else
                {
                    $script:addFolderPropsEws.Add($propdef, $value)
                }
                Write-Verbose "Added property $addProperty to add list; value is $($script:addFolderPropsEws[$propdef])"
            }
            else
            {
                Log "Failed to parse (or convert) property $addProperty" Red
            }
        }
    }

    # Now we add the properties to the folder
    foreach ($addProperty in $script:addFolderPropsEws.Keys)
    {
        $folder.SetExtendedProperty($addProperty, $script:addFolderPropsEws[$addProperty])
    }

    # Now update the folder
    try
    {
        if (!$WhatIf)
        {
            $folder.Update()
        }
        LogFolderAction $folder "added properties" Green
    }
    catch
    {
        LogFolderAction $folder "Error during update: $($Error[0])" Red
        $folder.Load() # We do this to clear the invalid update from the folder object
    }
}

Function DeleteFolderProperties($folder)
{
    # Delete any properties specified

    # Check if we are deleting any properties
    if ($null -eq $DeleteFolderProperties) { return }

    # We need to convert the properties to EWS extended properties
    if ($null -eq $script:deleteFolderPropsEws)
    {
        Write-Verbose "Building list of properties to delete"
        $script:deleteFolderPropsEws = @()
        foreach ($deleteProperty in $DeleteFolderProperties)
        {
            if ($deleteProperty.ToLower().StartsWith("0x"))
            {
                $deleteProperty = $deleteProperty.SubString(2)
            }
            $propId = [Convert]::ToInt32($deleteProperty.SubString(0,4),16)
            $propType = EWSPropertyType($deleteProperty.SubString(5))

            $propdef = $Null
            try
            {
                $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
            }
            catch {}
            if ($null -ne $propdef)
            {
                $script:deleteFolderPropsEws += $propdef
                Write-Verbose "Added property $deleteProperty to delete list"
            }
            else
            {
                Log "Failed to parse (or convert) property $deleteProperty" Red
            }
        }
    }

    $loadPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet($script:deleteFolderPropsEws)
    $loadPropSet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
    $folder.Load($loadPropSet)

    Write-Verbose "Deleting properties from $($folder.DisplayName)"

    # Process each of the properties and add delete request
    $foundProperties = @()
    foreach ($deleteProperty in $script:deleteFolderPropsEws)
    {
        $found = $False
        foreach ($extendedProperty in $folder.ExtendedProperties)
        {
            if ( ($deleteProperty.Tag -eq $extendedProperty.PropertyDefinition.Tag) -and ($deleteProperty.MapiType -eq $extendedProperty.PropertyDefinition.MapiType) )
            {
                $foundProperties += $deleteProperty
                $found = $True
            }
        }
        
        if (!$found)
        {
            $propTag = [String]::Format("0x{0}", [Convert]::ToString($deleteProperty.Tag, 16))
            LogFolderAction $folder "Property $propTag was not found" Red
        }
    }

    if ($foundProperties.Count -gt 0)
    {
        Write-Verbose "Applying changes"
        foreach ($foundProperty in $foundProperties)
        {
            $propTag = [String]::Format("0x{0}", [Convert]::ToString($foundProperty.Tag, 16))
            if ( $folder.RemoveExtendedProperty($foundProperty) )
            {
                LogFolderAction $folder "Property $propTag was found and added to delete request" Gray
            }
            else
            {
                LogFolderAction $folder "Property $propTag was found, but cannot be removed" Red
            }
        }

        # Now update the folder
        try
        {
            if (!$WhatIf)
            {
                $folder.Update()
            }
            LogFolderAction $folder "properties deleted" Green
        }
        catch
        {
            LogFolderAction $folder "Error while updating: $($Error[0])" Red
            $folder.Load($loadPropSet) # We do this to clear the invalid update from the folder object
        }
    }
    else
    {
        LogFolderAction $folder "No properties to delete, folder not updated" Green
    }
}

Function RemoveProcessedItemsFromList()
{
    # Process the results of a batch move/copy and remove any items that were successfully moved from our list of items to move
    param (
        $requestedItems,
        $results,
        $Items
    )

    if ($null -ne $results)
    {
        $failed = 0
        for ($i = 0; $i -lt $requestedItems.Count; $i++)
        {
            if ($results[$i].ErrorCode -eq "NoError")
            {
                $Items.Remove($requestedItems[$i])
            }
            else
            {
                if ( ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed") -or ($results[$i].ErrorCode -eq "ErrorInvalidOperation") -or ($results[$i].ErrorCode -eq "ErrorCannotDeleteObject") )
                {
                    # This is a permanent error, so we remove the item from the list
                    $Items.Remove($requestedItems[$i])
                }
                LogVerbose("Error $($results[$i].ErrorCode) reported for item: $($requestedItems[$i].UniqueId)")
                $failed++
            } 
        }
    }
    if ( $failed -gt 0 )
    {
        Log "$failed items reported error during batch request (if throttled, this is expected)" Yellow
    }
}

Function ThrottledBatchDelete()
{
    # Send request to delete items, allowing for throttling
    param (
        $ItemsToDelete
    )

    if ($script:currentBatchSize -lt 1) { $script:currentBatchSize = 500 }

	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $progressActivity = "Deleting items"
    $totalItems = $ItemsToDelete.Count
    LogVerbose "Batch deleting $totalItems items"
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    $deleteType = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
    if ($HardPurge)
    {
        LogVerbose "Setting delete type to HardDelete"
        $deleteType = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
    }

    while ( !$finished )
    {
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        
        for ([int]$i=0; $i -lt $script:currentBatchSize; $i++)
        {
            if ($null -ne $ItemsToDelete[$i])
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
            if (!$WhatIf)
            {
			    $results = $script:service.DeleteItems( $deleteIds, $deleteType, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly)
            }
            else
            {
                # As we are running in WhatIf mode, we remove all items from the delete list
                RemoveProcessedItemsFromList $deleteIds $deleteIds $ItemsToDelete
            }
        }
        catch
        {
            if ( Throttled )
            {
                # We've been throttled, nothing to do here as the check for throttling will also have backed off
            }
            elseif ($Error[0].Exception.InnerException.ToString().Contains("The operation has timed out"))
            {
                # We've probably been throttled, so we'll reduce the batch size and try again
                if ($script:currentBatchSize -gt 50)
                {
                    $script:currentBatchSize = 50
                    LogVerbose "Timeout error received"
                }
                else
                {
                    $finished = $true
                }
            }
            else
            {
                $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
                $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
                $responseXml = [xml]$lastResponse
	            if ($responseXml.Trace.Envelope.Body.Fault.detail.ResponseCode.Value -eq "ErrorNoRespondingCASInDestinationSite")
                {
                    # We get this error if the destination CAS (request was proxied) hasn't returned any data within the timeout (usually 60 seconds)
                    # Reducing the batch size should help here, and we want to reduce it quite aggressively
                    if ($script:currentBatchSize -gt 50)
                    {
                        LogVerbose "ErrorNoRespondingCASInDestinationSite error received"
                    }
                    else
                    {
                        $finished = $true
                    }
                }
                else
                {
                    Log "Unexpected error: $($Error[0].Exception.InnerException.ToString())" Red
                    $finished = $true
                }
            }
        }

        RemoveProcessedItemsFromList $deleteIds $results $ItemsToDelete
        LogVerbose "$($totalItems - $ItemsToDelete.Count) items deleted"

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
            Write-Progress -Activity $progressActivity -Status "100% complete" -Completed
        }
    }
}

Function PurgeFolder($Folder)
{
    LogVerbose "Purging $($Folder.DisplayName)"

    if (!($script:service.RequestedServerVersion.ToString().Contains("2007")) -and ($script:service.RequestedServerVersion -ne [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010))
    {
        # Schemas later than 2010 support folder empty, so we try that first
        try
        {
            if (!$WhatIf)
            {
                if ($HardPurge)
                {
                    $Folder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $False)
                }
                else
                {
                    $Folder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $False)
                }
                
                $Folder.Load()
            }
            LogFolderAction $Folder "emptied" Green
            return
        }
	    catch
	    {
            if (ErrorReported)
            {
		        LogFolderAction $Folder "unable to Empty(), attempting manual deletion" Yellow
            }
	    }
    }

    if ($WhatIf)
    {
        return
    }

    # Delete all items from the folder

	$Offset = 0
	$PageSize = 1000 # We're only querying Ids, so 1000 items at a time is reasonable
	$MoreItems = $true

    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time)
    $itemsToDelete = New-Object System.Collections.ArrayList
	
    $progressActivity = "Reading items in folder $(GetFolderPath($Folder))"
    LogVerbose "Building list of items to delete from $($Folder.DisplayName)"
    Write-Progress -Activity $progressActivity -Status "0 items found" -PercentComplete -1

	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)

        $FindResults = $null
        try
        {
		    $FindResults = $Folder.FindItems($View)
            Start-Sleep -Milliseconds $script:currentThrottlingDelay
        }
        catch
        {
            # We have an error, so check if we are being throttled
            if (Throttled)
            {
                $FindResults = $null # We do this to retry the request
            }
            else
            {
                Log "Error when querying items: $($Error[0])" Red
                $MoreItems = $false
            }
        }
		
        if ($FindResults)
        {
		    ForEach ($item in $FindResults.Items)
		    {
                [void]$itemsToDelete.Add($item.Id)
		    }
		    $MoreItems = $FindResults.MoreAvailable
            if ($MoreItems)
            {
                LogVerbose "$($itemsToDelete.Count) items found so far, more available"
            }
		    $Offset += $PageSize
        }
        Write-Progress -Activity $progressActivity -Status "$($itemsToDelete.Count) items found" -PercentComplete -1
	}
    Write-Progress -Activity $progressActivity -Status "$($itemsToDelete.Count) items found" -Completed

    if ( $itemsToDelete.Count -gt 0 )
    {
        Log "$($itemsToDelete.Count) item(s) found; attempting to delete" Green
        ThrottledBatchDelete $itemsToDelete
        LogFolderAction $Folder "manually emptied" Green
    }
    else
    {
        Log "No items found in $($Folder.DisplayName)"
    }
}

function RemoveCharactersFromDisplayName($Folder)
{
    # Check for any characters in the folder name that we need to remove

    if (!$RemoveCharactersFromDisplayName)
        { return }

    [string]$displayName = $Folder.DisplayName
    foreach ($char in $RemoveCharactersFromDisplayName)
    {
        if ($displayName.Contains($char))
        {
            $displayName = $displayName.Replace($char, '')
        }
    }
    if ($displayName.Equals($Folder.DisplayName))
    {
        # No change
        LogVerbose "No characters found to replace in DisplayName"
        return
    }
    LogFolderAction $Folder "DisplayName changed to $displayName"
    if (!$WhatIf)
    {
        $Folder.DisplayName = $displayName
        $Folder.Update()
    }
}

Function IsFolderExcluded()
{
    # Return $true if folder is in the excluded list

    param ($folder)

    if (!$ExcludedFolderPaths)
    {
        return $false
    }

    $folderPath = GetFolderPath($folder)

    # Check PR_FOLDER_TYPE (0x36010003) to see if this is a search folder (FOLDER_SEARCH=2)
    if (!$IncludeSearchFolders)
    {
        if ($folder.ExtendedProperties)
        {
            if ($folder.ExtendedProperties.Count -gt 0)
            {
                foreach ($prop in $folder.ExtendedProperties)
                {
                    if ($prop.PropertyDefinition -eq $script:PR_FOLDER_TYPE)
                    {
                        if ($prop.Value -eq 2)
                        {
                            Log "[IsFolderExcluded]Ignoring search folder: $folderPath" Gray
                            return $true
                        }
                        LogVerbose "[IsFolderExcluded]Folder is of type: $($prop.Value)"
                    }
                }
                LogVerbose "[IsFolderExcluded]Folder not identified as search folder"
            }
        }
        else
        {
            LogVerbose "[IsFolderExcluded]No extended properties for folder, can't test for search folder"
        }
    }

    # Check if this is a specifically excluded path
    if ($ExcludedFolderPaths -and ![String]::IsNullOrEmpty($folderPath) )
    {
        LogVerbose "Checking for exclusions: $($ExcludeFolderList -join ',')"
        ForEach ($excludedFolder in $ExcludedFolderPaths)
        {
            if ($folderPath.ToLower().Equals($excludedFolder.ToLower()))
            {
                Log "[IsFolderExcluded]Excluded folder being skipped: $folderPath" Gray
                return $true
            }
        }
    }


    return $false
}

Function FolderMatchesDateRequirement
{
    param ($folder)

    if (!$CreatedBefore -and -not $CreatedAfter)
    {
        return $true
    }

    # Check if we have creation date criteria
    $createdTime = $null
    if ($folder.ExtendedProperties.Count -gt 0)
    {
        foreach ($prop in $folder.ExtendedProperties)
        {
            if ($prop.PropertyDefinition -eq $script:PR_CREATION_TIME)
            {
                $createdTime = $prop.Value
                LogVerbose "[FolderMatchesDateRequirement]Folder created: $createdTime"
            }
        }
    }

    if ($null -eq $createdTime)
    {
        # If we can't read the creation time, we assume it doesn't match our criteria
        LogVerbose "[FolderMatchesDateRequirement]Unable to retrieve PR_CREATION_TIME: $($folder.DisplayName)"
        return $false
    }

    if ($CreatedAfter)
    {
        if ($createdTime -lt $CreatedAfter)
        {
            LogVerbose "[FolderMatchesDateRequirement]Folder does not match CreatedAfter requirement: $($folder.DisplayName)"
            return $false
        }
    }
    if ($CreatedBefore)
    {
        if ($createdTime -gt $CreatedBefore)
        {
            LogVerbose "[FolderMatchesDateRequirement]Folder does not match CreatedBefore requirement: $($folder.DisplayName)"
            return $false
        }
    }
    return $true
}

Function ProcessFolder()
{
	# Apply updates to the given folder

    param (
        $Folder,
        $defaultFolderClass
    )
	if ($null -eq $Folder)
	{
		throw "No folder specified"
	}
	
    if ( [String]::IsNullOrEmpty($defaultFolderClass) ) { $defaultFolderClass = "IPF.Note" }

	# Process any subfolders
    if ( !$(IsFolderExcluded $Folder) -or $DoNotExcludeSubfolders )
    {
	    if ($ProcessSubFolders)
	    {
		    if ($Folder.ChildFolderCount -gt 0)
		    {
                Log "Processing subfolders of: $($Folder.DisplayName)" Gray
                # We read the list of all folders first, so that we have the complete list before any processing
                $subfolders = @()
                $moreFolders = $True
			    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
                $FolderView.PropertySet = $script:requiredFolderProperties
                while ($moreFolders)
                {
			        $FindFoldersResults = $Folder.FindFolders($FolderView)
                    $subfolders += $FindFoldersResults.Folders
                    $moreFolders = $FindFoldersResults.MoreAvailable
                    $FolderView.Offset += 500
                }
                # Process the subfolders
                if ($subfolders.Count -gt 0)
                {
			        ForEach ($subFolder in $subfolders)
			        {
                        if ( !$(IsFolderExcluded $subFolder) )
                        {
				            ProcessFolder $subFolder $defaultFolderClass
                        }
			        }
                }
		    }
            else
            {
                LogVerbose "Folder has no subfolders to process: $($Folder.DisplayName)"
            }
	    }
    }
    if ( $(IsFolderExcluded $Folder) )
    {
        Log "Folder excluded: $($Folder.DisplayName)" Gray
        return
    }
    if ( !(FolderMatchesDateRequirement $Folder) )
    {
        Log "Folder doesn't match date requirement: $($Folder.DisplayName)" Gray
        return
    }
    if ( !(FolderHasRequiredProperties $Folder) )
    {
        Log "Folder doesn't have required properties: $($Folder.DisplayName)" Gray
        return
    }
    if ( !(FolderPropertiesMatchRequirements $Folder) )
    {
        Log "Folder properties don't match filter: $($Folder.DisplayName)" Gray
        return
    }
    

    Log "Processing folder: $($Folder.DisplayName)" Gray

    # Check for folder purge
    if ($Purge)
    {
        PurgeFolder $Folder
    }

    # If delete parameter is set, check if the source folder is empty (and if so, delete it)
    if ($Delete)
    {
	    if (($Folder.TotalCount -eq 0) -And ($Folder.ChildFolderCount -eq 0))
	    {
		    # Folder is empty, so can be safely deleted
		    try
		    {
                if (!$WhatIf)
                {
			        $Folder.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete)
                }
			    LogFolderAction $Folder "deleted" Green
		    }
		    catch
            {
                ReportError "deleting $($Folder.DisplayName)"
            }
	    }
	    else
	    {
		    # Folder is not empty
            if ($WhatIf)
            {
                LogFolderAction $Folder "deleted" Green
            }
            else
            {
		        LogFolderAction $Folder "could not be deleted as it is not empty." Red
            }
	    }
        return # If Delete is specified, any other parameter is irrelevant
    }

    DeleteFolderProperties($Folder)
    AddFolderProperties($Folder)

    if (![String]::IsNullOrEmpty($NewDisplayName))
    {
        # Update display name of the folder
        try
        {
            if (!$WhatIf)
            {
                $Folder.DisplayName = $NewDisplayName
                $folder.Update()
            }
            LogFolderAction $folder "Display name updated to $NewDisplayName" Green
        }
        catch {}
        ReportError "Updating display name"
    }

    RemoveCharactersFromDisplayName($Folder)

    if ( $RepairFolderClass )
    {
        if ( [String]::IsNullOrEmpty($Folder.FolderClass) )
        {
            # Empty folder class, so set it
            try
            {
                if (!$WhatIf)
                {
                    $Folder.FolderClass = $defaultFolderClass
                    $folder.Update()
                }
                LogFolderAction $Folder "Folder class updated to $defaultFolderClass" Green
            }
            catch {}
            ReportError "Updating folder class"
        }
        else
        {
            LogFolderAction $Folder "Folder class is $($Folder.FolderClass)" Green
        }
    }
    elseif (![String]::IsNullOrEmpty($FolderClass))
    {
        # Update default item class of the folder
        try
        {
            if (!$WhatIf)
            {
                $Folder.FolderClass = $FolderClass
                $folder.Update()
            }
            LogFolderAction $Folder "Folder class updated to $FolderClass" Green
        }
        catch {}
        ReportError "Updating folder class"
    }
}

function ValidateWellKnownFolderDefaultItemType()
{
    param (
        $FolderName,
        $FolderClass
    )

    $folder = $null
    $folder = GetFolder $null "WellKnownFolderName.$FolderName" $False
    if ( !($null -eq $folder) )
    {
        if ($folder.FolderClass -ne $FolderClass)
        {
            if ( [String]::IsNullOrEmpty($folder.FolderClass) )
            {
                # FolderClass is currently blank, so we should be able to set it
                try
                {
                    $folder.FolderClass = $FolderClass
                    $folder.Update()
                    LogFolderAction $folder "Missing folder class updated to $FolderClass" Green
                }
                catch {}
                ReportError "Updating folder class [$($folder.DisplayName)]"
            }
            else
            {
                # FolderClass is set, so we can't change it but will report that it is not the expected type
                LogFolderAction $folder "Unexpected FolderClass ($($folder.FolderClass); should be $FolderClass" Red
            }
        }
        else
        {
            LogFolderAction $folder "FolderClass is already $($folder.FolderClass)" Green
        }
    }
    else
    {
        Log "Failed to check folder type for WellKnownFolderName::$FolderName" Red
    }
}

function ValidateWellKnownFolderDefaultItemTypes()
{
    # Go through each of the well known folders and check the item class

    ValidateWellKnownFolderDefaultItemType "Calendar" "IPF.Appointment"
    ValidateWellKnownFolderDefaultItemType "Contacts" "IPF.Contact"
    ValidateWellKnownFolderDefaultItemType "Journal" "IPF.Journal"
    ValidateWellKnownFolderDefaultItemType "Notes" "IPF.StickyNote"
    ValidateWellKnownFolderDefaultItemType "Tasks" "IPF.Task"

    # All other folders have a default item class of IPF.Note, so we don't need to explicitly check them
}

function ProcessMailbox()
{
    # Process the mailbox
    Log([string]::Format("Accessing mailbox {0}", $Mailbox))  Gray
    $script:currentThrottlingDelay = 0
    $script:currentBatchSize = 500
	$script:service = CreateService($Mailbox)
	if ($null -eq $script:service)
	{
		Log "Failed to create ExchangeService, cannot continue" Red
        return
	}

    # Define the properties we need to retrieve
    $script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $script:PR_CREATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3007, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
    $script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount, $script:PR_FOLDER_TYPE, $script:PR_CREATION_TIME)
    CreatePropLists
	
	# Get our root folder
    $rootFolder = $Null
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    if ($PublicFolders)
    {
        Write-Host "Processing public folders" Gray
        $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
    }
    else
    {
        if ($Archive)
        {
            LogVerbose "Processing archive mailbox"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx )
            $rootFolder = ThrottledFolderBind $folderId
        }
        else
        {
            LogVerbose "Processing primary mailbox"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
            $rootFolder = ThrottledFolderBind $folderId
        }
    }

    # Check we could get the root folder (if not, there's nothing else we can do)
    if ($null -eq $rootFolder)
    {
        Log "Unable to bind to root folder.  No further processing possible" Red
        return
    }

    if ($RepairFolderClass)
    {
        # When validating the folder class of folders, we need to ensure that the well known folders are set correctly, otherwise the repair logic will set everything to IPF.Note
        ValidateWellKnownFolderDefaultItemTypes
    }


    # FolderPath could be an array...
    if ([String]::IsNullOrEmpty($FolderPaths))
    {
        ProcessFolder $rootFolder "IPF.Note"
    }
    else
    {
        foreach ($fPath in $FolderPaths)
        {
            # Now get our specific folder, if we have one
            $folder = $null
            if (![String]::IsNullOrEmpty(($fPath)))
            {
                if ($fPath.ToLower().StartsWith("wellknownfoldername"))
                {
                    # Well known folder, so bind to it directly
                    $wkf = $fPath.SubString(20)
                    if ( $wkf.Contains("\") )
                    {
                        $wkf = $wkf.Substring( 0, $wkf.IndexOf("\") )
                    }
                    LogVerbose "Attempting to bind to well known folder: $wkf"
                    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
                    $folder = ThrottledFolderBind $folderId #[Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
                    if ($null -ne $folder)
                    {
                        $fPath = $fPath.Substring(20+$wkf.Length)
                        if (![String]::IsNullOrEmpty($fPath))
                        {
                            LogVerbose "Remainder of path to match: $fPath"
                            $folder = GetFolder $folder $fPath $False
                        }
                    }
                }
                else
                {
                    $folder = GetFolder $folder $fPath $False
                }
            }

            # Now start processing the folder
            if ($null -ne $folder)
            {
                ProcessFolder $folder
            }
            else
            {
                Log "Unable to locate folder: $fPath" Red
            }
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

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($OAuth)
    {
        Write-Host "Please specify *either* -Credentials *or* -OAuth" Red
        Exit
    }
}


Write-Host ""

# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
    LogVerbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $Mailbox -Header "PrimarySmtpAddress"
	foreach ($entry in $csv)
	{
        LogVerbose $entry.PrimarySmtpAddress
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress))
        {
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress"))
            {
		        $Mailbox = $entry.PrimarySmtpAddress
			    ProcessMailbox
            }
        }
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox
}

if ($null -ne $script:Tracer)
{
    $script:Tracer.Close()
}


Log "Script finished in $([DateTime]::Now.SubTract($scriptStartTime).ToString())" Green
if ($script:logFileStreamWriter)
{
    $script:logFileStreamWriter.Close()
    $script:logFileStreamWriter.Dispose()
}
