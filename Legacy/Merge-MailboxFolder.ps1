#
# Merge-MailboxFolder.ps1
#
# By David Barrett, Microsoft Ltd. 2015-2023. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

# TO-DO
# 1. Add ability to copy specific properties from source folder to target (e.g. retention).
# 2. Add Azure application registration functionality

param (
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the source mailbox (from which items will be moved/copied)")]
    [ValidateNotNullOrEmpty()]
    [string]$SourceMailbox,

    [Parameter(Position=1,Mandatory=$False,HelpMessage="Specifies the target mailbox (if not specified, the source mailbox is also the target)")]
    [ValidateNotNullOrEmpty()]
    [string]$TargetMailbox,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the source is Public Folders (not a mailbox)")]
    [ValidateNotNullOrEmpty()]
    [switch]$SourcePublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the target is Public Folders (not a mailbox)")]
    [ValidateNotNullOrEmpty()]
    [switch]$TargetPublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="If specified, an additional DeleteItem will be sent after successful MoveItem between public folders")]
    [ValidateNotNullOrEmpty()]
    [switch]$DeleteMovedPublicFolderItems,

    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be merged")]
    [ValidateNotNullOrEmpty()]
    $MergeFolderList,

    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be excluded")]
    [ValidateNotNullOrEmpty()]
    $ExcludeFolderList,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the folder paths are relative to the mailbox root.  If not specified, they are assumed to be relative to message folder root (i.e. Top of Information Store)")]
    [ValidateNotNullOrEmpty()]
    [switch]$PathsRelativeToMailboxRoot,

    [Parameter(Mandatory=$False,HelpMessage="A list of message classes to be excluded from processing")]
    $ExcludedMessageClasses,

    [Parameter(Mandatory=$False,HelpMessage="A list of message classes to be processed.  Any that don't match will be ignored.")]
    $IncludedMessageClasses,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given AQS filter will be moved `r`n(see https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-an-aqs-search-by-using-ews-in-exchange )")]
    [string]$SearchFilter,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were sent from the specified sender will be moved (useful for sorting).  Note that currently only one sender can be specified.")]
    [string]$OnlyItemsFromSender,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were created before the specified date will be processed (useful for archiving)")]
    [DateTime]$OnlyItemsCreatedBefore,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were sent or received before the specified date will be processed (useful for archiving).")]
    [DateTime]$OnlyItemsSentReceivedBefore,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were created before the specified date will be processed (useful for archiving)")]
    [DateTime]$OnlyItemsCreatedAfter,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were sent or received after the specified date will be processed (useful for archiving).")]
    [DateTime]$OnlyItemsSentReceivedAfter,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were modified before the specified date will be processed.")]
    [DateTime]$OnlyItemsModifiedBefore,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that were modified after the specified date will be processed.")]
    [DateTime]$OnlyItemsModifiedAfter,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EwsId (not path).")]
    [switch]$ByFolderId,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EntryId (not path).")]
    [switch]$ByEntryId,

    [Parameter(Mandatory=$False,HelpMessage="When specified, subfolders will also be processed.")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, all items in subfolders of source will be moved to specified target folder (hierarchy will NOT be maintained).")]
    [alias("MergeSubfolders")]
    [switch]$CombineSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, if the target folder doesn't exist, then it will be created (if possible).")]
    [switch]$CreateTargetFolder,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the source mailbox being accessed will be the archive mailbox.")]
    [switch]$SourceArchive,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the target mailbox being accessed will be the archive mailbox.")]
    [switch]$TargetArchive,

    [Parameter(Mandatory=$False,HelpMessage="When specified, hidden (associated) items of the folder are processed (normal items are ignored).")]
    [switch]$AssociatedItems,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the source folder will be deleted after the move (can't be used with -Copy).")]
    [switch]$Delete,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the source items.  Can only be used with -Copy, triggers effects a client-side move.")]
    [switch]$DeleteItems,

    [Parameter(Mandatory=$False,HelpMessage="When specified, items are copied rather than moved (can't be used with -Delete).")]
    [switch]$Copy,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the script outputs a count of total number of items that were affected (useful for automation).")]
    [switch]$ReturnTotalItemsAffected,

    [Parameter(Mandatory=$False,HelpMessage="If specified, no moves will be performed (but actions that would be taken will be logged).")]
    [switch]$WhatIf,

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

    [Parameter(Mandatory=$False,HelpMessage="Batch size (number of items batched into one EWS request) - this will be decreased if throttling is detected")]	
    [int]$BatchSize = 50
)
$script:ScriptVersion = "1.4.0"
$scriptStartTime = [DateTime]::Now

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
# EWS/OAuth library version: 1.0.4

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
    LogVerbose "Requesting token using certificate auth"
    
    try
    {
        $script:oauthToken = $acquire.ExecuteAsync().Result
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
    Log "Failed to obtain OAuth token (no error thrown)" Red
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
        $traceFileForCode = ""

        if (![String]::IsNullOrEmpty($TraceFile))
        {
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
                private string _traceFileFullPath = String.Empty;

			    public EWSTracer(string traceFileName = "$traceFileForCode" )
			    {
				    try
				    {
                        if (!String.IsNullOrEmpty(traceFileName))
					        _traceStream = File.AppendText(traceFileName);
                        FileInfo fi = new FileInfo(traceFileName);
                        _traceFileFullPath = fi.Directory.FullName + "\\" + fi.Name;
				    }
				    catch { }
                }

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

                public string TraceFileFullPath
                {
                    get { return _traceFileFullPath; }
                }
		    }
"@

        if ("EWSTracer" -as [type]) {} else {
            Add-Type -TypeDefinition $TraceListenerClass -ReferencedAssemblies $EWSManagedApiPath
        }
        $script:Tracer=[EWSTracer]::new($traceFileForCode)

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
                #[byte[]]$ByteArr = ([byte[]])$extendedProperty.Value;
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


Function DecreaseBatchSize()
{
    param (
        $DecreaseMultiplier = 0.8
    )

    if ($script:currentBatchSize -lt 1)
    {
        $script:currentBatchSize = 1
    }
    if ($script:currentBatchSize -eq 1)
        { return }

    if ($script:currentBatchSize -lt 6) { $script:currentBatchSize = 1 }
    elseif ($script:currentBatchSize -lt 11) { $script:currentBatchSize = 5 }
    elseif ($script:currentBatchSize -lt 21) { $script:currentBatchSize = 10 }
    elseif ($script:currentBatchSize -gt 50) { $script:currentBatchSize = [int]($script:currentBatchSize * $DecreaseMultiplier) }
    else { $script:currentBatchSize = [int]($script:currentBatchSize - 10) }

    if ($script:currentBatchSize -lt 1) { $script:currentBatchSize = 1 }
    LogVerbose "Reducing batch size to $($script:currentBatchSize)"
}



$script:itemRetryCount = @{}
Function RemoveProcessedItemsFromList()
{
    # Process the results of a batch move/copy and remove any items that were successfully moved from our list of items to move
    param (
        $requestedItems,
        $results,
        $suppressErrors = $false,
        $Items,
        $keepItemsIfNoResults = $false
    )

    if ($null -ne $results)
    {
        $failed = 0
        $permanentFailures = 0
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

                    # Check if this item is in the deleted list (and if so remove it)
                    if ($null -ne $script:deleteIds -and $script:deleteIds.Contains($requestedItems[$i])) { [void]$script:deleteIds.Remove($requestedItems[$i]) }                            
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
                        $permanentFailures++
                    }
                }
                else
                {
                    $retryErrors = @("ErrorBatchProcessingStopped", "ErrorTimeoutExpired")
                    if ( $retryErrors.Contains($results[$i].ErrorCode.ToString()) )
                    {
                        # This is a known error to retry, so we don't remove the item from the list
                        LogVerbose "Retriable batch error reported: $($results[$i].ErrorCode.ToString())"
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
                            if ($null -ne $script:deleteIds -and $script:deleteIds.Contains($requestedItems[$i])) { [void]$script:deleteIds.Remove($requestedItems[$i]) }                            
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
                                $permanentFailures++
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
        if ($keepItemsIfNoResults)
        {
            LogVerbose "No results returned, whole batch will be retried"
        }
        else
        {
            Log "No results returned - assuming all items processed" Yellow
            for ($i = 0; $i -lt $requestedItems.Count; $i++)
            {
                [void]$Items.Remove($requestedItems[$i])
            }
        }

    }
    if ( ($failed -gt 0) -and !$suppressErrors )
    {
        Log "$failed item(s) reported error during batch request ($permanentFailures not retriable)" Yellow
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
        $TargetFolderId,
        $Copy
    )

    if ($script:currentBatchSize -lt 1) { $script:currentBatchSize = $BatchSize }
    $consecutive401Errors = 0

	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    if ($Copy)
    {
        $progressActivity = "Copying items"
    }
    else
    {
        $progressActivity = "Moving items"
    }
    $totalItems = $ItemsToMove.Count
    $timeoutErrorCount = 0

    $percentComplete = -1
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete $percentComplete

    $script:deleteIds = [Activator]::CreateInstance($genericItemIdList) # This is used to check that items were deleted once moved (only happens when moving between public folders)
    while ( !$finished )
    {
	    $script:moveIds = [Activator]::CreateInstance($genericItemIdList)

        LogVerbose "Current batch size is $($script:currentBatchSize)"
        
        for ([int]$i=0; $i -lt $script:currentBatchSize; $i++)
        {
            if ($nunll -ne $ItemsToMove[$i])
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
                    if ($WhatIf)
                    {
                        Log "Would move/copy $($ItemsToMove[$i])"
                    }
                    else
                    {
                        LogVerbose "Added to move/copy batch: $($ItemsToMove[$i])"
                        if (!$Copy -and $script:publicFolders -and $DeleteMovedPublicFolderItems)
                        {
                            $script:deleteIds.Add($ItemsToMove[$i])
                            LogVerbose "Added to delete (due to public folder move): $($ItemsToMove[$i])"
                        }
                        elseif ($Copy -and $DeleteItems)
                        {
                            $script:deleteIds.Add($ItemsToMove[$i])
                            LogVerbose "Added to delete: $($ItemsToMove[$i])"
                        }
                    }
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
            if (!$WhatIf)
            {
                $stopWatch = [Diagnostics.Stopwatch]::StartNew()
                SetClientRequestId $script:sourceService
                if ( $Copy )
                {
                    LogVerbose "Sending batch request to copy $($moveIds.Count) items ($($ItemsToMove.Count) remaining)"
			        $results = $script:sourceService.CopyItems( $moveIds, $TargetFolderId, $false )
                }
                else
                {
                    LogVerbose "Sending batch request to move $($moveIds.Count) items ($($ItemsToMove.Count) remaining)"
			        $results = $script:sourceService.MoveItems( $moveIds, $TargetFolderId, $false)
                }
                $stopWatch.Stop()
                LogVerbose "Batch request completed in $($stopWatch.Elapsed)"
            }
        }
        catch
        {
            if ( Throttled )
            {
                
                # We need to resend the request as previous request would not have been processed
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
                    elseif ($Error[0].Exception.InnerException -and $Error[0].Exception.InnerException.ToString().Contains("The operation has timed out"))
                    {
                        # We've probably been throttled, so we'll reduce the batch size and try again
                        if ($script:currentBatchSize -gt 1)
                        {
                            LogVerbose "Timeout error received"
                            DecreaseBatchSize
                        }
                        else
                        {
                            # We are at minimum batch size (1) and still timing out.  If we get too many of these, we'll stop processing as it implies an issue that needs investigating.
                            # With a batch size of 1, we'd expect a throttling response rather than a timeout
                            $timeoutErrorCount++
                            if ($timeoutErrorCount -gt 3)
                            {
                                Log "Too many timeout errors received, halting processing" Red
                                $finished = $true
                            }
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

        if (!$WhatIf)
        {
            RemoveProcessedItemsFromList $moveIds $results $false $ItemsToMove !$Copy

            if ($script:deleteIds.Count -gt 0)
            {
                # Remove any items still in the Move list from the Delete list (as they haven't moved)
                if ($ItemsToMove.Count -gt 0)
                {
                    # Some items failed to move, so remove these from the list of those to be deleted
                    ForEach ($moveId in $ItemsToMove)
                    {
                        if ($script:deleteIds.Contains($moveId))
                        {
                            LogVerbose "Removing item from delete list: $moveId"
                            [void]$script:deleteIds.Remove($moveId)
                        }
                    }
                }
            }
        }
        else
        {
            for ($i = 0; $i -lt $moveIds.Count; $i++)
            {
                $ItemsToMove.Remove($moveIds[$i])
            }
        }

        $percentComplete = ( ($totalItems - $ItemsToMove.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$($percentComplete.ToString("0.#"))% complete" -PercentComplete $percentComplete

        if ($ItemsToMove.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Completed

    if ($script:deleteIds.Count -gt 0)
    {
        # We have a list of items to delete
        ThrottledBatchDelete $script:deleteIds -SuppressNotFoundErrors $true
    }
}

Function ThrottledBatchDelete()
{
    # Send request to delete items, allowing for throttling (which in this case is likely to manifest as time-out errors)
    param (
        $ItemsToDelete,
        $BatchSize = 100,
        $SuppressNotFoundErrors = $false
    )

    if ($script:MaxBatchSize -gt 0)
    {
        # If we've had to reduce the batch size previously, we'll start with the last size that was successful
        $BatchSize = $script:MaxBatchSize
    }

    $progressActivity = "Deleting items"
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $totalItems = $ItemsToDelete.Count
    $consecutive401Errors = 0
    $timeoutErrorCount = 0

    $percentComplete = -1
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete $percentComplete   

    while ( !$finished )
    {
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        
        for ([int]$i=0; $i -lt $BatchSize; $i++)
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
            SetClientRequestId $script:sourceService
			$results = $script:sourceService.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
            $consecutive401Errors = 0 # Reset the consecutive error count, as if we reach this point then this request succeeded with no error
        }
        catch
        {
            if (Throttled)
            {
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
                    elseif ($Error[0].Exception.InnerException -and $Error[0].Exception.InnerException.ToString().Contains("The operation has timed out"))
                    {
                        # We've probably been throttled, so we'll reduce the batch size and try again
                        if ($script:currentBatchSize -gt 1)
                        {
                            LogVerbose "Timeout error received"
                            DecreaseBatchSize
                        }
                        else
                        {
                            # We are at minimum batch size (1) and still timing out.  If we get too many of these, we'll stop processing as it implies an issue that needs investigating.
                            # With a batch size of 1, we'd expect a throttling response rather than a timeout
                            $timeoutErrorCount++
                            if ($timeoutErrorCount -gt 3)
                            {
                                Log "Too many timeout errors received, halting processing" Red
                                $finished = $true
                            }
                        }
                    }
                }

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
        }

        RemoveProcessedItemsFromList $deleteIds $results $SuppressNotFoundErrors $ItemsToDelete

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$($percentComplete.ToString("0.#"))% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
}

Function IsFolderExcluded()
{
    # Return $true if folder is in the excluded list

    param ($folder)

    $folderPath = GetFolderPath($folder)

    # Check PR_FOLDER_TYPE (0x36010003) to see if this is a search folder (FOLDER_SEARCH=2)
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
                        Log "Ignoring search folder: $folderPath"
                        return $true
                    }
                    LogVerbose "Folder is of type: $($prop.Value)"
                }
            }
            LogVerbose "Folder not identified as search folder"
        }
    }
    else
    {
        LogVerbose "No extended properties for folder, can't test for search folder"
    }

    if ($ExcludeFolderList)
    {
        LogVerbose "Checking for exclusions: $($ExcludeFolderList -join ',')"
        $rootFolderName = $script:sourceMailboxRoot.DisplayName.ToLower()
        ForEach ($excludedFolder in $ExcludeFolderList)
        {
            LogDebug "Comparing $($folderPath.ToLower()) to $($excludedFolder.ToLower())"
            if ($folderPath.ToLower().EndsWith($excludedFolder.ToLower()))
            {
                # This could be a match
                $pathsMatch = $true
                if ($folderPath.Length -gt $excludedFolder.Length)
                {
                    $pathPrefix = $folderPath.SubString(0, $folderPath.Length-$excludedFolder.Length).ToLower()
                    if ( ($pathPrefix -ne "\$rootFolderName") -and ($pathPrefix -ne "\$rootFolderName\"))  { $pathsMatch = $false }
                }
                if ($pathsMatch)
                {
                    Log "Excluded folder being skipped: $folderPath"
                    return $true
                }
            }
        }
    }
    return $false
}

Function MoveItems()
{
    param (
        $SourceFolderObject,
        $TargetFolderObject
    )	
	# Process all the items in the given source folder, and move (or copy) them to the target
	
    if ($null -eq $SourceFolderObject)
    {
        Log "Source folder is null, cannot move items" Red
        return
    }	
    if ($null -eq $TargetFolderObject)
    {
        Log "Target folder is null, cannot move items" Red
        return
    }	
	if ( $SourceFolderObject.Id -eq $TargetFolderObject.Id )
	{
		Log "Cannot move or copy from/to the same folder (source folder Id and target folder Id are the same)" Red
		return
	}
	
    if ($Copy)
    {
        $action = "Copy"
        $actioning = "Copying"
    }
    else
    {
        $action = "Move"
        $actioning = "Moving"
    }

    $folderSourceInfo = $SourceMailbox
    if ($SourcePublicFolders)
    {
        $folderSourceInfo = "Public Folders ($SourceMailbox)"
    }
    elseif ($SourceArchive)
    {
        $folderSourceInfo = "Archive ($SourceMailbox)"
    }
    $folderTargetInfo = $TargetMailbox
    if ($TargetPublicFolders)
    {
        $folderTargetInfo = "Public Folders ($TargetMailbox)"
    }
    elseif ($TargetArchive)
    {
        $folderTargetInfo = "Archive ($TargetMailbox)"
    }

	Log "$actioning from $($folderSourceInfo):$(GetFolderPath($SourceFolderObject)) to $($folderTargetInfo):$(GetFolderPath($TargetFolderObject))" White
	
    if ($SourcePublicFolders)
    {
        SetPublicFolderContentHeaders $script:sourceService $SourceFolderObject
    }    
    
	# Set parameters - we will process in batches of 500 for the FindItems call
	$Offset = 0
	$PageSize = 1000 # We're only querying Ids, so 1000 items at a time is reasonable
	$MoreItems = $true

    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time)
    $itemsToMove = New-Object System.Collections.ArrayList
	
    $progressActivity = "Reading items in folder $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
    LogVerbose "Building list of items to $($action.ToLower())"
    Write-Progress -Activity $progressActivity -Status "0 items read (out of $($SourceFolderObject.TotalCount))" -PercentComplete -1

    $itemSearchFilter = $null
    $searchFilters = @()
    if ($OnlyItemsCreatedBefore)
    {
        $searchFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $OnlyItemsCreatedBefore)
    }
    if ($OnlyItemsCreatedAfter)
    {
        $searchFilters +=  New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $OnlyItemsCreatedAfter)
    }
    if ($OnlyItemsModifiedBefore)
    {
        $searchFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $OnlyItemsModifiedBefore)
    }
    if ($OnlyItemsModifiedAfter)
    {
        $searchFilters +=  New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $OnlyItemsModifiedAfter)
    }
    if ($OnlyItemsSentReceivedBefore)
    {
        $searchFilters +=  New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent, $OnlyItemsSentReceivedBefore)
        $searchFilters +=  New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $OnlyItemsSentReceivedBefore)
    }
    if ($OnlyItemsSentReceivedAfter)
    {
        $searchFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent, $OnlyItemsSentReceivedAfter)
        $searchFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $OnlyItemsSentReceivedAfter)
    }

    if (![String]::IsNullOrEmpty($OnlyItemsFromSender))
    {
        $searchFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From, $OnlyItemsFromSender)
    }

    if ($searchFilters.Count -gt 0)
    {
        LogVerbose "Search filters applied: $($searchFilters.Count)"
        $itemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, $searchFilters)
    }
    elseif ($SearchFilter)
    {
        $itemSearchFilter = $SearchFilter
        LogVerbose "Search query being applied: $itemSearchFilter"
    }

	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent, [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
        if ($AssociatedItems)
        {
            $View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        }

        $FindResults = $null
        try
        {
            ApplyEWSOauthCredentials
            SetClientRequestId $script:sourceService
            if ($itemSearchFilter)
            {
                # We have a search filter, so need to apply this
                $FindResults = $SourceFolderObject.FindItems($itemSearchFilter, $View)
            }

            else
            {
                # No search filter, we want everything
		        $FindResults = $SourceFolderObject.FindItems($View)
            }
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
                $skip = $False
                if ($null -ne $IncludedMessageClasses)
                {
                    # Check if this is an included message class
                    $skip = $true
                    foreach ($includedMessageClass in $IncludedMessageClasses)
                    {
                        if ($item.ItemClass -like $includedMessageClass)
                        {
                            LogVerbose "Included message class $($item.ItemClass)"
                            $skip = $false
                            break
                        }
                    }
                }
                else
                {
                    if ($null -ne $ExcludedMessageClasses)
                    {
                        # Check if this is an excluded message class
                        foreach ($excludedMessageClass in $ExcludedMessageClasses)
                        {
                            if ($item.ItemClass -like $excludedMessageClass)
                            {
                                LogVerbose "Skipping item with message class $($item.ItemClass)"
                                $skip = $True
                                break
                            }
                        }
                    }
                }
                if (!$skip)
                {
                    [void]$itemsToMove.Add($item.Id)
                }
		    }
		    $MoreItems = $FindResults.MoreAvailable
            if ($MoreItems)
            {
                LogVerbose "$($itemsToMove.Count) items read so far (out of $($SourceFolderObject.TotalCount))"
            }
		    $Offset += $PageSize
        }
        $percentComplete = 100
        if ($SourceFolderObject.TotalCount -gt 0)
        {
            $percentComplete = [int](($itemsToMove.Count/$SourceFolderObject.TotalCount)*100)
        }
        Write-Progress -Activity $progressActivity -Status "$($itemsToMove.Count) items read (out of $($SourceFolderObject.TotalCount))" -PercentComplete $percentComplete
	}
    Write-Progress -Activity $progressActivity -Status "$($itemsToMove.Count) items read (out of $($SourceFolderObject.TotalCount))" -Completed

    if ( $itemsToMove.Count -gt 0 )
    {
        if ($Copy -and $Delete)
        {
            Log "$($itemsToMove.Count) items found; attempting to copy then delete" Green
        }
        else {
            Log "$($itemsToMove.Count) items found; attempting to $($action.ToLower())" Green
        }
        $script:totalItemsAffected += $itemsToMove.Count
        ThrottledBatchMove $itemsToMove $TargetFolderObject.Id $Copy

        # Add a check for the number of items left in the folder (we expect it to be zero)
        if ($SourcePublicFolders)
        {
            # Set the public folder headers back to heirarchy
            SetPublicFolderHeirarchyHeaders $script:sourceService $SourceMailbox
        }  
        $SourceFolderObject = ThrottledFolderBind $SourceFolderObject.Id $null $script:sourceService
        Log "$($folderSourceInfo):$(GetFolderPath($SourceFolderObject)) processed, now contains $($SourceFolderObject.TotalCount) items(s)" White
    }
    else
    {
        Log "No matching items were found" Green
        if ($SourcePublicFolders)
        {
            # Set the public folder headers back to heirarchy
            SetPublicFolderHeirarchyHeaders $script:sourceService $SourceMailbox
        } 
    }

	# Now process any subfolders
	if ($ProcessSubFolders)
	{
		if ($SourceFolderObject.ChildFolderCount -gt 0)
		{
            LogVerbose "Processing subfolders of $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
			$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
            $FolderView.PropertySet = $script:requiredFolderProperties
            SetClientRequestId $script:sourceService
			$SourceFindFolderResults = $SourceFolderObject.FindFolders($FolderView)
			ForEach ($SourceSubFolderObject in $SourceFindFolderResults.Folders)
			{
                if ( !(IsFolderExcluded($SourceSubFolderObject)) )
                {
                    if ($CombineSubfolders)
                    {
                        # We are moving all subfolder items into the target folder (ignoring hierarchy)
                        MoveItems $SourceSubFolderObject $TargetFolderObject
                    }
                    else
                    {
                        # We need to recreate folder hierarchy in target folder
				        $Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SourceSubFolderObject.DisplayName)
				        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(2)
                        $FolderView.PropertySet = $script:requiredFolderProperties
                        $FindFolderResults = $null
                        try
                        {
                            $attempts = 0
                            while ($null -eq $FindFolderResults -and $attempts -lt 3)
                            {
                                SetClientRequestId $script:targetService
	                            $FindFolderResults = $TargetFolderObject.FindFolders($Filter, $FolderView)
                                $attempts++
                                if ($null -eq $FindFolderResults)
                                {
                                    if (!Throttled)
                                    {
                                        $attempts = 10
                                    }

                                }
                            }
                        }
                        catch {}
                        if ($null -eq $FindFolderResults)
                        {
                            if ($WhatIf -and $CreateTargetFolder)
                            {
                                Log "Target folder not created due to -WhatIf: $($SourceSubFolderObject.DisplayName)"
                                $TargetFolderObject = New-Object PsObject
                                $TargetFolderObject | Add-Member NoteProperty DisplayName $SourceSubFolderObject.DisplayName
                            }
                            else
                            {
                                Log "FAILED TO LOCATE TARGET FOLDER: $($SourceSubFolderObject.DisplayName)" Red
                                $TargetSubFolderObject = $null
                            }
                        }
                        elseif ($FindFolderResults.TotalCount -eq 0)
				        {
                            LogVerbose "Creating target folder $($SourceSubFolderObject.DisplayName)"
                            if ( $SourceSubFolderObject.FolderClass -eq "IPF.Task" )
                            {
                                # Task folders need to be created as a TasksFolder, otherwise the EWS API returns an error (even though the folder creation succeeds)
					            $TargetSubFolderObject = New-Object Microsoft.Exchange.WebServices.Data.TasksFolder($script:targetService)
                            }
                            else
                            {
					            $TargetSubFolderObject = New-Object Microsoft.Exchange.WebServices.Data.Folder($script:targetService)
                                $TargetSubFolderObject.FolderClass = $SourceSubFolderObject.FolderClass
                            }
					        $TargetSubFolderObject.DisplayName = $SourceSubFolderObject.DisplayName
                            try
                            {
                                SetClientRequestId $script:targetService
					            $TargetSubFolderObject.Save($TargetFolderObject.Id)
                            }
                            catch
                            {
                                if (Throttled)
                                {
                                    try
                                    {
                                        SetClientRequestId $script:targetService
					                    $TargetSubFolderObject.Save($TargetFolderObject.Id)
                                    }
                                    catch
                                    {
                                        Log "FAILED TO CREATE TARGET FOLDER: $($SourceSubFolderObject.DisplayName)"
                                        $TargetSubFolderObject = $null
                                    }
                                }
                                else
                                {
                                    Log "FAILED TO CREATE TARGET FOLDER: $($SourceSubFolderObject.DisplayName)"
                                    $TargetSubFolderObject = $null
                                }
                            }
				        }
				        else
				        {
                            LogVerbose "Target folder already exists"
					        $TargetSubFolderObject = $FindFolderResults.Folders[0]
				        }
                        if ($null -ne $TargetSubFolderObject)
                        {
				            MoveItems $SourceSubFolderObject $TargetSubFolderObject
                            if ($SourcePublicFolders)
                            {
                                # Set the public folder headers back to heirarchy
                                SetPublicFolderHeirarchyHeaders $script:sourceService $SourceMailbox
                            }
                        }
                    }
                }
                else
                {
                    LogVerbose "Folder $(GetFolderPath($SourceSubFolderObject)) on excluded list"
                }
			}
		}
        else
        {
            LogVerbose "No subfolders found: $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
        }
	}

    # If delete parameter is set, check if the source folder is now empty (and if so, delete it)
    if ($Delete)
    {
        SetClientRequestId $script:sourceService
	    $SourceFolderObject.Load()
	    if (($SourceFolderObject.TotalCount -eq 0) -And ($SourceFolderObject.ChildFolderCount -eq 0))
	    {
		    # Folder is empty, so can be safely deleted
		    try
		    {
                SetClientRequestId $script:sourceService
			    $SourceFolderObject.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete)
			    Log "$($SourceFolderObject.DisplayName) successfully deleted" Green
		    }
		    catch
		    {
			    Log "Failed to delete $($SourceFolderObject.DisplayName): $($Error[0])" Red
		    }
	    }
	    else
	    {
		    # Folder is not empty
		    Log "$($SourceFolderObject.DisplayName) could not be deleted as it is not empty." Red
	    }
    }
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        $RootFolder,
        [string]$FolderPath,
        [bool]$Create,
        [string]$Mailbox
    )	
	
    if ($null -eq $RootFolder)
    {
        LogVerbose "Root folder not initialised"
        return $null
    }

    LogVerbose "Locating folder: $FolderPath"
    if ($FolderPath.ToLower().StartsWith("wellknownfoldername"))
    {
        # Well known folder, so bind to it directly
        $wkf = $FolderPath.SubString(20)
        if ( $wkf.Contains("\") )
        {
            $wkf = $wkf.Substring( 0, $wkf.IndexOf("\") )
        }
        LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $Mailbox )
        $RootFolder = ThrottledFolderBind $folderId $null $RootFolder.Service
        if ($null -ne $RootFolder)
        {
            $FolderPath = $FolderPath.Substring(20+$wkf.Length)
            LogVerbose "Remainder of path to match: $FolderPath"
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
                    SetClientRequestId $RootFolder.Service
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                }
                catch {}
                if ($null -eq $FolderResults)
                {
                    if (Throttled)
                    {
                        try
                        {
                            SetClientRequestId $RootFolder.Service
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
					Log "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" Red
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
                            LogVerbose "Created folder $($PathElements[$i])"
                        }
                        catch
                        {
					        # Failed to create the subfolder
					        $Folder = $null
					        Log "Failed to create folder $($PathElements[$i]) in path $FolderPath" Red
					        break
                        }
                        $Folder = $subfolder
                    }
                    else
                    {
					    # Folder doesn't exist
					    $Folder = $null
					    Log "Folder $($PathElements[$i]) doesn't exist in path $FolderPath" Red
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

        if ([String]::IsNullOrEmpty($TargetMailbox) -or ($TargetMailbox -eq $SourceMailbox))
        {
            # Retrieve Top of Information store for both primary and archive mailbox so that we can distinguish between them easily
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $sourceMbx )
            $primaryRoot = ThrottledFolderBind $folderId $script:requiredFolderProperties $Folder.Service
            $script:FolderCache.Add($primaryRoot.Id.UniqueId, $primaryRoot)

            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $sourceMbx )
            $archiveRoot = ThrottledFolderBind $folderId $script:requiredFolderProperties $Folder.Service
            if ($archiveRoot)
            {
                $archiveRoot.DisplayName = "ARCHIVE($($archiveRoot.DisplayName))"
                $script:FolderCache.Add($archiveRoot.Id.UniqueId, $archiveRoot)
            }
        }
    }

    if ($script:folderCache.ContainsKey($Folder.Id.UniqueId))
    {
        $parentFolder = $script:folderCache[$Folder.Id.UniqueId]
    }
    else
    {
        $parentFolder = ThrottledFolderBind $Folder.Id $script:requiredFolderProperties $Folder.Service
    }

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

function ConvertId($entryId)
{
    # Use EWS ConvertId function to convert from EntryId to EWS Id

    $id = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
    $id.Mailbox = $SourceMailbox
    $id.UniqueId = $entryId
    $id.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EntryId
    $ewsId = $Null
    try
    {
        SetClientRequestId $script:sourceService
        $ewsId = $script:sourceService.ConvertId($id, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
    }
    catch {}
    LogVerbose "EWS Id: $($ewsId.UniqueId)"
    return $ewsId.UniqueId
}

function ProcessMailbox()
{
    # Process the mailbox

    $script:publicFolders = $false

    Write-Host "Processing mailbox $SourceMailbox" -ForegroundColor Gray

    if ( !([String]::IsNullOrEmpty($script:originalLogFile)) )
    {
        $script:LogFile = $script:originalLogFile.Replace("%mailbox%", $SourceMailbox)
    }

	$script:sourceService = CreateService($SourceMailbox)
	if ($null -eq $script:sourceService)
	{
		Write-Host "Failed to connect to source mailbox" -ForegroundColor Red
        return
	}
    if ($SourcePublicFolders)
    {
        SetPublicFolderHeirarchyHeaders $script:sourceService $SourceMailbox
    }
    
    # Bind to source mailbox root folder
    $sourceMbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $SourceMailbox )
    if ($SourceArchive)
    {
        if ($PathsRelativeToMailboxRoot)
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot, $sourceMbx )
        }
        else
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $sourceMbx )
        }
    }
    elseif ($SourcePublicFolders)
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
        $script:publicFolders = $true
    }
    else
    {
        if ($PathsRelativeToMailboxRoot)
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $sourceMbx )
        }
        else
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $sourceMbx )
        }
    }
    $script:sourceMailboxRoot = ThrottledFolderBind $folderId $null $script:sourceService

    if ($null -eq $script:sourceMailboxRoot)
    {
        Write-Host "Failed to open source message store ($SourceMailbox)" -ForegroundColor Red
        if ($Impersonate)
        {
            Write-Host "Please check that you have impersonation permissions" -ForegroundColor Red
        }
        return
    }

    # Bind to target mailbox root folder
    if ([String]::IsNullOrEmpty($TargetMailbox))
    {
        $TargetMailbox = $SourceMailbox
    }
    if ($TargetMailbox -ne $SourceMailbox)
    {
        # We impersonate the source mailbox when accessing the target mailbox
        $script:targetService = CreateService $TargetMailbox $SourceMailbox
    }
    else
    {
        $script:targetService = $script:sourceService
    }
    $targetMbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $TargetMailbox )
    if ($TargetArchive)
    {
        if ($PathsRelativeToMailboxRoot)
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot, $targetMbx )
        }
        else
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $targetMbx )
        }
    }
    elseif ($TargetPublicFolders)
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot )
        $script:publicFolders = $true
    }
    else
    {
        if ($PathsRelativeToMailboxRoot)
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $targetMbx )
        }
        else
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $targetMbx )
        }
    }

    $script:targetMailboxRoot = ThrottledFolderBind $folderId $null $script:targetService
    if ($null -eq $script:targetMailboxRoot)
    {
        Write-Host "Failed to open target message store ($TargetMailbox)" -ForegroundColor Red
        return
    }

    if ($null -eq $MergeFolderList)
    {
        # No folder list, this is a request to move the entire mailbox

        MoveItems $script:sourceMailboxRoot $script:targetMailboxRoot
        return
    }


    $MergeFolderList.GetEnumerator() | ForEach-Object {
        $PrimaryFolder = $_.Name
        LogVerbose "Target folder is $PrimaryFolder"
        $SecondaryFolderList = $_.Value
        LogVerbose "Source folder list is $SecondaryFolderList"

        if ( $TargetArchive -and $PrimaryFolder.ToLower().Contains("wellknownfoldername") )
        {
            # Sanity check to ensure that we are not referencing incorrect WKF (e.g. cannot use WKF.Calendar to access calendar of an archive mailbox, as it will return the main mailbox calendar)
            if ( !$PrimaryFolder.ToLower().Contains("archive") )
            {
                Log "Invalid folder reference: cannot target primary WellKnownFolderName folders in an Archive mailbox: $PrimaryFolder" Red
                $PrimaryFolder = $null
            }
        }

        # Check we can bind to the target folder (if not, stop now)
        if ( ![String]::IsNullOrEmpty($PrimaryFolder) )
        {
            $TargetFolderObject = $null
            if ($ByFolderId)
            {
                $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($PrimaryFolder, $TargetMailbox)
                $TargetFolderObject = ThrottledFolderBind $id
            }
            elseif ($ByEntryId)
            {
                $PrimaryFolderId = ConvertId($PrimaryFolder)
                $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($PrimaryFolderId)
                $TargetFolderObject = ThrottledFolderBind $id
            }
            else
            {
	            $TargetFolderObject = GetFolder $script:targetMailboxRoot $PrimaryFolder ($CreateTargetFolder -and -not $WhatIf) $TargetMailbox
                if ($null -eq $TargetFolderObject -and $WhatIf)
                {
                    Log "Folder not created due to -WhatIf: $PrimaryFolder"
                    $TargetFolderObject = New-Object PsObject
                    $TargetFolderObject | Add-Member NoteProperty DisplayName $PrimaryFolder
                }
            }
        }

        if ($TargetFolderObject)
        {
	        # We have the target folder, now check we can get the source folder(s)
	        LogVerbose "Target folder located: $($TargetFolderObject.DisplayName)"

            # Source folder could be a list of folders...
            $SecondaryFolderList | ForEach-Object {
                $SecondaryFolder = $_
                LogVerbose "Secondary folder is $SecondaryFolder"

                if ( $SourceArchive -and $SecondaryFolder.ToLower().Contains("wellknownfoldername") )
                {
                    # Sanity check to ensure that we are not referencing incorrect WKF (e.g. cannot use WKF.Calendar to access calendar of an archive mailbox, as it will return the main mailbox calendar)
                    if ( !$SecondaryFolder.ToLower().Contains("archive") )
                    {
                        Log "Invalid folder reference: cannot target primary WellKnownFolderName folders in an Archive mailbox: $PrimaryFolder" Red
                        $SecondaryFolder = $null
                    }
                }

                if ( ![String]::IsNullOrEmpty($SecondaryFolder) )
                {
                    $SourceFolderObject = $null
                    if ($ByFolderId)
                    {
                        $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($SecondaryFolder)
                        $SourceFolderObject = ThrottledFolderBind $id
                    }
                    elseif ($ByEntryId)
                    {
                        $SecondaryFolderId = ConvertId($SecondaryFolder)
                        $id = New-Object Microsoft.Exchange.WebServices.Data.FolderId($SecondaryFolderId)
                        $SourceFolderObject = ThrottledFolderBind $id
                    }
                    else
                    {
	                    $SourceFolderObject = GetFolder $script:sourceMailboxRoot $SecondaryFolder $false $SourceMailbox
                    }
	                if ($SourceFolderObject)
	                {
		                # Found source folder, now initiate move
		                LogVerbose "Source folder located: $($SourceFolderObject.DisplayName)"
		                MoveItems $SourceFolderObject $TargetFolderObject
                        if ($SourcePublicFolders)
                        {
                            # Set the public folder headers back to heirarchy
                            SetPublicFolderHeirarchyHeaders $script:sourceService $SourceMailbox
                        }                        
	                }
                    else
                    {
                        Write-Host "Merge parameters invalid: merge $SecondaryFolder into $PrimaryFolder" -ForegroundColor Red
                    }
                }
            }
        }
        else
        {
            Write-Host "Merge parameters invalid: merge $SecondaryFolder into $PrimaryFolder" -ForegroundColor Red
        }
    }
    Write-Host "Finished processing mailbox $SourceMailbox" -ForegroundColor Gray
}


# The following is the main script

if ($LogFile.Contains("%mailbox%"))
{
    # We replace mailbox marker with the SMTP address of the mailbox - this gives us a log file per mailbox
    $script:originalLogFile = $LogFile
    $LogFile = $script:originalLogFile.Replace("%mailbox%", "Merge-MailboxFolder")
}
else
{
    $script:originalLogFile = ""
}

if ( [string]::IsNullOrEmpty($SourceMailbox) )
{
    $SourceMailbox = CurrentUserPrimarySmtpAddress
    if ( [string]::IsNullOrEmpty($SourceMailbox) )
    {
	    throw "Source mailbox not specified.  Failed to determine current user's SMTP address."
    }
    else
    {
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $SourceMailbox)) -ForegroundColor Green
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
	throw "Failed to locate EWS Managed API, cannot continue"
}
  
# Check whether parameters make sense
if ($DeleteItems -and $Copy)
{
    Log "Items successfully copied will be deleted"
}
elseif ($Delete -and $Copy)
{
    Log "Cannot use -Delete with -Copy (folders cannot be deleted as they will not be empty)"
    exit
}
elseif ($DeleteItems -and !$Copy)
{
    Log "Cannot use -DeleteItems without -Copy"
    exit
}

if ($null -eq $MergeFolderList)
{
    # No folder list, this is a request to move the entire mailbox
    # Check -ProcessSubfolders and -CreateTargetFolder is set, otherwise we fail now (can't move a mailbox without processing subfolders!)
    if (!$ProcessSubfolders)
    {
        throw "Mailbox merge requested, but subfolder processing not specified.  Please retry using -ProcessSubfolders switch."
    }
    if (!$CreateTargetFolder)
    {
        throw "Mailbox merge requested, but folder creation not allowed.  Please retry using -CreateTargetFolder switch."
    }
}

# Set up script variables.  We set them here so that we can modify them depending upon what we need (some parameters mean we need to pull more properties back, and we add these as necessary)
$script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PR_FOLDER_TYPE)
if ($SourcePublicFolders -or $TargetPublicFolders)
{
    $script:requiredFolderProperties.Add($script:PR_REPLICA_LIST)
}
$script:requiredItemProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::Subject)

if ($BatchSize -gt 10 -and ($SourcePublicFolders -or $TargetPublicFolders) )
{
    $BatchSize = 10
    Log "Batch size adjusted to 10 as public folders are being accessed"
}

Write-Host ""

$script:totalItemsAffected = 0

# Check whether we have a CSV file as input...
$FileExists = Test-Path $SourceMailbox
If ( $FileExists )
{
	# We have a CSV to process
    LogVerbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $SourceMailbox -Header "SourceMailbox,TargetMailbox"
	foreach ($entry in $csv)
	{
        LogVerbose $entry.PrimarySmtpAddress
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress))
        {
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress"))
            {
		        $SourceMailbox = $entry.PrimarySmtpAddress
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

if ($null -eq $script:Tracer)
{
    $script:Tracer.Close()
}


Log "Script finished in $([DateTime]::Now.SubTract($scriptStartTime).ToString())" Green
if ($script:logFileStreamWriter)
{
    $script:logFileStreamWriter.Close()
    $script:logFileStreamWriter.Dispose()
}

if ($ReturnTotalItemsAffected)
{
    $script:totalItemsAffected
}
