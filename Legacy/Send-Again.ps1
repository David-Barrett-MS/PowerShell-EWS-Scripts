#
# Send-Again.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2022. Use at your own risk.  No warranties are given.
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
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$False,HelpMessage="Folder to search for NDRs - if omitted, the Inbox folder is assumed.")]
    [string]$FolderPath,

    [Parameter(Mandatory=$False,HelpMessage="If set, messages will be saved to this folder instead of sent from the mailbox.  You can specify multiple Pickup folders using an array, and a round robin process will be followed.")]
    $SaveToPickupFolder = $null,

    [Parameter(Mandatory=$False,HelpMessage="If set, any messages that can't be saved to Pickup folder will instead be saved to this folder (for debugging purposes).")]
    $FailPickupFolder = $null,

    [Parameter(Mandatory=$False,HelpMessage="If set, this return-path will be stamped on resent messages.")]
    [string]$ReturnPath = "",

    [Parameter(Mandatory=$False,HelpMessage="If set, we'll forward all messages directly to target server (based on MX or specified SMTP server list)")]
    [switch]$SendUsingSMTP,

    [Parameter(Mandatory=$False,HelpMessage="A list of SMTP servers for specific target email addresses (or domains).  Any listed here will be used in preference to MX.")]
    $SMTPServerList,

    [Parameter(Mandatory=$False,HelpMessage="If set, messages will be written directly into the recipients' mailbox(es).  Requires the authenticating account to have ApplicationImpersonation rights on those mailboxes.")]
    [switch]$WriteDirectlyToRecipientMailbox,

    [Parameter(Mandatory=$False,HelpMessage="Folder to move processed items into.")]
    [string]$MoveProcessedItemsToFolder = "",

    [Parameter(Mandatory=$False,HelpMessage="Folder to move failed items into (those we attempted to process but were unable to).")]
    [string]$MoveFailedItemsToFolder = "",

    [Parameter(Mandatory=$False,HelpMessage="Folder to move encrypted items into (we won't attempt to process them).")]
    [string]$MoveEncryptedItemsToFolder = "",

    [Parameter(Mandatory=$False,HelpMessage="If set, any items that are encrypted will have the encrypted content removed.")]
    [switch]$RemoveEncryptedAttachments,

    [Parameter(Mandatory=$False,HelpMessage="If an item is processed, but couldn't be moved, then the Id will be added to this file so that it can be ignored on future runs.")]	
    [string]$IgnoreIdsLog = "",

    [Parameter(Mandatory=$False,HelpMessage="If set, all items processed (or failed to process) will be logged to the ignore file (recommended if messages are not being moved once processed).")]	
    [switch]$AddAllItemsToIgnoreLog,

    [Parameter(Mandatory=$False,HelpMessage="Batch size for processing NDRs (the number of items queried from the Inbox at one time).")]
    [int]$BatchSize = -1,

    [Parameter(Mandatory=$False,HelpMessage="If specified, checks for the messageclass are done clientside so that no search is required on the server.")]
    [switch]$FilterNDRsClientside,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only this number of items will be processed (script will stop when this number is reached).")]
    [int]$MaxItemsToProcess = -1,

    [Parameter(Mandatory=$False,HelpMessage="If specified, any messages larger than this will be failed (without being sent).")]
    [int]$MaxMessageSize = -1,

    [Parameter(Mandatory=$False,HelpMessage="If specified, message will only be resent to the provided recipient(s).")]
    $OnlyResendTo,

    [Parameter(Mandatory=$False,HelpMessage="If specified, specified recipient(s) will be added to the message.")]
    $AddResendTo,

    [Parameter(Mandatory=$False,HelpMessage="If specified, any messages found that have a blank From: header will have this address applied as the sender.")]	
    [string]$DefaultFromAddress = "",

    [Parameter(Mandatory=$False,HelpMessage="If specified, message will only be resent if the recipient specified in OnlyResendTo parameter was an original recipient of the email.  If this isn't specified, then all messages will be resent.")]
    [switch]$ConfirmResendAddress,

    [Parameter(Mandatory=$False,HelpMessage="If original message not included as attachment, attempt to find it in Sent Items.")]
    [switch]$SearchSentItems,

    [Parameter(Mandatory=$False,HelpMessage="Output statistics to the specified CSV file.")]
    [string]$StatsCSV,

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

    [Parameter(Mandatory=$False,HelpMessage="If specified, no actions (e.g. sending on) will be performed (but actions that would be taken will be logged).")]	
    [switch]$WhatIf
)
$script:ScriptVersion = "1.2.7"

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

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$RootFolder = $null,
        [String]$FolderPath = "",
        [switch]$Create
    )
        	
    if ( $null -eq $RootFolder )
    {
        # If we don't have a root folder, we assume the root of the message store
        if ( $null -eq $script:msgFolderRoot)
        {
            LogVerbose "Attempting to locate root message folder"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox )
            $script:msgFolderRoot = ThrottledFolderBind $folderId $null $script:service
            if ($null -eq $script:msgFolderRoot)
            {
                Log "Failed to bind to message root folder" Red
                return $null
            }
            LogVerbose "Retrieved root message folder"
        }
        $RootFolder = $script:msgFolderRoot
    }

	$Folder = $RootFolder
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
                LogDebug "Finding folder $($PathElements[$i])"
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
                $FolderResults = $Null
                try
                {
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                    Start-Sleep -Milliseconds $script:throttlingDelay
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
                if ( $null -eq $FolderResults)
                {
                    return $null
                }

				if ($FolderResults.TotalCount -gt 1)
				{
					# We have more than one folder returned... We shouldn't ever get this, as it means we have duplicate folders
					$Folder = $null
					Log "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red
					break
				}
                elseif ( $FolderResults.TotalCount -eq 0 )
                {
                    if ($Create)
                    {
                        # Folder not found, so attempt to create it
					    $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($script:service)
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
				    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id $null $script:service
                }
			}
		}
	}
	
	return $Folder
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

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)

    if ($Folder -eq "\")
    {
        # Special handling for root folder
        if  ($script:folderCache.ContainsKey("\"))
        {
            return $script:folderCache["\"]
        }
        $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $rootFolder = ThrottledFolderBind $folderId $propset $script:service
        if ($rootFolder)
        {
            $folderPath = "\$($rootFolder.DisplayName)"
            $script:folderCache.Add("\", $folderPath)
            $script:FolderCache.Add($rootFolder.Id.UniqueId, $rootFolder)
            return $folderPath
        }
        return ""
    }
    else
    {
        $parentFolder = ThrottledFolderBind $Folder.Id $propset $script:service
        $folderPath = $Folder.DisplayName
        $parentFolderId = $Folder.Id
    }

    while ($parentFolder.ParentFolderId -ne $parentFolderId)
    {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId))
        {
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
        }
        else
        {
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset $script:service
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = "$($parentFolder.DisplayName)\$folderPath"
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function RecipientIsInCollection()
{
    param (
        [Microsoft.Exchange.WebServices.Data.EmailAddressCollection]$RecipientCollection,
        [String]$Recipient
    )

    $recipientLC = $Recipient.ToLower()
    LogVerbose "Looking for $recipientLC in message recipients"

    foreach ($rcpt in $RecipientCollection)
    {
        LogVerbose "Checking $($rcpt.Address.ToLower()) against $recipientLC"
        if ($rcpt.Address.ToLower() -eq $recipientLC)
        {
            return $True
        }
    }
    return $false
}

function ResendMessage()
{
    param (
        [Microsoft.Exchange.WebServices.Data.EmailMessage]$Message
    )

    if (![String]::IsNullOrEmpty($OnlyResendTo))
    {
        if ($ConfirmResendAddress)
        {
            # We need to check whether the resend address was originally a recipient of this message
            $recipientProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ToRecipients,
                [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::CcRecipients, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::BccRecipients, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
            $Message.Load($recipientProperties)
            $resendAddressFound = RecipientIsInCollection $Message.ToRecipients $OnlyResendTo
            if (!$resendAddressFound)
            {
                $resendAddressFound = RecipientIsInCollection $Message.CcRecipients $OnlyResendTo
            }
            if (!$resendAddressFound)
            {
                $resendAddressFound = RecipientIsInCollection $Message.BccRecipients $OnlyResendTo
            }
            if (!$resendAddressFound)
            {
                Log "Not resending message $($Message.Subject) as $OnlyResendTo was not originally a recipient of the message"
                $script:skippedItems++
                return
            }
        }
    }

    try
    {
        # Copy message for sending (we don't want to affect the original message)
        $newMessage = $Message.Copy([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts)
        $newMessage.Load()
        LogVerbose "Copied message to resend"

        if (![String]::IsNullOrEmpty($OnlyResendTo))
        {
            $newMessage.ToRecipients.Clear()
            LogVerbose "Cleared ToRecipients"
            $newMessage.CcRecipients.Clear()
            LogVerbose "Cleared CcRecipients"
            $newMessage.BccRecipients.Clear()
            LogVerbose "Cleared BccRecipients"
            [void]$newMessage.ToRecipients.Add($OnlyResendTo)
            LogVerbose "Added $OnlyResendTo to ToRecipients"
        }
        $newMessage.Send()
        Log "Message `"$($newMessage.Subject)`" has been resent"
        $script:processedItems
    }
    catch
    {
        Log "Error occurred on send: $Error[0]" Red
        $script:skippedItems++
    }
}

function FindAndResendMessage()
{
    param (
        [String]$MessageId,
        [String]$SenderAddress )

    if ($SenderAddress -match "\<(.+)\>")
    {
        $senderEmail = $matches[1]
        LogVerbose "Email sender: $senderEmail"
    }
    else
    {
        $script:errorItems++
        return
    }

    LogVerbose "Getting ExchangeService for $senderEmail"
    $senderService = CreateService $senderEmail
    if ($null -eq $senderService)
    {
        Log "Failed to open mailbox of sender: $senderEmail" Red
        $script:errorItems++
        return
    }

    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $senderEmail )
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $mbx )
    LogVerbose "Connecting to Sent Items folder of $senderEmail"
    $folder = ThrottledFolderBind $folderId $null $senderService
    if ($null -eq $folder)
    {
        Log "Failed to connect to SentItems folder of $senderEmail" Red
        $script:errorItems++
        return
    }

    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId, $MessageId)

	$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(2, 0, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
	$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)

    $FindResults = $null
    try
    {
        $FindResults = $folder.FindItems($SearchFilter, $View)
        Start-Sleep -Milliseconds $script:currentThrottlingDelay
    }
    catch
    {
        # We have an error, so check if we are being throttled
        if (Throttled)
        {
            $FindResults = $folder.FindItems($SearchFilter, $View)
        }
        else
        {
            Log "Error when searching for message (id $MessageId): $($Error[0])" Red
            $script:errorItems++
            return
        }
    }
		
    if ($FindResults.Count -eq 1)
    {
        LogVerbose "Found message with id $MessageId in mailbox $senderEmail"
        ResendMessage $FindResults.Items[0]
    }
    else
    {
        if ($FindResults.Count -gt 1)
        {
            Log "More than one message found with id $MessageId in the Sent Items of mailbox $senderEmail" Red
            $script:errorItems++

        }
        else
        {
            Log "Message id $MessageId does not exist in the Sent Items of mailbox $senderEmail" Yellow
            $script:errorItems++
        }
    }
}

function ExtractHeaderValue()
{
   param (
        [String]$headers,
        [String]$headerName )

    # Extract the value of the message header.

    # Check if we have the full MIME or just headers (we are only interested in the headers)
    $endOfHeaders = $headers.IndexOf("$([Environment]::NewLine)$([Environment]::NewLine)")
    if ($endOfHeaders -gt 0)
    {
        #$content = $MIME.SubString($endOfHeaders+2)
        $headers = $MIME.SubString(0,$endOfHeaders)
    }
    $headerLines = $headers -split "`r`n|`r|`n"

    LogVerbose "Analysing header block (contains $($headerLines.Count) lines)"
    $i=0
    do {
        #LogVerbose "$($headerLines[$i])"
        if ( $headerLines[$i].StartsWith("$($headerName): ", [System.StringComparison]::OrdinalIgnoreCase) )
        {
            # This is the header we want
            $foundHeader = ""
            do {
                if ( [String]::IsNullOrEmpty($foundHeader) )
                {
                    $foundHeader += $headerLines[$i]
                }
                else
                {
                    $foundHeader += $headerLines[$i].TrimStart()
                }
                $i++
            } while ( ($i -lt $headerLines.Count) -and ($headerLines[$i+1].StartsWith("`t") -or $headerLines[$i+1].StartsWith(" ")) )
            $foundHeader = $foundHeader.Substring($headerName.Length+2)
            LogVerbose "Found header: $headerName"
            return $foundHeader
        }
        $i++
    } until ($i -ge $headerLines.Count)

    return $null
}

function ReplaceMIMEHeader()
{
   param (
        [String]$MIME,
        [String]$HeaderName,
        [String]$HeaderValue
    )

    # First of all we extract the headers from the MIME parts, as we don't want to process all the MIME.  Then we process each header in turn while building the new header block.
    $endOfHeaders = $MIME.IndexOf("$([Environment]::NewLine)$([Environment]::NewLine)")
    if ($endOfHeaders -lt 0)
    {
        Log "Failed to extract MIME headers, cannot update $HeaderName header" Yellow
        return $MIME
    }
    $headers = $MIME.SubString(0,$endOfHeaders)
    $content = $MIME.SubString($endOfHeaders+2)
    $headerLines = $headers -split "`r`n|`r|`n"

    $updatedHeaders = New-Object System.Text.StringBuilder
    $headerFound = $false

    LogVerbose "Analysing header block"
    $i=0
    do {
        if ($headerLines[$i].StartsWith("$($HeaderName): ") )
        {
            # This is the header to replace
            $headerFound = $true
            if (![String]::IsNullOrEmpty($HeaderValue))
            {
                $updatedHeaders.AppendLine("$($HeaderName): $HeaderValue") | out-null
                LogVerbose "Found $HeaderName header, replaced: $($HeaderName): $HeaderValue"
            }
            else
            {
                LogVerbose "Removed $HeaderName header"
            }
            do {
                $i++
            } while ( $headerLines[$i+1].StartsWith("`t") -or $headerLines[$i+1].StartsWith(" ") )
        }
        else
        {
            if (![String]::IsNullOrEmpty($headerLines[$i]))
            {
                $updatedHeaders.AppendLine($headerLines[$i]) | out-null
            }
        }
        $i++
    } until ($i -ge $headerLines.Count)

    if (!$headerFound)
    {
        if (![String]::IsNullOrEmpty($HeaderValue))
        {
            $updatedHeaders.AppendLine("$($HeaderName): $HeaderValue") | out-null
            LogVerbose "Added header: $($HeaderName): $HeaderValue"
        }
    }
    LogVerbose "Header block analysis complete; $i lines processed"

    # Now we just need to put the new header and the content back together
    return "$($updatedHeaders.ToString())$content"
}

function SendUsingSMTP()
{
   param (
        [String]$MIME,
        $recipients,
        [String]$senderSMTP
    )

    # Send-MailMessage and SmtpClient don't support sending MIME directly, so this is more complex than I thought
    return $false

    # Resolve-DnsName contoso.com -Type MX

    foreach ($recipient in $recipients)
    {
        Send-MailMessage -To $recipient 
    }

    return $true
}

function StripEncryptedAttachmentsFromMime([String]$mime)
{
    [System.IO.StringReader]$reader = [System.IO.StringReader]::new($mime)
    [System.Text.StringBuilder]$updatedMIME = [System.Text.StringBuilder]::new()
    [System.Text.StringBuilder]$unwrittenData = [System.Text.StringBuilder]::new()

    $boundary = ""
    $startBoundaryFound = $false
    $encryptedAttachmentFoundInMIMEPart = $false

    do
    {
        $line = $reader.ReadLine()

        if ($null -ne $line)
        {
            if ( [String]::IsNullOrEmpty($boundary) )
            {
                if ($line.StartsWith("Content-Type:"))
                {
                    while ($reader.Peek().ToInt32($null).Equals(9))
                    {
                        $line = "$line`r`n$($reader.ReadLine())"
                    }
                    if ($line -match "boundary=`"(.*?)\`"")
                    {
                        $boundary = $Matches[1]
                        LogVerbose "Found MIME boundary: $boundary"
                    }
                }
                [void]$updatedMIME.AppendLine($line)
            }
            else
            {
                # We have the MIME boundary (we are only interested in the outer boundary), so we need to identify the MIME part that has the encrypted attachment
                # Once identified, we remove the whole MIME part (based on outer boundaries)

                #Content-Description: message.rpmsg
                #Content-Type: application/x-microsoft-rpmsg-message
                #Content-Disposition: attachment; filename="message.rpmsg"
                #Content-Transfer-Encoding: base64

                [void]$unwrittenData.AppendLine($line)
                if ($line.Equals("--$boundary"))
                {
                    LogVerbose "Found MIME boundary"
                    if (!$startBoundaryFound)
                    {
                        $startBoundaryFound = $true
                        [void]$updatedMIME.Append($unwrittenData.ToString())
                        $unwrittenData = [System.Text.StringBuilder]::new()
                    }
                    else
                    {
                        if ($encryptedAttachmentFoundInMIMEPart)
                        {
                            # Remove this MIME part, as it contains an encrypted attachment
                            LogVerbose "Encrypted MIME part (length $($unwrittenData.Length)) removed"
                            $unwrittenData = [System.Text.StringBuilder]::new()
                            $encryptedAttachmentFoundInMIMEPart = $false
                        }
                        else
                        {
                            # Keep this MIME part
                            LogVerbose "MIME part contained no encrypted data"
                            [void]$updatedMIME.Append($unwrittenData.ToString())
                            $unwrittenData = [System.Text.StringBuilder]::new()
                        }
                    }
                }
                if ($line.StartsWith("Content-Type: application/x-microsoft-rpmsg-message"))
                {
                    LogVerbose "Encrypted MIME part found: $line"
                    $encryptedAttachmentFoundInMIMEPart = $true
                }
            }
        }

    } while ($null -ne $line)

    if ($unwrittenData.Length -gt 0)
    {
        # We still have some data that we need to deal with
        if ($encryptedAttachmentFoundInMIMEPart)
        {
            LogVerbose "Encrypted MIME part (length $($unwrittenData.Length)) removed"
            [void]$updatedMIME.AppendLine("--$boundary--") # We need to add back the closing MIME boundary
        }
        else
        {
            [void]$updatedMIME.Append($unwrittenData.ToString())
        }
    }
    return $updatedMIME.ToString()
}

function ValidateFolderMoveParameter($TargetFolder, $folderObject)
{
    if ( ($null -ne $folderObject) -or [String]::IsNullOrEmpty(($TargetFolder)) )
    {
        return $folderObject
    }

    LogVerbose "Locating folder: $TargetFolder"
    $folderObject = GetFolder $null $TargetFolder -Create
    if ($null -eq $folderObject)
    {
        Log "Unable to find/create target folder specified in parameters" Red
        exit
    }
    Log "Folder located: $TargetFolder"
    return $folderObject
}

function ResendMessages()
{
    param ( [System.Collections.ArrayList]$NDRs )

    if (!$script:ignoreIds)
    {
        $script:ignoreIds = @()
        if (![String]::IsNullOrEmpty($IgnoreIdsLog))
        {
            # We need to read all the existing item Ids that we will ignore
            if ( $(Test-Path $IgnoreIdsLog -PathType Leaf) )
            {
                
                $ignoreIdList = Get-Content -Path $IgnoreIdsLog
                foreach ($id in $ignoreIdList)
                {
                    $script:ignoreIds += $id
                }
                Log "$($script:ignoreIds.Count) item(s) being ignored, read from:  $IgnoreIdsLog" Green
            }
            else
            {
                # Need to create empty Ignore list so that we can append to it
                Out-File -FilePath $IgnoreIdsLog
                if ( -not $(Test-Path $IgnoreIdsLog -PathType Leaf) )
                {
                    ReportError
                    Log "Failed to create ignore file:  $IgnoreIdsLog" Red
                    exit
                }
                else
                {
                    Log "No existing items to ignore, created ignore file:  $IgnoreIdsLog" Green
                }
            }
        }
    }

    $progressActivity = "Processing NDRs"
    LogVerbose "Processing $($NDRs.Count) items"
    $ndrNum = 1
    foreach ($NDR in $NDRs)
    {
        # We need to read the message body and extract the Message Id and Sender from the message headers
        # We use these to trace the original message and resend it

        Write-Progress -Activity $progressActivity -Status "$ndrNum of $($NDRs.Count).  $(PerfReport)" -PercentComplete -1
        $ndrNum++
        $ndrProcessFail = $false
        $ndrProcessed = $false
        $ndrEncrypted = $false

        if ($script:ignoreIds.Contains($NDR.Id.UniqueId))
        {
            # We have already processed this item, but were unable to move it
            LogVerbose "Ignoring item: $($NDR.Id.UniqueId)"
            $script:ignoredItems++
            continue
        }

        # Load the message body (we only need the text version)
        $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ToRecipients, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ParentFolderId, $PidTagBody)
        Log "Retrieving message Id: $($NDR.Id.UniqueId)" Gray
        $NDR.Load($propset)

        # An NDR should have the failed recipients as the recipients, so we read these
        $toHeader = ""
        $resendTo = @()

        if ($OnlyResendTo)
        {
            # We are forwarding all messages to the specified recipients, so add them here (we ignore recipients on the message already)
            foreach ($resendToAddress in $OnlyResendTo)
            {
                if (![String]::IsNullOrEmpty($resendToAddress))
                {
                    LogVerbose "Resending to $($resendToAddress) based on OnlyResendTo parameter"
                    $resendTo += $resendToAddress
                }
            }
            if ( $resendTo.Count -lt 1 )
            {
                # Parsing the OnlyResendTo recipients didn't return any recipients...
                Log "OnlyResendTo parameter invalid: $OnlyResendTo" Red
                exit
            }
        }
        else
        {
            # We need to read the recipients to send to from the message
            if ($NDR.ToRecipients.Count -gt 0)
            {
                foreach ($recipient in $NDR.ToRecipients)
                {
                    if ($null -ne $recipient.Address)
                    {
                        if ($recipient.Address.StartsWith("IMCEAINVALID"))
                        {
                            Log "Invalid To recipient: $($recipient.Address)" Red
                        }
                        else
                        {
                            $address = $recipient.Address
                            if ( $address.StartsWith("=SMTP:") ) { $address = $address.SubString(6) }
                            LogVerbose "Resending to $($address) based on message recipients"
                            $resendTo += $address
                        }
                    }
                }
            }
            else
            {
                # If we haven't got recipients, check the mail body for mailto: links
                if ([String]::IsNullOrEmpty($toHeader))
                {
                    $mailToMatches = [System.Text.RegularExpressions.Regex]::Matches($NDR.Body.Text,"mailto:(.+?)`">",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($mailToMatches.Count -gt 0)
                    {
                        foreach ($mailToMatch in $mailToMatches)
                        {
                            if ( ![String]::IsNullOrEmpty($mailToMatch.Groups[1].Value) )
                            {
                                LogVerbose "Resending to $($mailToMatch.Groups[1].Value) based on message content"
                                $resendTo += $mailToMatch.Groups[1].Value
                            }
                        }
                    }
                }
            }
            if ($AddResendTo)
            {
                # We have recipient(s) that we need to add
                foreach ($addResend in $AddResendTo)
                {
                    if (![String]::IsNullOrEmpty($addResend))
                    {
                        $resendTo += $addResend
                    }
                }
            }
        }

        if ( $resendTo.Count -lt 1 )
        {
            # We couldn't determine who to send this to, so we fail
            Log "Could not read failed recipients from ndr" Red
            $ndrProcessFail = $true
        }
        else
        {
            foreach ($recipient in $resendTo)
            {
                if ([String]::IsNullOrEmpty($toHeader))
                {
                    $toHeader = "<$recipient>"
                }
                else
                {
                    $toHeader = "$toheader, <$recipient>"
                }
            }
            LogVerbose "Updated To header value: $toHeader"

            if ($NDR.Attachments.Count -eq 1)
            {
                # Attachment is most likely the original message, so resend that
                LogVerbose "Original message attached to NDR"
                $Error.Clear()
                try
                {
                    $itemAttachment = $null
                    $itemAttachment = [Microsoft.Exchange.WebServices.Data.ItemAttachment]$NDR.Attachments[0]
                    if ($MaxMessageSize -gt 0)
                    {
                        if ($itemAttachment.Size -gt $MaxMessageSize)
                        {
                            Log "Item too large ($($itemAttachment.Size))" Red
                            $itemAttachment = $null
                            $ndrProcessFail = $true
                        }
                    }

                    if ($null -ne $itemAttachment)
                    {
                        $itemAttachment.Load([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
                    }

                    if ($null -ne $itemAttachment)
                    {
                        $MIME = $itemAttachment.Item.MimeContent.ToString()
                        if ( $itemAttachment.Item.Attachments.Count -eq 2 )
                        {
                            if ( $itemAttachment.Item.Attachments[0].Name.Equals($itemAttachment.Item.Attachments[1].Name) )
                            {
                                # Item has two attachments with the same name, which means it is most likely encrypted

                                if ($RemoveEncryptedAttachments)
                                {
                                    $clearMIME = StripEncryptedAttachmentsFromMime $MIME
                                    if ($clearMIME.Length -ne $MIME.Length)
                                    {
                                        LogVerbose "Encrypted item found, encrypted attachment removed. Original MIME length: $($MIME.Length)  Updated MIME length: $($clearMIME.Length)"
                                        $MIME = $clearMIME
                                        $ndrEncrypted = $true
                                    }
                                    else
                                    {
                                        Log "Item with two attachments is not encrypted, processing as normal message. Original MIME length: $($MIME.Length)  Updated MIME length: $($clearMIME.Length)" Yellow
                                    }
                                }
                                else
                                {
                                    LogVerbose "Encrypted item detected, will be ignored"
                                }
                            }
                        }
                        $script:totalBytesRetrieved += $MIME.Length

                        if ($WriteDirectlyToRecipientMailbox)
                        {
                            if ($resendTo.Length -gt 0)
                            {
                                foreach ($targetMailbox in $resendTo)
                                {
                                    $targetService = CreateService $targetMailbox -ForceImpersonation
                                    if ($null -ne $targetService)
                                    {
                                        try
                                        {
                                            LogVerbose "Writing message into mailbox: $targetMailbox"
                                            $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::new($targetService)
                                            $mail.MimeContent = $MIME
                                            $mail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
                                            Log "Message successfully saved to $targetMailbox" Green
                                        }
                                        catch
                                        {
                                            ReportError
                                        }
                                    }
                                }
                                $ndrProcessed = $true
                            }
                            else
                            {
                                Log "Cannot save directly to mailbox as recipients could not be read" Red
                            }
                        }
                        else
                        {
                            $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "CC" -HeaderValue ""
                            $saveToPickupMime = $MIME # Workaround weird PowerShell issue where content of $MIME changes between here and save to pickup code.  Probably scope, but haven't worked it out yet...
                            if ( ![String]::IsNullOrEmpty($ReturnPath) )
                            {
                                $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "Return-Path" -HeaderValue $ReturnPath
                            }

                            if (![String]::IsNullOrEmpty($toHeader))
                            {
                                $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "To" -HeaderValue $toHeader
                            }
                            if (!$WhatIf)
                            {
                                # We don't resend the message if we are only collecting statistics
                                if ($SendUsingSMTP)
                                {
                                    LogVerbose "Resending message over SMTP"
                                    if ( SendUsingSMTP -Mime $saveToPickupMime -recipients $resendTo -Sender $NDR.Sender.Address )
                                    {
                                        $ndrProcessed = $true
                                    }
                                    else
                                    {
                                        $ndrProcessFail = $true
                                    }
                                }
                                else
                                {
                                    if ( $null -eq $SaveToPickupFolder )
                                    {
                                        # Send message from the mailbox
                                        LogVerbose "Resending message"
                                        $EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($script:service)
                                        $bytes = [byte[]]::new($saveToPickupMime.Length)
                                        [System.Text.Encoding]::UTF8.GetBytes($saveToPickupMime, 0, $saveToPickupMime.Length, $bytes, 0)
                                        $EmailMessage.MimeContent = New-Object Microsoft.Exchange.WebServices.Data.MimeContent("UTF-8", $bytes)
                                        try
                                        {
                                            $EmailMessage.Send()
                                            $ndrProcessed = $true
                                        } catch
                                        {
                                            ReportError "Send message"
                                            $ndrProcessFail = $true
                                        }
                                    }
                                    else
                                    {
                                        # Save message to pickup folder
                                        $ndrProcessed = SaveMIMEToPickupFolder -mime $saveToPickupMime -WasEncrypted $ndrEncrypted
                                        LogVerbose "Save to pickup folder success: $ndrProcessed"
                                        $ndrProcessFail = !$ndrProcessed
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    LogVerbose "Failed to read attached message: $Error[0]"
                    $ndrProcessFail = $true
                }
            }
            else
            {
                $ndrProcessFail = $true
                LogVerbose "Original message not attached to NDR"
            }
        }
        
        if ($SearchSentItems)
        {
            $ndrBody = [String]::Empty
            foreach ($prop in $NDR.ExtendedProperties)
            {
                if ($prop.PropertyDefinition -eq $PidTagBody)
                {
                    $ndrBody = $prop.Value
                }
            }

            if ([String]::IsNullOrEmpty($ndrBody))
            {
                # Failed to read the text body, so we'll try to read it from MessageBody
                $ndrBody = $NDR.Body.Text
            }

            #LogVerbose $ndrBody

            if (![String]::IsNullOrEmpty($ndrBody))
            {
                # Read the info
                LogVerbose "Attempting to extract message Id and sender from NDR"
                $messageId = ExtractHeaderValue $ndrBody "Message-ID"
                $fromHeader = ExtractHeaderValue $ndrBody "From"
                if (![String]::IsNullOrEmpty($messageId) -and ![String]::IsNullOrEmpty($fromHeader))
                {
                    LogVerbose "Attempting to resend message $messageId from $fromHeader"
                    FindAndResendMessage $messageId $fromHeader
                }
                else
                {
                    LogVerbose "Unable to read required information from NDR for resending"
                }
            }
            else
            {
                LogVerbose "Failed to read body of NDR"
            }
        }

        if ($WhatIf)
        {
            # We work out whether this message was processed or failed by its location (we are collecting statistics only)
            if ($NDR.ParentFolderId -eq $script:moveProcessedItemsToFolderFolder)
            {
                # This item was successfully processed
                $ndrProcessed = $true
            }
            elseif ($NDR.ParentFolderId -eq $script:moveErrorItemsToFolderFolder)
            {
                # This was an error item (we did process it, but encountered an error)
                $ndrProcessed = $true
                $ndrProcessFail = $true
            }
            elseif ($NDR.ParentFolderId -eq $script:moveEncryptedItemsToFolderFolder)
            {
                # This was an encrypted item (which we didn't process originally)
                $ndrEncrypted = $true
            }
        }

        $addItemToIgnoreList = $AddAllItemsToIgnoreLog
        if ($ndrProcessed)
        {
            # NDR has been processed, so check if we need to move it
            $script:processedItems++
            if ( ($null -ne $script:moveProcessedItemsToFolderFolder) -and (!$WhatIf) )
            {
                LogVerbose "Moving processed item"
                try
                {
                    [void]$NDR.Move($script:moveProcessedItemsToFolderFolder.Id)
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                }
                ReportError
            }

            # Update our resent to stats
            # $script:resentCountByEmailAddress = @{}
            foreach ($recipient in $resendTo)
            {
                if ($script:resentCountByEmailAddress.ContainsKey($recipient) )
                {
                    $script:resentCountByEmailAddress[$recipient]++
                }
                else
                {
                    $script:resentCountByEmailAddress[$recipient] = 1
                }
            }
        }

        if ($ndrEncrypted -and !$ndrProcessed)
        {
            # NDR is encrypted, so check if we need to move it
            $script:ignoredItems++
            $script:encryptedItems++
            if ( ($null -ne $script:moveEncryptedItemsToFolderFolder) -and (!$WhatIf))
            {
                LogVerbose "Moving encrypted item"
                try
                {
                    [void]$NDR.Move($script:moveEncryptedItemsToFolderFolder.Id)
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                }
                ReportError
            }
        }

        if ($ndrProcessFail)
        {
            # We encountered an issue processing this NDR, so we move to error folder
            
            $script:errorItems++
            if ( ($null -ne $script:moveErrorItemsToFolderFolder) -and (!$WhatIf) )
            {
                LogVerbose "Moving error item"
                try
                {
                    $movedItem = $NDR.Move($script:moveErrorItemsToFolderFolder.Id)
                    Log "Failed item id (moved to error folder): $($movedItem.Id.UniqueId)" Red
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                    Log "Failed item id (move failed): $($NDR.Id.UniqueId)" Red
                }
                ReportError
            }
        }

        # Check if the statistics file already exists (if not, we want to write CSV headers)
        if (-not [String]::IsNullOrEmpty($StatsCSV))
        {
            if (-not (Test-Path $StatsCSV))
            {
                "`"Message Id`",`"Sender`",`"Subject`",`"Original Sent Date`",`"Journal Address`",`"Processed`",`"Failed`",`"Encrypted`"" | Out-File $StatsCSV
                if (-not (Test-Path $StatsCSV))
                {
                    # Failed to create the CSV, so we can't write it
                    Log "Statistics CSV creation failed, statistics will not be collected" Red
                    $StatsCSV = ""
                }
            }
        }

        if (![String]::IsNullOrEmpty($StatsCSV))
        {
            # Collect the item statistics - we want the target email address, subject, message id, time originally sent
            # $resendTo contains the target email address(es)
            $messageId = ExtractHeaderValue $MIME "Message-ID"
            $fromHeader = ExtractHeaderValue $MIME "From"
            $subject = ExtractHeaderValue $MIME "Subject"
            $sentTime = ExtractHeaderValue $MIME "Date"
            foreach ($targetAddress in $resendTo)
            {
                "`"$messageId`",`"$fromHeader`",`"$subject`",`"$sentTime`",`"$targetAddress`",`"$ndrProcessed`",`"$ndrProcessFail`",`"$ndrEncrypted`"" | Out-File $StatsCSV -Append
            }
        }

        if ($addItemToIgnoreList -and (![String]::IsNullOrEmpty($IgnoreIdsLog)))
        {
            $script:ignoreIds += $NDR.Id.UniqueId
            if (![String]::IsNullOrEmpty($IgnoreIdsLog))
            {
                # Write this Id to the log file so that we ignore it on future runs
                $NDR.Id.UniqueId | out-file -FilePath $IgnoreIdsLog -Append
            }
        }

        if ($MaxItemsToProcess -gt 0)
        {
            if ($script:processedItems -ge $MaxItemsToProcess)
            {
                # We've processed our maximum number of items
                Log "$($script:processedItems) item(s) processed.  Stopping further processing." Green
                break
            }
        }
    }
    Write-Progress -Activity $progressActivity -Status "$($script:processedItems) NDRs processed" -Completed
    Log $(PerfReport) White
}

function SaveMIMEToPickupFolder()
{
    param (
        [String]$mime,
        [bool]$WasEncrypted
    )

    if (!$script:pickUpFolderList)
    {
        # Set up our list of folders for round robin
        $script:pickupFolderIndex = 0
        $script:pickUpFolderList = @()
        foreach ($pickupFolder in $SaveToPickupFolder)
        {
            if (![String]::IsNullOrEmpty($pickupFolder))
            {
                LogVerbose "Adding Pickup folder: $pickupFolder"
                $script:pickUpFolderList += $pickupFolder
            }
        }
        Log "$($pickUpFolderList.Length) pickup folder(s) being used (round robin)"
    }

    $messageIsValid = $true

    if ( ![String]::IsNullOrEmpty($ReturnPath) )
    {
        # We need to update Sender also, otherwise Exchange will overwrite Return-Path
        $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "Sender" -HeaderValue $ReturnPath
    }

    #$mime | out-file "c:\temp\sendagain\attachmimeinpickup.txt"
    #exit
    # Check that the message has a valid From header (Sender and Return-Path are optional, so we don't check these)
    $from = ExtractHeaderValue -headers $mime -HeaderName "From"
    if ([String]::IsNullOrEmpty($from))
    {
        if (![String]::IsNullOrEmpty($DefaultFromAddress))
        {
            # No from address found, but we have a default one to apply
            $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "From" -HeaderValue "From: $DefaultFromAddress"
            Log "From header was empty, replaced with `"From: $DefaultFromAddress`"" Yellow
        }
        else
        {
            Log "From header was empty, message not saved to pickup folder: $from" Red
            $messageIsValid = $false
        }
    }
    LogVerbose "From header: $from"
    if (!$RemoveEncryptedAttachments -and $WasEncrypted)
    {
        $messageIsValid = $false
        Log "Encrypted message not processed as -RemoveEncryptedAttachments not specified (encrypted attachments must be removed for successful processing)" Yellow
        return $false # We don't save these messages to any Pickup debug folder as we know why it has failed
    }

    if ( $messageIsValid )
    {
        $filename = "$($script:pickUpFolderList[$script:pickupFolderIndex])\$([DateTime]::Now.Ticks).eml"
        $script:pickupFolderIndex++
        if ($script:pickupFolderIndex -ge $pickUpFolderList.Length) { $script:pickupFolderIndex = 0 }

        try
        {
            Log "Saving email to: $fileName" Gray
            [IO.File]::WriteAllText($fileName, $mime)
            return $true
        }
        catch
        {
            ReportError
            return $false # No point in debugging a write failure, as this will be an IO issue                                     
        }
    }

    # If we get to this point, the message failed validation
    if ( ![String]::IsNullOrEmpty($FailPickupFolder) )
    {
        $filename = "$FailPickupFolder\$([DateTime]::Now.Ticks).eml"
        try
        {
            Log "Saving debug email to: $fileName" Gray
            [IO.File]::WriteAllText($fileName, $mime)
        }
        catch
        {
            ReportError
        }
    }
    return $false
}

$script:processedItems = 0
$script:ignoredItems = 0
$script:errorItems = 0
$script:encryptedItems = 0
$script:totalBytesRetrieved = 0
$script:resentCountByEmailAddress = @{}
function PerfReport()
{
    $elapsedTime = [DateTime]::Now.Subtract($script:startProcessTime)
    $totalProcessed = $script:processedItems + $script:errorItems + $script:encryptedItems
    if ($script:processedItems -gt 0)
    {
        if ($elapsedTime.TotalMinutes -gt 0)
        {
            $processedPerMinute = [Math]::Round($script:processedItems / $elapsedTime.TotalMinutes)
            $totalProcessedPerMinute = [Math]::Round($totalProcessed / $elapsedTime.TotalMinutes)
        }
        else
        {
            $processedPerMinute = $script:processedItems
            $totalProcessedPerMinute = $totalProcessed
        }
    }
    else
    {
        $processedPerMinute = 0
        $totalProcessedPerMinute = 0
    }

    $average = 0
    $averageMb = "0Mib"
    if ($script:processedItems -gt 0)
    {
        $average = [Math]::Round(($script:totalBytesRetrieved / $script:processedItems) / (1024*1024),2)
        $averageMb = "$($average)Mib"
    }

    $mBRetrieved = "$([Math]::Round($script:totalBytesRetrieved / (1024*1024),2))Mib"
    return "Runtime: $($elapsedTime.ToString("hh\:mm\:ss"))  Resent $processedPerMinute per minute (processed $totalProcessedPerMinute). Resent: $($script:processedItems)  Errors: $($script:errorItems)  Encrypted: $($script:encryptedItems)  Ignored: $($script:ignoredItems)  Total: $($script:processedItems + $script:errorItems)  Total size: $mbRetrieved  Average size: $averageMb"
}

function ProcessNDRs()
{
    param ( [Microsoft.Exchange.WebServices.Data.Folder]$folder )

    # Performance tracking
    $script:startProcessTime = [DateTime]::Now


	# Set parameters - we will process in batches of 500 for the FindItems call
	$Offset = 0
	$PageSize = 1000
    if ($BatchSize -gt 0) { $PageSize = $BatchSize }
    if ($MaxItemsToProcess -gt 0)
    {
        if ($PageSize -gt $MaxItemsToProcess) { $PageSize = $MaxItemsToProcess }
    }
	$MoreItems = $true

    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time)
    $itemsToResend = New-Object System.Collections.ArrayList
	
    $progressActivity = "Reading NDRs in folder $($Mailbox):$(GetFolderPath($folder))"
    LogVerbose "Building list of NDRs"
    Write-Progress -Activity $progressActivity -Status "0 NDRs found" -PercentComplete -1

    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "REPORT.IPM.Note.NDR")
    $ndrsFound = 0
    $noneNdrsFound = 0
    $retries = 0

	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,[Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)

        $FindResults = $null
        try
        {
            if ($FilterNDRsClientside)
            {
                $FindResults = $folder.FindItems($View)
            }
            else
            {
                $FindResults = $folder.FindItems($SearchFilter, $View)
            }
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
                $retries++
                if ($retries -lt 3)
                {
                    Log "Waiting 5 minutes until retry" Yellow
                    Start-Sleep -Seconds 360
                }
                else
                {
                    $MoreItems = $false # We've retried 3 times, so we fail
                }
                $FindResults = $null
            }
        }
		
        if ($FindResults)
        {
		    ForEach ($item in $FindResults.Items)
		    {
                if ($item.ItemClass.ToUpper().Equals("REPORT.IPM.NOTE.NDR"))
                {
                    [void]$itemsToResend.Add($item)
                    $ndrsFound++
                }
                else
                {
                    $noneNdrsFound++
                }
		    }
		    $MoreItems = $FindResults.MoreAvailable
            if ($MoreItems)
            {
                LogVerbose "$($itemsToResend.Count) items found so far, more available"
            }
		    $Offset += $PageSize
        }

        if ($itemsToResend.Count -ge $BatchSize)
        {
            # We have at least 1000 message to process, so we'll do this now
            LogVerbose "$($itemsToResend.Count) NDR(s) in batch; attempting to resend"
            ResendMessages $itemsToResend
            $itemsToResend = New-Object System.Collections.ArrayList
        }

        Write-Progress -Activity $progressActivity -Status "$($itemsToResend.Count) NDRs found" -PercentComplete -1
        if ($MaxItemsToProcess -gt 0)
        {
            if ($script:processedItems -ge $MaxItemsToProcess)
            {
                # We've processed our maximum number of items
                Log "Limit reached: $script:processedItems processed"
                break
            }
        }
	}
    Write-Progress -Activity $progressActivity -Status "$($itemsToResend.Count) NDRs found" -Completed

    if ( $itemsToResend.Count -gt 0 )
    {
        if ( ($MaxItemsToProcess -lt 1) -or ($script:processedItems -lt $MaxItemsToProcess) )
        {
            Log "$($itemsToResend.Count) NDR(s) in final batch; attempting to resend" Green
            ResendMessages $itemsToResend
        }
        else
        {
            Log "$($itemsToResend.Count) items in final batch, but already reached maximum processed limit" Yellow
        }
    }

    Log "Folder analysis complete.  Analysed $ndrsFound NDRs, ignored $noneNdrsFound other item(s).  Total items analysed $($ndrsFound+$noneNdrsFound)" Green
}

function ProcessMailbox()
{
    # Process the mailbox
    if ( [string]::IsNullOrEmpty($Mailbox) )
    {
        Log "ProcessMailbox called with no mailbox set" Red
        return
    }
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($null -eq $script:service)
	{
		Log "Failed to create ExchangeService" Red
        return
	}

    $script:throttlingDelay = 0

    # Bind to root folder	
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    $Folder = $Null
    if ([String]::IsNullOrEmpty($FolderPath))
    {
        $FolderPath = "wellknownfoldername.Inbox"
    }
    if ($FolderPath.ToLower().StartsWith("wellknownfoldername."))
    {
        # Well known folder specified (could be different name depending on language, so we bind to it using WellKnownFolderName enumeration)
        $wkf = $FolderPath.SubString(20)
        LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind $folderId $null $script:service
    }
    else
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
        $Folder = ThrottledFolderBind $folderId $null $script:service
        if ($Folder -and ($FolderPath -ne "\"))
        {
	        $Folder = GetFolder $Folder $FolderPath
        }
    }

	if (!$Folder)
	{
		Log "Failed to find folder $FolderPath" Red
		return
	}

    # Check we can access our processed and error folders before we do any processing
    $script:moveProcessedItemsToFolderFolder = ValidateFolderMoveParameter $MoveProcessedItemsToFolder $script:moveProcessedItemsToFolderFolder
    $script:moveErrorItemsToFolderFolder = ValidateFolderMoveParameter $MoveFailedItemsToFolder $script:moveErrorItemsToFolderFolder
    $script:moveEncryptedItemsToFolderFolder = ValidateFolderMoveParameter $MoveEncryptedItemsToFolder $script:moveEncryptedItemsToFolderFolder

    if ($WhatIf)
    {
        # If we are only collecting stats, then we simply reprocess each of the target folders without resending
        if ($script:moveProcessedItemsToFolderFolder)
        {
            ProcessNDRs $script:moveProcessedItemsToFolderFolder
        }
        if ($script:moveErrorItemsToFolderFolder)
        {
            ProcessNDRs $script:moveErrorItemsToFolderFolder
        }
        if ($script:moveEncryptedItemsToFolderFolder)
        {
            ProcessNDRs $script:moveEncryptedItemsToFolderFolder
        }
    }
    else
    {
        ProcessNDRs $Folder
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

Add-Type -AssemblyName System.Web

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($OAuth)
    {
        Write-Host "Please specify *either* -Credentials *or* -OAuth" Red
        Exit
    }
}

$PidTagBody = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1000, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String) 
$script:currentThrottlingDelay = 500

Write-Host ""

# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
    Write-Verbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $Mailbox -Header "PrimarySmtpAddress"
	foreach ($entry in $csv)
	{
        Write-Verbose $entry.PrimarySmtpAddress
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


if ($script:resentCountByEmailAddress.Count -gt 0)
{
    Log "Resend Statistics by email address" -SuppressWriteToScreen
    foreach ($recipientAddress in $script:resentCountByEmailAddress.Keys)
    {
        Log "$($recipientAddress): $($script:resentCountByEmailAddress[$recipientAddress])" -SuppressWriteToScreen
    }
    $script:resentCountByEmailAddress
}
else
{
    Log "No resend statistics collected (implies no messages were resent)"
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
