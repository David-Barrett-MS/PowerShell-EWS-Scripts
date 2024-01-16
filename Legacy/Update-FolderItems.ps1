#
# Update-FolderItems.ps1
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
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox will be accessed (instead of the main mailbox).")]
    [switch]$Archive,
		
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder.")]
    [switch]$PublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox message root folder is assumed.")]
    $FolderPath,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, subfolders will also be processed.")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, hidden (associated) items of the folder are processed (normal items are ignored).")]
    [switch]$AssociatedItems,
	
    # Generic item processing

    [Parameter(Mandatory=$False,HelpMessage="If specified, all EWS first class properties will be requested when retrieving items.  Useful when using -ListMatches.")]
    [switch]$LoadAllItemProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the item(s). Default is a soft delete.")]
    [switch]$Delete,
    
    [Parameter(Mandatory=$False,HelpMessage="When used with -Delete, forces a hard delete of the item(s).")]
    [switch]$HardDelete,

    [Parameter(Mandatory=$False,HelpMessage="Adds the given property(ies) to the item(s) (must be supplied as hash table @{}).")]
    $AddItemProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the given property(ies) from the item(s).")]
    $DeleteItemProperties,

    # Calendar processing

    [Parameter(Mandatory=$False,HelpMessage="Accepts any matching appointment requests.")]
    [switch]$AcceptCalendarInvite,

    [Parameter(Mandatory=$False,HelpMessage="Performs a check for conflicts before accepting a meeting (if there is a conflict, the meeting will be declined).  Must be used with -AcceptCalendarInvite.")]
    [switch]$DeclineCalendarInviteIfConflict,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the subject is removed from the appointment (only for the accepted meeting).")]
    [switch]$DeleteSubject,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the organizer is appended to the appointment subject (only for the accepted meeting).")]
    [switch]$AddOrganizerToSubject,

    [Parameter(Mandatory=$False,HelpMessage="Declines any matching appointment requests.")]
    [switch]$DeclineCalendarInvite,

    # Contact processing

    [Parameter(Mandatory=$False,HelpMessage="Actions will only apply to contact objects that have the given SMTP address as their email address.  Supports multiple SMTP addresses passed as an array.")]
    $MatchContactAddresses,

    [Parameter(Mandatory=$False,HelpMessage="If any matching contact object contains a contact photo, the photo is deleted.")]
    [switch]$DeleteContactPhoto,
    
    # Email processing
    
    [Parameter(Mandatory=$False,HelpMessage="Marks the item(s) as read.")]
    [switch]$MarkAsRead,
    
    [Parameter(Mandatory=$False,HelpMessage="Marks the item(s) as unread.")]
    [switch]$MarkAsUnread,

    # Email Resend

    [Parameter(Mandatory=$False,HelpMessage="Resends the message (resend options must also be set).")]
    [switch]$Resend,

    [Parameter(Mandatory=$False,HelpMessage="Creates a draft of the message that will be resent (in the Drafts folder of the mailbox).  Message will not be sent.")]
    [switch]$ResendCreateDraftOnly,

    [Parameter(Mandatory=$False,HelpMessage="Sets the sender for the resent message.")]
    $ResendFrom = "",

    [Parameter(Mandatory=$False,HelpMessage="Sets the recipient for the resent message.")]
    $ResendTo = "",

    [Parameter(Mandatory=$False,HelpMessage="Prepends the provided text to the resent message body.")]
    $ResendPrependText = "",

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, the text to be prepended will be modified per field values.  e.g. <!-- %ORIGINALSENDER% --> will be replaced with the original sender email.")]
    [switch]$ResendUpdatePrependTextFields,

    [Parameter(Mandatory=$False,HelpMessage="Resends the message to the recipient declared in the message Received: header (if present).")]
    [switch]$ResendToForInReceivedHeader,

    # Filters

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given AQS filter will be processed `r`n(see https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-an-aqs-search-by-using-ews-in-exchange ).  Do not use any other restrictions with this option.")]
    [string]$SearchFilter,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that have recipients not from the listed domains will be matched.")]
    $RecipientsNotFromDomains,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that have recipients not from the listed addresses will be matched.")]
    $RecipientsNotFromAddresses,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that have recipients from the listed domains will be matched.")]
    $RecipientsFromDomains,

    [Parameter(Mandatory=$False,HelpMessage="If specified, CC recipients will not be checked for domain or email address matches.")]
    $ExcludeCCRecipients,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items where the sender is not from one of the listed domains will be matched.")]
    $SenderNotFromDomains,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items where the sender is not from one of the listed addresses will be matched.")]
    $SenderNotFromAddresses,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items where the sender is from one of the listed domains will be matched.")]
    $SenderFromDomains,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items where the sender is from one of the listed addresses will be matched.")]
    $SenderFromAddresses,

    [Parameter(Mandatory=$False,HelpMessage="Only processes items created after this date.")]
    [datetime]$CreatedAfter,
	
    [Parameter(Mandatory=$False,HelpMessage="Only processes items created before this date.")]
    [datetime]$CreatedBefore,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that have values in the given properties will be updated.")]
    $PropertiesMustExist,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given values in the given properties will be updated.  Properties must be supplied as a Dictionary: @{""propId"" = ""value""}")]
    $PropertiesMustMatch,

    [Parameter(Mandatory=$False,HelpMessage="Outputs any matching items (can be collected for further processing).")]
    [switch]$ListMatches,

    [Parameter(Mandatory=$False,HelpMessage="If set, a separate GetItem request is sent to retrieve each item.  Much slower (batch processing is used otherwise), but may need to be used if querying large properties.")]
    [switch]$LoadItemsIndividually,

    [Parameter(Mandatory=$False,HelpMessage="If this is set to any value higher than 0, then the script will go into -WhatIf mode once that many items have been processed.")]
    $MaximumNumberOfItemsToProcess = 0,

    [Parameter(Mandatory=$False,HelpMessage="If set, the script will stop processing further items once MaximumNumberOfItemsToProcess limit is reached.")]
    [switch]$StopAfterMaximumNumberOfItemsProcessed,

    [Parameter(Mandatory=$False,HelpMessage="The number of items that will be requested in a single GetItem call.  Reduce this significantly (e.g. 10) if items are large and need to be retrieved.")]
    $GetItemBatchSize = 500,
    
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

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, no items will be changed (but any processing that would occur will be logged).")]	
    [switch]$WhatIf
)
$script:ScriptVersion = "1.3.6"

if ($ForceTLS12)
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
else
{
    Write-Host "If having connection/auth issues for Exchange Online or hybrid, you may need -ForceTLS12 switch" -ForegroundColor Gray
}


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
    if ($script:oAuthAccessToken -ne $null)
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

        if ("EWSTracer" -as [type]) {} else {
            Add-Type -TypeDefinition $TraceListenerClass -ReferencedAssemblies $EWSManagedApiPath
        }
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

Function Throttled()
{
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning

    if ([String]::IsNullOrEmpty($script:Tracer.LastResponse))
    {
        if ( $Error[0].Exception.Message.Contains("The server cannot service this request right now.") )
        {
            # We've got a generic throttling error, so we'll pause for 15 seconds
            Start-Sleep -Seconds 15
            return $true
        }
        return $false # Throttling does return a response, if we don't have one then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds")
    {
        # We are throttled, and the server has told us how long to back off for
        # We back off for the time given by the server
        Log "Throttling detected; server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
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
        $propset = $null
    )

    ApplyEWSOAuthCredentials
    LogVerbose "Attempting to bind to folder $folderId"
    try
    {
        if ($propset -eq $null)
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
        else
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset)
        }
        Start-Sleep -Milliseconds $script:throttlingDelay
        if (-not ($folder -eq $null))
        {
            LogVerbose "Successfully bound to $($folderId): $($folder.DisplayName)"
        }
        return $folder
    }
    catch
    {
    }

    if (Throttled)
    {
        ApplyEWSOAuthCredentials
        try
        {
            if ($propset -eq $null)
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
            }
            else
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId, $propset)
            }
            if (-not ($folder -eq $null))
            {
                LogVerbose "Successfully bound to $($folderId): $($folder.DisplayName)"
            }
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    return $null
}

function ThrottledItemBind()
{
    param (
        [Microsoft.Exchange.WebServices.Data.ItemId]$itemId,
        $propset = $null
    )

    ApplyEWSOAuthCredentials
    LogVerbose "Attempting to bind to item $itemId"
    if ($propset -eq $null)
    {
        $propset = $script:RequiredPropSet
    }
    try
    {
        $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($script:service, $itemId, $propset)
        Start-Sleep -Milliseconds $script:throttlingDelay
        if (-not ($item -eq $null))
        {
            LogVerbose "Successfully bound to item $($itemId): $($item.Subject)"
        }
        return $item
    }
    catch
    {
    }

    if (Throttled)
    {
        ApplyEWSOAuthCredentials
        try
        {
            $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($script:service, $itemId, $propset)
            if (-not ($item -eq $null))
            {
                LogVerbose "Successfully bound to item $($itemId): $($item.Subject)"
            }
            return $item             
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    return $null
}

function ThrottledItemDelete()
{
    param (
        $item
    )

    ApplyEWSOAuthCredentials

    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems
    if ($HardDelete)
    {
        $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
    }
    try
    {
        $item.Delete($deleteMode)
        LogVerbose "Item deleted"
        return $True
    }
    catch {}

    if (Throttled)
    {
        ApplyEWSOAuthCredentials
        try
        {
            $item.Delete($deleteMode)
            LogVerbose "Item deleted"
            return $True
        }
        catch
        {
            Log "Failed to delete item: $($Error[0])" Red
        }
    }
    return $false
}

function ThrottledItemUpdate()
{
    param (
        $item
    )

    $isAppointment = ($item.GetType() -eq [Microsoft.Exchange.WebServices.Data.Appointment])
    
    if ($WhatIf)
    {
        if ($isAppointment)
        {
            Log "Appointment would be updated, but -WhatIf specified" Green
        }
        else
        {
            Log "Item would be updated, but -WhatIf specified" Green
        }
        return $True
    }

    ApplyEWSOAuthCredentials
    try
    {
        if ($isAppointment)
        {
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone)
            LogVerbose "Appointment updated"
        }
        else
        {
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
            LogVerbose "Item updated"
        }
        return $True
    }
    catch {}

    if (Throttled)
    {
        ApplyEWSOAuthCredentials
        try
        {
            if ($isAppointment)
            {
                $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone)
                LogVerbose "Appointment updated"
            }
            else
            {
                $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
                LogVerbose "Item updated"
            }
            return $True
        }
        catch
        {
            Log "Failed to update item: $($Error[0])" Red
        }
    }

    try
    {
        $item.Load() # We do this to clear the failed update from the item object
    }
    catch {}
    return $false
}

Function AddItemProperties($item)
{
    # Add the specified properties to the item

    # First of all ensure we have some properties to add...
    if ($AddItemProperties -eq $Null) { return $False }

    # We need to convert the properties to EWS extended properties
    if ($script:addItemPropsEws -eq $Null)
    {
        Write-Verbose "Building list of properties to add"
        $script:addItemPropsEws = @{}
        foreach ($addProperty in $AddItemProperties.Keys)
        {
            $value = $AddItemProperties[$addProperty]
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
            if ($propdef -ne $Null)
            {
                $script:addItemPropsEws.Add($propdef, $value)
                Write-Verbose "Added property $addProperty to add list; value is $($script:addItemPropsEws[$propdef])"
            }
            else
            {
                Log "Failed to parse (or convert) property $addProperty" Red
            }
        }
    }

    # Now we add the properties to the item
    foreach ($addProperty in $script:addItemPropsEws.Keys)
    {
        LogVerbose "Property $($addProperty) set to $($script:addItemPropsEws[$addProperty])"
        $item.SetExtendedProperty($addProperty, $script:addItemPropsEws[$addProperty])
    }

    # Now update the item
    if (ThrottledItemUpdate $item)
    {
        LogVerbose "Item updated (properties added)" Green
        return $True
    }
    return $False
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
    $EWSPropDef = $null

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
        # Assume MAPI property (e.g. 0x00360003)
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

    if ($propdef -ne $Null)
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
        if ($EWSPropDef -ne $Null)
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
            if ($EWSPropDef -ne $Null)
            {
                $script:propertiesMustMatchEws.Add($EWSPropDef, $PropertiesMustMatch[$PropDef])
            }

        }
    }
}

Function DeleteItemProperties($item)
{
    # Delete the specified properties from the item

    # Ensure we have some properties to delete...
    if ( ($script:deleteItemPropsEws -eq $Null) -or ($item -eq $null) )
    {
        return
    }

    $propDeleted = $False
    # Delete the properties from the item
    LogVerbose "Checking for properties to delete"
    foreach ($deleteProperty in $script:deleteItemPropsEws)
    {
        foreach ($extendedProperty in $item.ExtendedProperties)
        {
            # Check if this extended property is one marked for deletion

            LogVerbose "Item property $($extendedProperty.PropertyDefinition), delete property $($deleteProperty)"

            if ( $deleteProperty.Equals($extendedProperty.PropertyDefinition) )
            {
                if (!$item.RemoveExtendedProperty($deleteProperty))
                {
                    Log "Failed to remove property $($deleteProperty)" Red
                }
                else
                {
                    LogVerbose "Property $($deleteProperty) deleted"
                    $propDeleted = $True
                }
                break
            }
        }
    }

    # Now update the item
    if ($propDeleted)
    {
        if (ThrottledItemUpdate $item)
        {
            LogVerbose "Item updated (properties deleted)" Green
            return $True
        }
    }
    return $False
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath, $Create = $args[0]
	
    if ( $RootFolder -eq $null )
    {
        LogVerbose "GetFolder called with null root folder"
        return $null
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
                ApplyEWSOAuthCredentials
                try
                {
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                    Start-Sleep -Milliseconds $script:throttlingDelay
                }
                catch {}
                if ($FolderResults -eq $Null)
                {
                    if (Throttled)
                    {
                        ApplyEWSOAuthCredentials
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
					Write-Host "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red
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
				    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id
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

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
    $parentFolder = ThrottledFolderBind $Folder.Id $propset
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
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

Function DeleteContactPhoto($item)
{
    # If the item is a contact, check for and delete any photo attached to the item

    if (!$item.ItemClass.Equals("IPM.Contact") )
        { return $false }

    if (!$item.HasPicture)
    {
        LogVerbose "Contact object has no picture"
        return $false
    }

    # We need to load the Contact object and attachments
    LogVerbose "Checking for contact photo"
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ContactSchema]::Attachments)
    $contact = ThrottledItemBind($item.Id, $propset)

    foreach ($attachment in $contact.Attachments)
    {
        if ($attachment.IsContactPhoto)
        {
            # This is the attachment we need to delete
            if ($WhatIf)
            {
                LogVerbose "Contact photo would be removed"
            }
            else
            {
                try
                {
                    $contact.Attachments.Remove($attachment)
                    ThrottledItemUpdate $contact
                    LogVerbose "Contact photo removed"
                }
                catch
                {
                    Log "Error removing contact photo: $($Error[0])" Red
                    return $false
                }
            }
            return $true
        }
    }

    return $false
}

function MarkWhetherRead($item)
{
    # Check if we need to update the item's read/unread status

    if (!$MarkAsRead -and !$MarkAsUnread) { return $false }

    $currentStatus = $item.IsRead
    if ($item.IsRead -and $MarkAsUnread)
    {
        $item.IsRead = $false
    }
    elseif (!$item.IsRead -and $MarkAsRead)
    {
        $item.IsRead = $true
    }
    if ($currentStatus -ne $item.IsRead)
    {
        # Now update the item
        if ($WhatIf)
        {
            LogVerbose "Would update read status of message, IsRead = $($item.IsRead)"
            return $true
        }
        else
        {
            if (ThrottledItemUpdate $item)
            {
                LogVerbose "Updated read status of message, IsRead = $($item.IsRead)"
                return $true
            }
        }
    }
    return $false
}

function ItemHasRequiredProperties($item)
{
    # Check that this item matches any property requirements
    if ($script:propertiesMustExistEws -ne $null)
    {
        foreach ($requiredProperty in $script:propertiesMustExistEws)
        {
            # Check the item has this property
            $propExists = $false
            if (![String]::IsNullOrEmpty(($requiredProperty.PropertySetId)))
            {
                LogDebug "Checking for $($requiredProperty.PropertySetId)"
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
            elseif ($requiredProperty.Tag -ne $null)
            {
                LogDebug "Checking for $($requiredProperty.Tag)"
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
    }
    return $true
}

function ItemPropertiesMatchRequirements($item)
{
    if ($script:propertiesMustMatchEws -ne $null)
    {
        foreach ($requiredProperty in $script:propertiesMustMatchEws.Keys)
        {
            # Check the item has this property
            $propMatches = $false

            foreach ($itemProperty in $item.ExtendedProperties)
            {
                if (![String]::IsNullOrEmpty(($requiredProperty.PropertySetId)))
                {
                    if ($requiredProperty.PropertySetId -eq $itemProperty.PropertyDefinition.PropertySetId)
                    {
                        # Same property set, check the value
                        if (![String]::IsNullOrEmpty(($requiredProperty.Id)))
                        {
                            if ($requiredProperty.Id -eq $itemProperty.PropertyDefinition.Id)
                            {
                                $propMatches = ($itemProperty.Value -eq $script:propertiesMustMatchEws[$requiredProperty])
                                break
                            }
                        }
                        elseif (![String]::IsNullOrEmpty(($requiredProperty.Name)))
                        {
                            if ($requiredProperty.Name -eq $itemProperty.PropertyDefinition.Name)
                            {
                                $propMatches = ($itemProperty.Value -eq $script:propertiesMustMatchEws[$requiredProperty])
                                break
                            }
                        }
                    }
                }
                elseif ($requiredProperty.Tag -ne $null -and $requiredProperty.Tag -eq $itemProperty.PropertyDefinition.Tag)
                {
                    # Check MAPI extended property
                    if ($itemProperty.PropertyDefinition.MapiType -eq [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
                    {
                        # This is a binary property, so we use Compare-Object (which will return null if the objects are identical)
                        if ( (Compare-Object -ReferenceObject $script:propertiesMustMatchEws[$requiredProperty] -DifferenceObject $itemProperty.Value) -eq $null)
                        {
                            $propMatches = $true
                        }
                    }
                    else
                    {
                        $propMatches = ($itemProperty.Value -eq $script:propertiesMustMatchEws[$requiredProperty])
                    }
                    break
                }
            }
            if (!$propMatches)
            {
                # If any single property does not match, we don't bother to check any further
                Write-Verbose "$requiredProperty does not match, ignoring item"
                return $false
            }
        }
    }
    return $true    
}

Function ItemMatchesDateRequirement
{
    param ($item)

    if (!$CreatedAfter -and !$CreatedBefore)
    {
        return $true
    }

    # Check if we have creation date criteria
    $createdTime = $null
    if ($item.ExtendedProperties.Count -gt 0)
    {
        foreach ($prop in $item.ExtendedProperties)
        {
            if ($prop.PropertyDefinition -eq $script:PR_CREATION_TIME)
            {
                $createdTime = $prop.Value
                LogVerbose "Folder created: $createdTime"
            }
        }
    }

    if ($createdTime -eq $null)
    {
        # If we can't read the creation time, we assume it doesn't match our criteria
        LogVerbose "Unable to retrieve PR_CREATION_TIME: $($item.Subject)"
        return $false
    }

    if ($CreatedAfter)
    {
        if ($createdTime -lt $CreatedAfter)
        {
            LogVerbose "Folder does not match CreatedAfter requirement: $($item.Subject)"
            return $false
        }
    }
    if ($CreatedBefore)
    {
        if ($createdTime -gt $CreatedBefore)
        {
            LogVerbose "Folder does not match CreatedBefore requirement: $($item.Subject)"
            return $false
        }
    }
    return $true
}

Function InitRecipientMatchInfo()
{
    $script:filterRecipients = $false
    $script:wildcardRecipientsNotFromDomains = @()
    $script:exactRecipientsNotFromDomains = @()
    $script:wildcardSenderNotFromDomains = @()
    $script:exactSenderNotFromDomains = @()    

    if ($RecipientsNotFromDomains -or $RecipientsFromDomains -or $RecipientsNotFromAddresses)
    {
        # We have recipient filters, so ensure we get recipient properties and initialise our checks

        $script:filterRecipients = $true

        if ($RecipientsNotFromDomains)
        {
            # We split wildcard domain matches into a separate list as these need special handling.  The only support wildcard format is *.domain.com

            foreach ($notFromDomain in $RecipientsNotFromDomains)
            {
                if ($notFromDomain.StartsWith("*."))
                {
                    # Wildcard domain match
                    $script:wildcardRecipientsNotFromDomains += $notFromDomain.Substring(1).ToLower()
                    # We also need to add the main domain to the exact match (otherwise we won't match the main domain, only sub-domains)
                    $script:exactRecipientsNotFromDomains += $notFromDomain.Substring(2).ToLower()
                }
                else
                {
                    $script:exactRecipientsNotFromDomains += $notFromDomain.ToLower()
                }
            }
        }
    }

    if ($SenderNotFromDomains -or $SenderFromDomains -or $SenderNotFromAddresses)
    {
        # We have sender filters
        $script:filterSender = $true

        if ($SenderNotFromDomains)
        {
            # We split wildcard domain matches into a separate list as these need special handling.  The only support wildcard format is *.domain.com

            foreach ($notFromDomain in $SenderNotFromDomains)
            {
                if ($notFromDomain.StartsWith("*."))
                {
                    # Wildcard domain match
                    $script:wildcardSenderNotFromDomains += $notFromDomain.Substring(1).ToLower()
                    # We also need to add the main domain to the exact match (otherwise we won't match the main domain, only sub-domains)
                    $script:exactSenderNotFromDomains += $notFromDomain.Substring(2).ToLower()
                }
                else
                {
                    $script:exactSenderNotFromDomains += $notFromDomain.ToLower()
                }
            }
        }
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
    }

    if ($script:filterRecipients)
    {
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ToRecipients)
        if (!$ExcludeCCRecipients)
        {
            $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::CcRecipients)
        }
    }
}

Function ItemMatchesRecipientRequirements($item)
{
    if ( !$script:filterRecipients -and !$script:filterSender)
    {
        return $true
    }

    # Perform sender checks
    $senderSMTPAddress = $null
    try
    {
        $senderSMTPAddress = $item.Sender.Address.ToLower()
    }
    catch
    {
        $script:LastError = $Error[0]
    }

    # Sender checks
    if (![string]::IsNullOrEmpty($senderSMTPAddress))
    {
        $senderDomain = $senderSMTPAddress.Substring($senderSMTPAddress.IndexOf('@')+1)
        if ($SenderNotFromDomains)
        {
            # If the sender is not from one of the given domains, then this message matches our filter
            $wildcardMatch = $false
            foreach ($wildcardDomain in $script:wildcardSenderNotFromDomains)
            {
                if ( $senderDomain.EndsWith($wildcardDomain) )
                {
                    $wildcardMatch = $true
                    break
                }
            }

            if (!$wildcardMatch -and -not ($script:exactSenderNotFromDomains.Contains($senderDomain)) )
            {
                if ($SenderNotFromAddresses)
                {
                    # If we have specific recipient addresses specified, we check the sender doesn't match those
                    if (!$SenderNotFromAddresses.Contains($senderSMTPAddress))
                    {
                        return $true
                    }
                }
                else
                {
                    return $true
                }
            }
        }
        elseif ($SenderNotFromAddresses)
        {
            if ($SenderNotFromAddresses.Contains($senderSMTPAddress))
            {
                return $false
            }
        }
    }

    if (!$script:filterRecipients)
    {
        return $true
    }

    # Get the domain part of each recipient
    $recipientDomains = @()

    if ($item.ToRecipients -ne $null -and $item.ToRecipients.Count -gt 0)
    {
        foreach ($recipient in $item.ToRecipients)
        {
            if ($recipient.RoutingType -eq "SMTP")
            {
                $recipientSMTPAddress = $recipient.Address
                if ($RecipientsNotFromAddresses)
                {
                    # If we have specific recipient addresses specified, we check those here and exclude from further checks
                    if (!$RecipientsNotFromAddresses.Contains($recipientSMTPAddress.ToLower()))
                    {
                        $recipientDomain = $recipientSMTPAddress.Substring($recipientSMTPAddress.IndexOf('@')+1)
                        $recipientDomains += $recipientDomain.ToLower()
                    }
                }
            }
        }
    }
    if (!$ExcludeCCRecipients -and $item.CcRecipients -ne $null -and $item.CcRecipients.Count -gt 0)
    {
        foreach ($recipient in $item.CcRecipients)
        {
            if ($recipient.RoutingType -eq "SMTP")
            {
                $recipientSMTPAddress = $recipient.Address
                $recipientDomain = $recipientSMTPAddress.Substring($recipientSMTPAddress.IndexOf('@')+1)
                $recipientDomains += $recipientDomain.ToLower()
            }
        }
    }

    $recipientMatch = $false
    if ($recipientDomains.Count -lt 1)
    {
        # No recipients to check
        return $false
    }

    if ($RecipientsNotFromDomains)
    {
        # If any recipients are not from the given domains, then this message matches our filter
        foreach ($checkDomain in $recipientDomains)
        {
            $wildcardMatch = $false
            foreach ($wildcardDomain in $script:wildcardRecipientsNotFromDomains)
            {
                if ( $checkDomain.EndsWith($wildcardDomain) )
                {
                    $wildcardMatch = $true
                    break
                }
            }

            if (!$wildcardMatch -and -not ($script:exactRecipientsNotFromDomains.Contains($checkDomain)) )
            {
                return $true
            }
        }
    }

    if ($RecipientsFromDomains)
    {
        # If any recipients are from the given domains, then this message matches our filter
        foreach ($checkDomain in $recipientDomains)
        {
            if ( $RecipientsFromDomains.Contains($checkDomain) )
            {
                $recipientMatch = $true
                break
            }
        }
    }

    return $recipientMatch
}

function ResendItem()
{
    # Attempt to resend the item
    $item = $args[0]

    if ($ResendToForInReceivedHeader)
    {
        # We need to parse the receieved headers to determine the original recipient of this message.  The headers are included in the MIME content, so we read from there
        $mimeContent = $item.MimeContent.ToString()
        $mimeHeaders = ""
        $mimeHeadersEnd = $mimeContent.IndexOf("`r`n`r`n")
        if ($mimeHeadersEnd -gt -1)
        {
            $mimeHeaders = $mimeContent.Substring(0,$mimeHeadersEnd)
        }
        if ([String]::IsNullOrEmpty($mimeHeaders))
        {
            Log "Failed to read MIME headers" Red
            return $false
        }

        # Parse the MIME headers
        # We are looking for the first Received: header that has for <x@x.com> (which should be the original intended recipient)
        $originalForAddress = ""
        $headerLines = $mimeHeaders -split "`r`n"
        $receivedForHeader = ""
        for ( $i = $headerLines.Count-1; $i -ge 0; $i--)
        {
            if ($headerLines[$i].StartsWith("Received:"))
            {
                $receivedForHeader = $headerLines[$i]
                $j = 1
                while ($headerLines[$i+$j].StartsWith("`t"))
                {
                    $receivedForHeader = "$receivedForHeader$($headerLines[$i+$j].SubString(1))"
                    if ($headerLines[$i+$j].StartsWith("`tfor <"))
                    {
                        # This is the for address we need to extract
                        $resendToAddressEnd = $headerLines[$i+$j].IndexOf(">")
                        if ($resendToAddressEnd -gt 6)
                        {
                            $ResendTo = $headerLines[$i+$j].Substring(6, $resendToAddressEnd-6)
                        }
                        $j = 0
                        $i = -1
                    }
                    else
                    {
                        $j++
                    }
                }
            }
            if (![String]::IsNullOrEmpty($ResendTo))
            {
                break
            }
        }

        if ([String]::IsNullOrEmpty($ResendTo))
        {
            Log "Failed to determine recipient to send to" Red
            return $false
        }

        if ($WhatIf)
        {
            Log "Would resend message to $ResendTo"
            return $true
        }

        LogVerbose "Resending message to $ResendTo"
        $resendMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $script:service
        $resendMessage.MimeContent = $item.MimeContent
        $resendMessage.ToRecipients.Clear()
        $resendMessage.ToRecipients.Add($ResendTo) | out-null

        if (![String]::IsNullOrEmpty($ResendFrom))
        {
            LogVerbose "Setting sender to $ResendFrom"
            $resendMessage.Sender = $ResendFrom
            $resendMessage.From = $ResendFrom
        }        

        $itemSaved = $false
        ApplyEWSOAuthCredentials
        if (![String]::IsNullOrEmpty($ResendPrependText))
        {
            # Prepend the given text to the message body
            # To do this, we need to save the item and then reload it so that we can retrieve the message body
            try
            {
                $resendMessage.Save()
                $itemSaved = $true
            }
            catch {}
            if (ErrorReported("ResendItem"))
            {
                return $false
            }

            $resendMessage = ThrottledItemBind($resendMessage.Id)
            if ($resendMessage -eq $null)
            {
                return $false
            }

            $prependText = $ResendPrependText
            if ($ResendUpdatePrependTextFields)
            {
                # Replace any fields that are defined in the prepended text
                # <!-- %ORIGINALSENDER% -->
                # <!-- %ORIGINALRECIPIENTS% -->
                # <!-- %ORIGINALSENTTIME% -->

                $prependText = $prependText.Replace("<!-- %ORIGINALSENDER% -->", $item.From)
                $prependText = $prependText.Replace("<!-- %ORIGINALRECIPIENTS% -->", $resendTo)
                $prependText = $prependText.Replace("<!-- %ORIGINALSENTTIME% -->", $item.DateTimeSent)
            }

            if ($resendMessage.Body.BodyType -eq [Microsoft.Exchange.WebServices.Data.BodyType]::HTML)
            {
                # Update HTML message body
                LogVerbose "Body type is HTML" -ForegroundColor Cyan
                $startOfHTMLBody = $resendMessage.Body.Text.IndexOf("<body")
                if ($startOfHTMLBody -gt -1)
                {
                    $insertionPoint = $resendMessage.Body.Text.IndexOf(">",$startOfHTMLBody)+1
                    if ($insertionPoint -gt -1)
                    {
                        LogVerbose "Text prepended" -ForegroundColor Cyan
                        $resendMessage.Body.Text = "$($resendMessage.Body.Text.Substring(0, $insertionPoint))<p>$prependText</p>$($resendMessage.Body.Text.Substring($insertionPoint))"
                    }
                }
            }
            else
            {
                # Update text message body
                LogVerbose "Body type is text, prepending text" -ForegroundColor Cyan
                $resendMessage.Body.Text = "$prependText`r`n`r`n$($resendMessage.Body.Text)"
            }
        }                        

        if ($ResendCreateDraftOnly)
        {
            try
            {
                if (!$itemSaved)
                {
                    ApplyEWSOAuthCredentials
                    try
                    {
                        $resendMessage.Save()
                        $itemSaved = $true
                    }
                    catch {}
                    if (ErrorReported("ResendItem"))
                    {
                        return $false
                    }
                }
                else
                {
                    ThrottledItemUpdate $resendMessage | out-null
                }
                $script:itemsResent++
            } catch {}
            if (!(ErrorReported("ResendItem")))
            {
                Log "Draft Resend message created" Green
                return $true
            }
        }
        else
        {
            ApplyEWSOAuthCredentials
            try
            {
                $resendMessage.Send()
                $script:itemsResent++
            }
            catch {}
            if (!(ErrorReported("ResendItem")))
            {
                Log "Message resent to $resendTo" Green
                return $true
            }
        }
    }
    return $false
}

Function AppointmentHasConflict($meetingInvitation)
{
    # Return $True if there is a conflict, $False otherwise

    LogVerbose "Checking for appointment conflict"
    $calendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($meetingInvitation.Start, $meetingInvitation.End, 2)
    $items = $script:service.FindAppointments([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $calendarView)
    return $items.Count -gt 1
}

Function ProcessItem()
{
	# Apply updates to the given item

    $item = $args[0]
	if ($item -eq $null)
	{
		throw "No item specified"
	}

    if ($script:itemLimitHit) { return }
    if ( -not (ItemHasRequiredProperties($item)) -or -not (ItemPropertiesMatchRequirements($item)) ) { return }
    if ( -not (ItemMatchesRecipientRequirements($item)) ) { return }
    if ( -not (ItemMatchesDateRequirement($item)) ) { return }

    LogVerbose "Processing item: $($item.Subject)"
    $script:itemsMatched++

    if ($MaximumNumberOfItemsToProcess -gt 0)
    {
        if ($script:itemsMatched -gt $MaximumNumberOfItemsToProcess)
        {
            # We've processed maximum number of items, so turn -WhatIf on            
            $script:WhatIf = $true

            if ($StopAfterMaximumNumberOfItemsProcessed)
            {
                Log "$MaximumNumberOfItemsToProcess items processed, halting further action" Green
                $script:itemLimitHit = $true
                return
            }
            else
            {
                $MaximumNumberOfItemsToProcess = 0
                Log "$MaximumNumberOfItemsToProcess items processed, -WhatIf enabled for further processing" Green
            }
        }
    }

    if ($ListMatches)
    {
        $item
    }

    $madeChanges = $false

    # Check calendar invitation processing
    if ($item.ItemClass.Equals("IPM.Schedule.Meeting.Request"))
    {
        if ($AcceptCalendarInvite)
        {
            if ($DeclineCalendarInviteIfConflict -and (AppointmentHasConflict $item))
            {
                Log "Declining calendar invitation due to conflict"
                if (!$WhatIf) { $item.Decline($true) | out-null }
            }
            else
            {
                Log "Accepting calendar invitation"
                $acceptedMeeting = $null
                if (!$WhatIf) { $acceptedMeeting = $item.Accept($true) }

                # We only process the subject when the meeting is accepted, as when declined there is nothing to update...
                if ($acceptedMeeting -ne $null)
                {
                    $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer,
                        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
                    $meeting = ThrottledItemBind $acceptedMeeting.Appointment.Id $propSet
                    $global:debugMeeting = $meeting

                    if ($meeting -ne $null)
                    {
                        $meetingSubject = $meeting.Subject
                        if ($DeleteSubject)
                        {
                            LogVerbose "Deleting meeting subject"
                            $meetingSubject = ""
                        }

                        if ($AddOrganizerToSubject)
                        {
                            LogVerbose "Adding organizer to subject"
                            if (![String]::IsNullOrEmpty($meetingSubject)) { $meetingSubject = "$meetingSubject - " }
                            $meetingSubject = "$meetingSubject$($meeting.Organizer.Name)"
                        }

                        if ($meetingSubject -ne $meeting.Subject)
                        {
                            # Update the appointment with the changes
                            Log "Updating meeting subject to: $meetingSubject"
                            $meeting.Subject = $meetingSubject
                            ThrottledItemUpdate $meeting | out-null
                        }
                    }
                    else
                    {
                        Log "Failed to post process meeting as unable to bind to appointment" Red
                    }
                }
            }
            $madeChanges = $true
        }
        elseif ($DeclineCalendarInvite)
        {
            Log "Declining calendar invitation"
            if (!$WhatIf) { $item.Decline($true) | out-null }
            $madeChanges = $true
        }
    }

    # Check for Resend
    if ($Resend)
    {
        if (!(ResendItem $item))
        {
            if ($Delete)
            {
                Log "Not deleting item as resend failed" Red
                return
            }
        }
    }

    # Check for delete
    if ($Delete)
    {
        if (-not $WhatIf)
        {
            if (-not $Resend)
            {
                [void]$script:itemsToDelete.Add($item.Id)
                Log "`"$($item.Subject)`" added to list of items to be deleted" Gray
            }
            else
            {
                # If we are resending, we delete the message immediately instead of batching (so that in the event of an issue, script can rerun and pick up where it left off)
                if (ThrottledItemDelete $item)
                {
                    Log "`"$($item.Subject)`" deleted" Gray
                    $script:itemsDeleted++
                }
                else
                {
                    Log "FAILED to delete: $($item.Subject)" Red
                }
            }
        }
        else
        {
            Log "`"$($item.Subject)`" would be deleted" Gray
            if ($MaximumNumberOfItemsToProcess -lt 1)
            {
                $script:itemsDeleted++
            }
        }

        return # If Delete is specified, any other parameter is irrelevant
    }

    if ( DeleteItemProperties $item ) { $madeChanges = $True }
    if ( AddItemProperties $item ) { $madeChanges = $True }

    if ($DeleteContactPhoto) { $madeChanges = DeleteContactPhoto $item }
    if (MarkWhetherRead $item) { $madeChanges = $true }

    if ($madeChanges)
    {
        $script:itemsAffected++
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

    if ($results -ne $null)
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
                if ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed" -or $results[$i].ErrorCode -eq "ErrorInvalidOperation")
                {
                    # This is a permanent error, so we remove the item from the list
                    $Items.Remove($requestedItems[$i])
                    $script:itemsWithError++
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
    # Send request to move/copy items, allowing for throttling (which in this case is likely to manifest as time-out errors)
    param (
        $ItemsToDelete,
        $BatchSize = 500
    )

    if ($script:MaxBatchSize -gt 0)
    {
        # If we've had to reduce the batch size previously, we'll start with the last size that was successful
        $BatchSize = $script:MaxBatchSize
    }

    $progressActivity = "Deleting items"
    
    $finished = $false
    $totalItems = $ItemsToDelete.Count
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    LogVerbose "Deleting $totalItems item(s)"
    if ( $totalItems -gt 10000 )
    {
        if ( $script:throttlingDelay -lt 1000 )
        {
            $script:throttlingDelay = 1000
            LogVerbose "Large number of items will be processed, so throttling delay set to 1000ms"
        }
    }
    $consecutiveErrors = 0

    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems
    if ($HardDelete)
    {
        $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
    }

    while ( !$finished )
    {
        $deleteIds = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
        
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
        ApplyEWSOAuthCredentials
        try
        {
            LogVerbose "Sending batch request to delete $($deleteIds.Count) items ($($ItemsToDelete.Count) remaining)"
			$results = $script:service.DeleteItems( $deleteIds, $deleteMode, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
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
        }

        $script:itemsDeleted += $ItemsToDelete.Count
        RemoveProcessedItemsFromList $deleteIds $results $ItemsToDelete
        $script:itemsDeleted -= $ItemsToDelete.Count

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
}

Function InitialiseItemPropertySet()
{
    if ($script:RequiredPropSet -ne $null)
    {
        return $script:RequiredPropSet
    }
    if ($LoadAllItemProperties)
    {
        $script:RequiredPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    }
    else
    {
        $script:RequiredPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,[Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
            [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead,[Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,[Microsoft.Exchange.WebServices.Data.ContactSchema]::HasPicture)
    }

    if ($CreatedAfter -or $CreatedBefore)
    {
        $script:RequiredPropSet.Add($script:PR_CREATION_TIME)
    }

    if ($script:deleteItemPropsEws -ne $null) # We retrieve any properties that we want to delete
    {
        foreach ($deleteProperty in $script:deleteItemPropsEws)
        {
            $script:RequiredPropSet.Add($deleteProperty)
        }
    }
    if ($script:propertiesMustExistEws -ne $null) # We retrieve any properties that must exist (so that we can tell if they exist!)
    {
        foreach ($requiredProperty in $script:propertiesMustExistEws)
        {
            if (-not ($script:RequiredPropSet.Contains($requiredProperty)) )
            {
                $script:RequiredPropSet.Add($requiredProperty)
            }
        }
    }
    if ($script:propertiesMustMatchEws -ne $null) # We retrieve any properties that we need to check the value of
    {
        foreach ($propMustMatch in $script:propertiesMustMatchEws.Keys)
        {
            #LogVerbose "$propMustMatch"
            if (-not ($script:RequiredPropSet.Contains($propMustMatch)) )
            {
                $script:RequiredPropSet.Add($propMustMatch)
            }
        }
    }

    if ($DeclineCalendarInviteIfConflict)
    {
        # To check for conflicts, we need start and end time of the invitation
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ParentFolderId)
    }

    InitRecipientMatchInfo

    if ($Resend)
    {
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::MimeContent)
        $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Body)
        if ($ResendUpdatePrependTextFields)
        {
            $script:RequiredPropSet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent)
        }
    }
}

Function ProcessFolder()
{
	# Process this folder

    $Folder = $args[0]
	if ($Folder -eq $null)
	{
		throw "No folder specified"
	}
	
    Log "Processing folder: $($Folder.DisplayName)" Gray
    $progressActivity = "$($Folder.DisplayName):"

	# Process any subfolders
	if ($ProcessSubFolders)
	{
		if ($Folder.ChildFolderCount -gt 0)
		{
            # We read the list of all folders first, so that we have the complete list before any processing
            $subfolders = @()
            $moreFolders = $True
			$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
            while ($moreFolders)
            {
                ApplyEWSOAuthCredentials
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
				    ProcessFolder $subFolder
			    }
            }
		}
	}

    # Now process all items in this folder
	$Offset=0
	$PageSize=1000
	$MoreItems=$true

    # We create a list of all the items we are going to process (this means we don't have to allow for any delete actions, etc.)
    $itemsToProcess = @()
    $script:itemsToDelete = New-Object System.Collections.ArrayList # Any items to delete we process in batch (we can't easily do this for updates)
    $i = 0

    LogVerbose "Building list of items"
    $filters = @()
    if ($MatchContactAddresses)
    {
        # Add filter for contact address matching (a contact address can be in one of three properties)
        $contactFilters = @()
        foreach ($contactAddress in $MatchContactAddresses)
        {
            LogVerbose "Adding SMTP address search: $smtpAddress"
            $contactFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress1, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
            $contactFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress2, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
            $contactFilters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress3, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)

        }
        if ( $contactFilters.Count -gt 0 )
        {
            $contactFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
            foreach ($filter in $contactFilters)
            {
                $contactFilter.Add($filter)
            }
            $filters += $contactFilter
        }
    }

    # Add filter(s) for creation time
    if ( $CreatedAfter )
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan($script:PR_CREATION_TIME, $CreatedAfter)
    }
    if ( $CreatedBefore )
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan($script:PR_CREATION_TIME, $CreatedBefore)
    }

    if (![String]::IsNullOrEmpty($SearchFilter))
    {
        LogVerbose "Search query being applied: $SearchFilter"
    }
    elseif ( $filters.Count -gt 0 )
    {
        # Create the search filter
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        foreach ($filter in $filters)
        {
            $SearchFilter.Add($filter)
        }
    }    

    Write-Progress -Activity "$progressActivity reading items" -Status "0 items found" -PercentComplete -1
	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
        # As some properties (e.g. recipients) cannot be retrieved using FindItem, we only retrieve item Ids here and perform a GetItem later to get the properties
        $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly

        if ($AssociatedItems)
        {
            $View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        }

        try
        {
            if ($MatchContactAddresses -or ![String]::IsNullOrEmpty($SearchFilter))
            {
                # We have a search filter, so need to apply this
                $FindResults=$Folder.FindItems($SearchFilter, $View)
            }
            else
            {
                # No search filter, we want everything
		        $FindResults=$Folder.FindItems($View)
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
            $itemsToProcess += $FindResults.Items
		    $MoreItems=$FindResults.MoreAvailable
            if ($MoreItems)
            {
                LogVerbose "$($itemsToProcess.Count) items found so far, more available"
            }
		    $Offset+=$PageSize
        }
        Write-Progress -Activity "$progressActivity reading items" -Status "$($itemsToProcess.Count) item(s) found"
	}
    Write-Progress -Activity "$progressActivity reading items" -Status "Complete" -Completed
    
    # Now go through each item and process it
    Write-Progress -Activity "$progressActivity processing items" -Status "0 items processed" -PercentComplete 0
    $i = 0

    if ($LoadItemsIndividually)
    {
        # Send a GetItem request for each item
        ForEach ($item in $itemsToProcess)
        {
            ProcessItem (ThrottledItemBind($item.Id))
            $i++
            if ($i%10 -eq 0)
            {
                Write-Progress -Activity "$progressActivity processing items" -Status "$i items processed" -PercentComplete (($i/$itemsToProcess.Count)*100)
            }
            if ($script:itemLimitHit) { break }
        }
    }
    else
    {
        # We send GetItem request for multiple items at a time

        $itemIds = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
        ForEach ($item in $itemsToProcess)
        {
            $itemIds.Add($item.Id)
            $i++
            if ($itemIds.Count -ge $GetItemBatchSize)
            {
                # We have a full batch, so retrieve these items and process
                ApplyEWSOAuthCredentials
                $fullItems = $script:service.BindToItems( $itemIds, $script:RequiredPropSet )
                foreach ($fullItem in $fullItems)
                {
                    ProcessItem $fullItem.Item
                    if ($script:itemLimitHit) { break }
                }
                $itemIds = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
                Write-Progress -Activity "$progressActivity processing items" -Status "$i items processed" -PercentComplete (($i/$itemsToProcess.Count)*100)

            }
        }
        if ($itemIds.Count -gt 0 -and -not $script:itemLimitHit)
        {
            # Process the remaining items
            $fullItems = $script:service.BindToItems( $itemIds, $script:RequiredPropSet )
            
            foreach ($fullItem in $fullItems)
            {
                ProcessItem $fullItem.Item
                if ($script:itemLimitHit) { break }
            }
            $itemIds = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
            Write-Progress -Activity "$progressActivity processing items" -Status "$i items processed" -PercentComplete (($i/$itemsToProcess.Count)*100)
        }
    }

    Write-Progress -Activity "$progressActivity processing items" -Status "Complete" -Completed
    Log "Completed processing folder $($Folder.DisplayName)" Gray

    if ($script:itemsToDelete.Count -gt 0)
    {
        ThrottledBatchDelete $script:itemsToDelete
    }
}

function ProcessMailbox()
{
    # Process the mailbox

    Log "Processing mailbox $Mailbox" Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
	}

    $script:throttlingDelay = 0
    $script:itemsProcessedCount = 0

    # Bind to root folder
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    if ($PublicFolders)
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot, $mbx )
    }
    else
    {
        if ($Archive)
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx )
            LogVerbose "Attempting to bind to archive message root folder ($($Mailbox))"
        }
        else
        {
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
            LogVerbose "Attempting to bind to message root folder ($($Mailbox))"
        }
    }

    $rootFolder = $Null
    $rootFolder = ThrottledFolderBind($folderId)

    # Check we could get the root folder (if not, there's nothing else we can do)
    if ($rootFolder -eq $Null)
    {
        if ($Impersonate)
        {
            Log "Unable to bind to root folder.  Please check whether ApplicationImpersonation permissions have been granted to authenticating account." Red
        }
        else
        {
            Log "Unable to bind to root folder.  No further processing possible." Red
        }
        Log "If accessing a mailbox that has MFA enabled, you must use OAuth" Yellow
        return
    }

    $script:itemsAffected = 0
    $script:itemsDeleted = 0
    $script:itemsResent = 0
    $script:itemsMatched = 0
    $script:itemsWithError = 0
    $script:itemLimitHit = $false

    # FolderPath can support arrays (list of folders)
    if ([String]::IsNullOrEmpty($FolderPath))
    {
        # No folder path specified, so process root
        ProcessFolder $rootFolder
    }
    else
    {
        foreach ($fPath in $FolderPath)
        {
            if ($fPath.ToLower().StartsWith("wellknownfoldername."))
            {
                # Well known folder specified (could be different name depending on language, so we bind to it using WellKnownFolderName enumeration)
                $wkf = $fPath.SubString(20)
                $restOfPath = ""
                $restOfPathStart = $wkf.IndexOf("\")
                if ($restOfPathStart -gt -1)
                {
                    $restOfPath = $wkf.SubString($restOfPathStart+1)
                    $wkf = $wkf.SubString(0, $restOfPathStart)
                }
                Write-Verbose "Attempting to bind to well known folder: $wkf"
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
                $Folder = ThrottledFolderBind($folderId)
                if ($folder -and ![String]::IsNullOrEmpty($restOfPath))
                {
                    $Folder = GetFolder($Folder, $restOfPath, $false)
                }
            }
            else
            {
	            $Folder = GetFolder($rootFolder, $fPath, $false)
            }
	        if (!$Folder)
	        {
		        Write-Host "Failed to find folder $fPath" -ForegroundColor Red
	        }
            else
            {
                ProcessFolder $Folder
            }
        }
    }
    if ($WhatIf -and $MaximumNumberOfItemsToProcess -eq 0)
    {
        Log "$($Mailbox): $($script:itemsMatched) item(s) matched specified criteria"
        Log "$($Mailbox): $($script:itemsAffected) item(s) would be changed (but -WhatIf was specified so no action was taken)"
        if ($Resend)
        {
            Log "$($Mailbox): $($script:itemsResent) item(s) would be resent (but -WhatIf was specified so no action was taken)"
        }
        Log "$($Mailbox): $($script:itemsDeleted) item(s) would be deleted (but -WhatIf was specified so no action was taken)"
    }
    else
    {
        Log "$($Mailbox): $($script:itemsMatched) item(s) matched specified criteria"
        Log "$($Mailbox): $($script:itemsAffected) item(s) were changed"
        if ($Resend)
        {
            Log "$($Mailbox): $($script:itemsResent) item(s) were resent"
        }
        Log "$($Mailbox): $($script:itemsDeleted) item(s) were deleted"
    }
    if ($script:itemsWithError -gt 0)
    {
        Log "$($Mailbox): $($script:itemsWithError) item(s) FAILED TO PROCESS" Red
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
    Write-Host "The API can be downloaded from the Microsoft Download Centre: http://www.microsoft.com/en-us/search/Results.aspx?q=exchange%20web%20services%20managed%20api&form=DLC"
    Write-Host "Use the latest version available"
	Exit
}

$script:PR_CREATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3007, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime) 

Write-Host ""
CreatePropLists
InitialiseItemPropertySet

# Check whether we have a CSV file as input...
If ( $(Test-Path $Mailbox) )
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

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}

if ($script:logFileStreamWriter)
{
    $script:logFileStreamWriter.Close()
    $script:logFileStreamWriter.Dispose()
}
