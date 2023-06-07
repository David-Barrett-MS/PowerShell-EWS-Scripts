#
# Merge-MailboxFolder.ps1
#
# By David Barrett, Microsoft Ltd. 2015-2022. Use at your own risk.  No warranties are given.
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

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given AQS filter will be moved `r`n(see https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-an-aqs-search-by-using-ews-in-exchange )")]
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

    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EwsId (not path)")]
    [switch]$ByFolderId,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the folders in MergeFolderList are identified by EntryId (not path)")]
    [switch]$ByEntryId,

    [Parameter(Mandatory=$False,HelpMessage="When specified, subfolders will also be processed")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, all items in subfolders of source will be moved to specified target folder (hierarchy will NOT be maintained)")]
    [alias("MergeSubfolders")]
    [switch]$CombineSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="When specified, if the target folder doesn't exist, then it will be created (if possible)")]
    [switch]$CreateTargetFolder,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the source mailbox being accessed will be the archive mailbox")]
    [switch]$SourceArchive,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the target mailbox being accessed will be the archive mailbox")]
    [switch]$TargetArchive,

    [Parameter(Mandatory=$False,HelpMessage="When specified, hidden (associated) items of the folder are processed (normal items are ignored)")]
    [switch]$AssociatedItems,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the source folder will be deleted after the move (can't be used with -Copy)")]
    [switch]$Delete,

    [Parameter(Mandatory=$False,HelpMessage="When specified, items are copied rather than moved (can't be used with -Delete)")]
    [switch]$Copy,

    [Parameter(Mandatory=$False,HelpMessage="If specified, no moves will be performed (but actions that would be taken will be logged)")]
    [switch]$WhatIf,

    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,

    [Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts).")]
    [switch]$OAuth,

    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
    [string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id in which the application is registered.  If missing, application is assumed to be multi-tenant and the common log-in URL will be used.")]
    [string]$OAuthTenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$OAuthSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.  Please note that certificate auth requires the MSAL dll to be available.")]
    $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
    [switch]$Impersonate,

    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, and -Office365 not specified, then autodiscover is used)")]
    [string]$EwsUrl,

    [Parameter(Mandatory=$False,HelpMessage="If specified, requests are directed to Office 365 endpoint (this overrides -EwsUrl)")]
    [switch]$Office365,

    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]
    [string]$EWSManagedApiPath = "",
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]
    [switch]$IgnoreSSLCertificate,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]
    [switch]$AllowInsecureRedirection,
	
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="Enable verbose log file.  Verbose logging is written to the log whether -Verbose is enabled or not.")]	
    [switch]$VerboseLogFile,

    [Parameter(Mandatory=$False,HelpMessage="Enable debug log file.  Debug logging is written to the log whether -Debug is enabled or not.")]	
    [switch]$DebugLogFile,

    [Parameter(Mandatory=$False,HelpMessage="If selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled)")]
    [switch]$FastFileLogging,

    [Parameter(Mandatory=$False,HelpMessage="Enables token debugging (for development purposes - do not use)")]	
    [switch]$DebugTokenRenewal,

    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]
    [string]$TraceFile,

    [Parameter(Mandatory=$False,HelpMessage="Batch size (number of items batched into one EWS request) - this will be decreased if throttling is detected")]	
    [int]$BatchSize = 100
)
$script:ScriptVersion = "1.3.0"
$scriptStartTime = [DateTime]::Now

# Define our functions

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

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    $logInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
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
        try
        {
            $dll = Get-ChildItem $dllName -ErrorAction SilentlyContinue
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
        $script:LastError = $Error[0] # We do this to suppress any errors encountered during the search above

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

    $cca1 = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca2 = $cca1.WithCertificate($OAuthCertificate)
    $cca3 = $cca2.WithTenantId($OAuthTenantId)
    $cca = $cca3.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    $script:Impersonate = $true

    if ($DebugTokenRenewal)
    {
        $global:certAuthResult = $authResult
    }
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$OAuthRedirectUri&response_mode=query&scope=openid%20profile%20email%20offline_access%20https://outlook.office365.com/.default"
    Write-Host "Please complete log-in via the web browser, and then paste the redirect URL (including auth code) here to continue" -ForegroundColor Green
    Start-Process $authUrl

    $authcode = Read-Host "Auth code"
    $codeStart = $authcode.IndexOf("?code=")
    if ($codeStart -gt 0)
    {
        $authcode = $authcode.Substring($codeStart+6)
    }
    $codeEnd = $authcode.IndexOf("&session_state=")
    if ($codeEnd -gt 0)
    {
        $script:AuthCode = $authcode.Substring(0, $codeEnd)
    }
    Write-Verbose "Using auth code: $authcode"

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$script:AuthCode;redirect_uri=$OAuthRedirectUri}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
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
      "client_secret" = "$OAuthSecretKey";
      "scope"         = "https://outlook.office365.com/.default"
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token: $Error" -ForegroundColor Red
        exit # Failed to obtain a token
    }
    $script:Impersonate = $true
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

    if ($OAuthCertificate -ne $null)
    {
        GetTokenWithCertificate
    }
    elseif (![String]::IsNullOrEmpty($OAuthSecretKey))
    {
        GetTokenWithKey
    }
    else
    {
        if ($RenewToken)
        {
            RenewOAuthToken
        }
        else
        {
            GetTokenViaCode
        }
    }

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

$script:oAuthDebugStop = $false
function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $script:services -eq $null ) { return }

    
    if ($DebugTokenRenewal -and $script:oauthToken)
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
        }
        else
        {
            $script:oAuthDebugStop = $true
        }
    }
    
    if ($OAuthCertificate -ne $null)
    {
        if ( [DateTime]::UtcNow -ge $script:oauthToken.ExpiresOn.UtcDateTime) { return }
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

    if ($DebugTokenRenewal)
    {
        $global:oAuthTokenDebug = $script:oauthToken
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
    if ( $script:Tracer -eq $null -or [String]::IsNullOrEmpty($script:Tracer.LastResponse))
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
        ApplyEWSOauthCredentials
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
            ApplyEWSOauthCredentials
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
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    while ( !$finished )
    {
	    $script:moveIds = [Activator]::CreateInstance($genericItemIdList)
        $script:deleteIds = [Activator]::CreateInstance($genericItemIdList) # This is used to check that items were deleted once moved (only happens when moving between public folders)

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
                    if ($WhatIf)
                    {
                        Log "Would move/copy $($ItemsToMove[$i])"
                    }
                    else
                    {
                        if (!$Copy -and $script:publicFolders)
                        {
                            $deleteIds.Add($ItemsToMove[$i])
                        }
                        LogVerbose "Added to batch: $($ItemsToMove[$i])"
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
                LogVerbose "Batch request completed"
            }
        }
        catch
        {
            if ( Throttled )
            {
                # We've been throttled, so we reduce batch size (to a minimum size of 50) and try again
                if ($script:currentBatchSize -gt 50)
                {
                    DecreaseBatchSize
                }
                else
                {
                    #$finished = $true
                }
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

        if (!$WhatIf)
        {
            RemoveProcessedItemsFromList $moveIds $results $false $ItemsToMove

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
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToMove.Count -eq 0)
        {
            $finished = $True
            Write-Progress -Activity $progressActivity -Status "100% complete" -Completed
        }
    }

    if ($script:deleteIds.Count -gt 0)
    {
        # We have a list of items to delete (i.e. Move succeeded, but we are processing public folders so need to ensure that the source item no longer exists)
        ThrottledBatchDelete $script:deleteIds -SuppressNotFoundErrors $true
    }

    # Restore the throttling delay (in case we changed it)
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
			$results = $script:sourceService.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
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
                        Log "[IsFolderExcluded]Ignoring search folder: $folderPath"
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

    if ($ExcludeFolderList)
    {
        LogVerbose "Checking for exclusions: $($ExcludeFolderList -join ',')"
        $rootFolderName = $script:sourceMailboxRoot.DisplayName.ToLower()
        ForEach ($excludedFolder in $ExcludeFolderList)
        {
            LogDebug "[IsFolderExcluded]Comparing $($folderPath.ToLower()) to $($excludedFolder.ToLower())"
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
                    Log "[IsFolderExcluded]Excluded folder being skipped: $folderPath"
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
	
    if ( $SourceFolderObject -eq $null )
    {
        Log "[MoveItems]Source folder is null, cannot move items" Red
        return
    }	
    if ( $TargetFolderObject -eq $null )
    {
        Log "[MoveItems]Target folder is null, cannot move items" Red
        return
    }	
	if ( $SourceFolderObject.Id -eq $TargetFolderObject.Id )
	{
		Log "[MoveItems]Cannot move or copy from/to the same folder (source folder Id and target folder Id are the same)" Red
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
	Log "[MoveItems]$actioning from $($SourceMailbox):$(GetFolderPath($SourceFolderObject)) to $($TargetMailbox):$(GetFolderPath($TargetFolderObject))" White
	
	# Set parameters - we will process in batches of 500 for the FindItems call
	$Offset = 0
	$PageSize = 1000 # We're only querying Ids, so 1000 items at a time is reasonable
	$MoreItems = $true
    $moveCountSuccess = 0
    $moveCountFail = 0

    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time)
    $itemsToMove = New-Object System.Collections.ArrayList
    $i = 0
	
    $progressActivity = "Reading items in folder $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
    LogVerbose "[MoveItems]Building list of items to $($action.ToLower())"
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
        LogVerbose "[MoveItems]Search filters applied: $($searchFilters.Count)"
        $itemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, $searchFilters)
    }
    elseif ($SearchFilter)
    {
        $itemSearchFilter = $SearchFilter
        LogVerbose "[MoveItems]Search query being applied: $itemSearchFilter"
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
                Log "[MoveItems]Error when querying items: $($Error[0])" Red
                $MoreItems = $false
            }
        }
		
        if ($FindResults)
        {
		    ForEach ($item in $FindResults.Items)
		    {
                $skip = $False
                if ($IncludedMessageClasses -ne $null)
                {
                    # Check if this is an included message class
                    $skip = $true
                    foreach ($includedMessageClass in $IncludedMessageClasses)
                    {
                        if ($item.ItemClass -like $includedMessageClass)
                        {
                            LogVerbose "[MoveItems]Included message class $($item.ItemClass)"
                            $skip = $false
                            break
                        }
                    }
                }
                else
                {
                    if ($ExcludedMessageClasses -ne $Null)
                    {
                        # Check if this is an excluded message class
                        foreach ($excludedMessageClass in $ExcludedMessageClasses)
                        {
                            if ($item.ItemClass -like $excludedMessageClass)
                            {
                                LogVerbose "[MoveItems]Skipping item with message class $($item.ItemClass)"
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
                LogVerbose "[MoveItems]$($itemsToMove.Count) items read so far (out of $($SourceFolderObject.TotalCount))"
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
        Log "$($itemsToMove.Count) items found; attempting to $($action.ToLower())" Green
        ThrottledBatchMove $itemsToMove $TargetFolderObject.Id $Copy

        # Add a check for the number of items left in the folder (we expect it to be zero)
        $SourceFolderObject = ThrottledFolderBind $SourceFolderObject.Id $null $script:sourceService
        Log "[MoveItems]$($SourceMailbox):$(GetFolderPath($SourceFolderObject)) processed, now contains $($SourceFolderObject.TotalCount) items(s)" White
    }
    else
    {
        Log "No matching items were found" Green
    }

	# Now process any subfolders
	if ($ProcessSubFolders)
	{
		if ($SourceFolderObject.ChildFolderCount -gt 0)
		{
            LogVerbose "[MoveItems]Processing subfolders of $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
			$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
            $FolderView.PropertySet = $script:requiredFolderProperties
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
                            while ($FindFolderResults -eq $null -and $attempts -lt 3)
                            {
	                            $FindFolderResults = $TargetFolderObject.FindFolders($Filter, $FolderView)
                                $attempts++
                                if ($FindFolderResults -eq $null)
                                {
                                    if (!Throttled)
                                    {
                                        $attempts = 10       
                                    }

                                }
                            }
                        }
                        catch {}
                        if ($FindFolderResults -eq $null)
                        {
                            if ($WhatIf -and $CreateTargetFolder)
                            {
                                Log "Target folder not created due to -WhatIf: $($SourceSubFolderObject.DisplayName)"
                                $TargetFolderObject = New-Object PsObject
                                $TargetFolderObject | Add-Member NoteProperty DisplayName $SourceSubFolderObject.DisplayName
                            }
                            else
                            {
                                Log "[MoveItems]FAILED TO LOCATE TARGET FOLDER: $($SourceSubFolderObject.DisplayName)" Red
                                $TargetSubFolderObject = $null
                            }
                        }
                        elseif ($FindFolderResults.TotalCount -eq 0)
				        {
                            LogVerbose "[MoveItems]Creating target folder $($SourceSubFolderObject.DisplayName)"
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
					            $TargetSubFolderObject.Save($TargetFolderObject.Id)
                            }
                            catch {}
                            if ( $(ErrorReported "MoveItems") )
                            {
                                Log "[MoveItems]FAILED TO CREATE TARGET FOLDER: $($SourceSubFolderObject.DisplayName)"
                                $TargetSubFolderObject = $null
                            }
				        }
				        else
				        {
                            LogVerbose "[MoveItems]Target folder already exists"
					        $TargetSubFolderObject = $FindFolderResults.Folders[0]
				        }
                        if ($TargetSubFolderObject -ne $null)
                        {
				            MoveItems $SourceSubFolderObject $TargetSubFolderObject
                        }
                    }
                }
                else
                {
                    LogVerbose "[MoveItems]Folder $(GetFolderPath($SourceSubFolderObject)) on excluded list"
                }
			}
		}
        else
        {
            LogVerbose "[MoveItems]No subfolders found: $($SourceMailbox):$(GetFolderPath($SourceFolderObject))"
        }
	}

    # If delete parameter is set, check if the source folder is now empty (and if so, delete it)
    if ($Delete)
    {
	    $SourceFolderObject.Load()
	    if (($SourceFolderObject.TotalCount -eq 0) -And ($SourceFolderObject.ChildFolderCount -eq 0))
	    {
		    # Folder is empty, so can be safely deleted
		    try
		    {
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
        $RootFolder = ThrottledFolderBind $folderId $null $RootFolder.Service
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
        $ewsId = $script:service.ConvertId($id, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
    }
    catch {}
    LogVerbose "EWS Id: $($ewsId.UniqueId)"
    return $ewsId.UniqueId
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

function ProcessMailbox()
{
    # Process the mailbox

    $script:publicFolders = $false

    Write-Host "Processing mailbox $SourceMailbox" -ForegroundColor Gray

    if ( !([String]::IsNullOrEmpty($script:originalLogFile)) )
    {
        $LogFile = $script:originalLogFile.Replace("%mailbox%", $SourceMailbox)
    }

	$script:sourceService = CreateService($SourceMailbox)
	if ($script:sourceService -eq $Null)
	{
		Write-Host "Failed to connect to source mailbox" -ForegroundColor Red
        return
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

    if ( $script:sourceMailboxRoot -eq $null )
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
    if ( $script:targetMailboxRoot -eq $null )
    {
        Write-Host "Failed to open target message store ($TargetMailbox)" -ForegroundColor Red
        return
    }

    if ($MergeFolderList -eq $Null)
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
                if ($TargetFolderObject -eq $null -and $WhatIf)
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
    $sourceMailbox
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
	    Write-Host "Source mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red
	    Exit
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
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}
  
# Check whether parameters make sense
if ($Delete -and $Copy)
{
    Write-Host "Cannot -Delete and -Copy, please use only one of these switches and try again." -ForegroundColor Red
    exit
}

if ($MergeFolderList -eq $Null)
{
    # No folder list, this is a request to move the entire mailbox
    # Check -ProcessSubfolders and -CreateTargetFolder is set, otherwise we fail now (can't move a mailbox without processing subfolders!)
    if (!$ProcessSubfolders)
    {
        Write-Host "Mailbox merge requested, but subfolder processing not specified.  Please retry using -ProcessSubfolders switch." -ForegroundColor Red
    }
    if (!$CreateTargetFolder)
    {
        Write-Host "Mailbox merge requested, but folder creation not allowed.  Please retry using -CreateTargetFolder switch." -ForegroundColor Red
        exit
    }
    if (!$ProcessSubfolders) { exit }
}

# Set up script variables.  We set them here so that we can modify them depending upon what we need (some parameters mean we need to pull more properties back, and we add these as necessary)
$script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PR_FOLDER_TYPE)
$script:requiredItemProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::Subject)

Write-Host ""

# Check whether we have a CSV file as input...
$FileExists = Test-Path $SourceMailbox
If ( $FileExists )
{
	# We have a CSV to process
    LogVerbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $SourceMailbox -Header "PrimarySmtpAddress"
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

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}


Log "Script finished in $([DateTime]::Now.SubTract($scriptStartTime).ToString())" Green
if ($script:logFileStreamWriter)
{
    $script:logFileStreamWriter.Close()
    $script:logFileStreamWriter.Dispose()
}