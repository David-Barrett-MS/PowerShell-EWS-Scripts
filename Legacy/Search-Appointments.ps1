#
# Search-Appointments.ps1
#
# By David Barrett, Microsoft Ltd. 2015 - 2023. Use at your own risk.  No warranties are given.
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
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,
		
    [Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox calendar folder is assumed")]
    [string]$FolderPath,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder")]
    [switch]$PublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="Subject of the appointment(s) being searched")]
    [string]$Subject,	

    [Parameter(Mandatory=$False,HelpMessage="Organizer of the appointment(s) being searched")]
    [string]$Organizer,

    [Parameter(Mandatory=$False,HelpMessage="Location of the appointment(s) being searched")]
    [string]$Location,

    [Parameter(Mandatory=$False,HelpMessage="Category of the appointment(s) being searched")]
    [string]$Category,

    [Parameter(Mandatory=$False,HelpMessage="Start date for the appointment(s) must be after this date")]
    [string]$StartsAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="Start date for the appointment(s) must be before this date")]
    [string]$StartsBefore,
	
    [Parameter(Mandatory=$False,HelpMessage="End date for the appointment(s) must be after this date")]
    [string]$EndsAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="End date for the appointment(s) must be before this date")]
    [string]$EndsBefore,
	
    [Parameter(Mandatory=$False,HelpMessage="Only appointments created before the given date will be returned")]
    [string]$CreatedBefore,
    
    [Parameter(Mandatory=$False,HelpMessage="Only appointments created after the given date will be returned")]
    [string]$CreatedAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="Only recurring appointments with a last occurrence date before the given date will be returned")]
    [string]$LastOccurrenceBefore,
    
    [Parameter(Mandatory=$False,HelpMessage="Only recurring appointments with a last occurrence date after the given date will be returned")]
    [string]$LastOccurrenceAfter,
    
    [Parameter(Mandatory=$False,HelpMessage="If specified, only appointments with at least this number of exceptions will be returned.  Exceptions include both modified and deleted occurrences.")]
    [int]$HasExceptions = -1,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only appointments with at least this number of attachments will be returned (NOT IMPLEMENTED)")]
    [int]$HasAttachments = -1,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, only recurring appointments are returned")]
    [switch]$IsRecurring,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, only recurring appointments that have no end date are returned (must be used with -IsRecurring)")]
    [switch]$HasNoEndDate,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, only all-day appointments are returned")]
    [switch]$IsAllDay,

    [Parameter(Mandatory=$False,HelpMessage="If specified, any matched appointments will be deleted")]
    [switch]$Delete,
	
    [Parameter(Mandatory=$False,HelpMessage="The SendCancellationsMode option that will be used when deleting items.")]
    $SendCancellationsMode = "SendToNone",
	
    [Parameter(Mandatory=$False,HelpMessage="Folder path to which matching appointments will be moved")]
    [string]$MoveToFolder,
	
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
	
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
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

    [Parameter(Mandatory=$False,HelpMessage="CSV Export - appointments are exported to this file")]	
    [string]$ExportCSV = "",

    [Parameter(Mandatory=$False,HelpMessage="If this parameter is specified, exported times are in UTC")]	
    [switch]$ExportUTC,
	
    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
    [string]$TraceFile
)
$script:ScriptVersion = "1.2.1"
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

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        $RootFolder,
        [string]$FolderPath,
        [bool]$Create
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
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
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

function LoadItem( $item )
{
    # Load the item... We are doing this here so that we only load it once (i.e. we check if it needs loading first)
    if ($script:loadedItems -eq $Null)
    {
        # We only store an array of Ids - we don't need to store the whole item, as we will only encounter an item once
        $script:loadedItems = New-Object 'System.Collections.Generic.List[System.String]'
    }

    if ( $script:loadedItems.Contains( $item.Id.UniqueId ) )
    {
        # Already loaded this item
        LogVerbose "Item already loaded, not reloading"
        return
    }

    $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::AppointmentType)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::LastOccurrence)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Recurrence)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ModifiedOccurrences)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::DeletedOccurrences)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location)

    ApplyEWSOAuthCredentials
    $item.Load($propSet)
    $script:loadedItems.Add($item.Id.UniqueId)
}

function ExportTime($timeValue)
{
    if ($ExportUTC)
    {
        return $timeValue.ToUniversalTime()
    }
    return $timeValue
}

function LogToCSV()
{
    if (-not (Test-Path $ExportCSV))
    {
        # File doesn't exist, so write CSV headers
        if ( $script:CSVHeaders -eq $Null )
        {
            $script:CSVHeaders = """Mailbox"",""Subject"",""Sent"",""Received"",""Sender"",""Organizer"",""Start"",""End"",""IsAllDay"",""AppointmentType"""
        }
        $script:CSVHeaders | Out-File -FilePath $ExportCSV
    }

    $args | Out-File -FilePath $ExportCSV -Append
}

function ProcessItem( $item )
{
	# We have found an item, so this function handles any processing

    if (![String]::IsNullOrEmpty($ExportCSV))
    {
        LoadItem $item
	    LogToCSV "`"$Mailbox`",`"$($item.Subject)`",`"$(ExportTime($item.DateTimeSent))`",`"$(ExportTime($item.DateTimeReceived))`",`"$($item.Sender)`",`"$($item.Organizer)`",`"$((ExportTime($item.Start)))`",`"$((ExportTime($item.End)))`",`"$($item.IsAllDayEvent)`",`"$($item.AppointmentType.ToString())`""
    }

    if (!$Delete -and [String]::IsNullOrEmpty($MoveToFolder))
    {
        # No actions are specified, so we just log this appointment
        Log "$($item.Start) $($item.Subject)"
        return
    }

	# Add the item to our list of matches (for batch processing later)
    if ( $script:matches.ContainsKey($item.Id.UniqueId) )
    {
        Log "Item not added to match list as matching Id already present"
        if ($script:matches[$item.Id.UniqueId].Subject -ne $item.Subject)
        {
            Log "Subject of matching item: $($script:matches[$item.Id.UniqueId].Subject)" Red
        }
    }
    else
    {
        $script:matches.Add($item.Id.UniqueId, $item)
		LogVerbose "Item added to match list: $($item.Subject)"

    }
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

function SearchForAppointments($Folder)
{
    # Search for the appointment

    $startsBeforeDate = ParseDate $StartsBefore "StartsBefore"
    $startsAfterDate = ParseDate $StartsAfter "StartsAfter"
    $endsBeforeDate = ParseDate $EndsBefore "EndsBefore"
    $endsAfterDate = ParseDate $EndsAfter "EndsAfter"
    $createdBeforeDate = ParseDate $CreatedBefore "CreatedBefore"
    $createdAfterDate = ParseDate $CreatedAfter "CreatedAfter"
    $lastOccurrenceBeforeDate = ParseDate $LastOccurrenceBefore "LastOccurrenceBefore"
    $lastOccurrenceAfterDate = ParseDate $LastOccurrenceAfter "LastOccurrenceAfter"

    # Use FindItems as opposed to FindAppointments.  In this case, we process all items in the folder, and manually check if they meet our criteria
    
    $offset = 0
    $moreItems = $true
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, 0)
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
    $Propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent)
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::AppointmentType)
    

    # Set the search filter - this limits some of the results, not all the options can be filtered
    $filters = @()
    if ($createdBeforeDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $createdBeforeDate)
    }
    if ($createdAfterDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $createdBeforeDate)
    }
    if ($startsBeforeDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $startsBeforeDate)
    }
    if ($startsAfterDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $startsAfterDate)
    }
    if ($endsBeforeDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, $endsBeforeDate)
    }
    if ($endsAfterDate -ne $Null)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, $endsAfterDate)
    }
    if ($IsAllDay)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent, $true)
    }
    if ($HasAttachments -gt -1)
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments, $true)
    }
    if ($HasExceptions -gt -1)
    {
        # Only recurring appointments can have exceptions, so we add a filter for that and then check exceptions
        $PidLidIsRecurring = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8223, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidLidIsRecurring, $true)
    }
    if (![String]::IsNullOrEmpty($Location))
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location, $Location)
    }
    if (![String]::IsNullOrEmpty($Category))
    {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Categories, $Category)
    }

    $view.PropertySet = $propset
    $searchFilter = $Null
    if ( $filters.Count -gt 0 )
    {
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        foreach ($filter in $filters)
        {
            $searchFilter.Add($filter)
        }
    }

    # Now retrieve the matching items and process
    while ($moreItems)
    {
        # Get the next batch of items to process
        ApplyEWSOAuthCredentials
        if ( $searchFilter )
        {
            $results = $Folder.FindItems($searchFilter, $view)
        }
        else
        {
            $results = $Folder.FindItems($view)
        }
        $moreItems = $results.MoreAvailable
        $view.Offset = $results.NextPageOffset

        # Loop through each item and check if it matches criteria
        foreach ($item in $results)
        {
            LogVerbose "Checking item: $($item.Subject)"
            LogVerbose "Item Id: $($item.Id.UniqueId)"
            $match = $True

            # Check creation date
            if ( $createdBeforeDate -ne $Null )
            {
                
                if ($item.CreationTime -ge $createdBeforeDate) { $match = $False }
                LogVerbose "CreatedBefore test: match=$match"
            }

            if ( $match -and ($createdAfterDate -ne $Null) )
            {
                if ($item.CreationTime -le $createdAfterDate) { $match = $False }
                LogVerbose "CreatedAfter test: match=$match"
            }

            # Check for start date match
            if ( $match -and ($startsBeforeDate -ne $Null) )
            {
                if ($item.Start -ge $startsBeforeDate) { $match = $False }
                LogVerbose "StartsBefore test: match=$match"
            }

            if ( $match -and ($startsAfterDate -ne $Null) )
            {
                if ($item.Start -le $startsAfterDate) { $match = $False }
                LogVerbose "StartsAfter test: match=$match"
            }

            # Check for end date match
            if ( $match -and ($endsBeforeDate -ne $Null) )
            {
                if ($item.End -ge $endsBeforeDate) { $match = $False }
                LogVerbose "EndsBefore test: match=$match"
            }

            if ( $match -and ($endsAfterDate -ne $Null) )
            {
                if ($item.End -gt $endsAfterDate) { $match = $False }
                LogVerbose "EndsAfter test: match=$match"
            }

            # Check the number of attachments
            if ( $match -and ($HasAttachments -gt -1) )
            {
                # Not implemented yet
                LogVerbose "HasAttachments test (NOT IMPLEMENTED)"
            }

            # Check for subject match
            if ( $match -and (![String]::IsNullOrEmpty($Subject)) )
            {
                if ($item.Subject -notlike $Subject) { $match = $False }
                LogVerbose "Subject test: match=$match"
            }
            
            # Check for Organizer match
            if ( ![String]::IsNullOrEmpty($Organizer) )
            {
                if ($item.Organizer.Address -like ($Organizer)) { $match = $True } else { $mismatch = $True }
                LogVerbose "Organizer test: match=$match, mismatch=$mismatch"
            }

            # Check last occurrence (we can't do this in FindItems, and must request the property by binding to the item)
            if ( $match -and ($lastOccurrenceBeforeDate -or $lastOccurrenceAfterDate) )
            {
                LoadItem $item
                if ( $item.AppointmentType -ne [Microsoft.Exchange.WebServices.Data.AppointmentType]::RecurringMaster )
                {
                    $match = $False
                    LogVerbose "LastOccurrence test: match=$match"
                }
                else
                {
                    LogVerbose "Last occurrence: $($item.LastOccurrence.End)"
                    if ( $lastOccurrenceBeforeDate -ne $Null )
                    {
                        if ($item.LastOccurrence.End -ge $lastOccurrenceBeforeDate) { $match = $False }
                        LogVerbose "LastOccurrenceBefore test: match=$match"
                    }
                    if ( $lastOccurrenceAfterDate -ne $Null )
                    {
                        if ($item.LastOccurrence.End -le $lastOccurrenceAfterDate) { $match = $False }
                        LogVerbose "LastOccurrenceAfter test: match=$match"
                    }
                }
            }

            # Check if recurring
            if ( $match -and $IsRecurring )
            {
                if ( $item.AppointmentType -ne [Microsoft.Exchange.WebServices.Data.AppointmentType]::RecurringMaster ) { $match = $False }
                LogVerbose "IsRecurring test: match=$match"
            }
            
            # Check if appointment has exceptions
            if ( $match -and ($HasExceptions -gt -1) )
            {
                LoadItem $item
                if ( $item.AppointmentType -ne [Microsoft.Exchange.WebServices.Data.AppointmentType]::RecurringMaster ) { $match = $False }
                $exceptionsCount = $item.ModifiedOccurrences.Count + $item.DeletedOccurrences.Count
                LogVerbose "Appointment has $exceptionsCount exceptions"
                if ($exceptionsCount -lt $HasExceptions) { $match = $false }
                LogVerbose "HasExceptions test: match=$match"
            }
           
            # Check if all-day event
            if ( $match -and $IsAllDay )
            {
                if ( !$item.IsAllDayEvent ) { $match = $False }
                LogVerbose "IsAllDayEvent test: match=$match"
            }

            if ( $match )
            {
                ProcessItem $item
            }
        }
    }

    return

}

function ProcessBatches()
{
    # Take care of any batch processing (e.g. deleting), which is done this way for greater efficiency (and performance)

	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$baseList = [System.Collections.Generic.List``1]
	$genericItemIdList = $baseList.MakeGenericType(@($itemIdType))
    $batchSize = 100

    if ( ![String]::IsNullOrEmpty($MoveToFolder) )
    {
        # Move the items to the specified folder
        $targetFolder = GetFolder($MoveToFolder)
        if ($targetFolder -eq $Null)
            { return }

	    $moveIds = [Activator]::CreateInstance($genericItemIdList)
        $i = $script:matches.Count
	    ForEach ($item in $script:matches.Values)
	    {
            Log "Moving: $($item.Subject), $($item.Start) - $($item.End)" White
		    $moveIds.Add($item.Id)
            $i--
		    if ($moveIds.Count -ge $batchSize)
		    {
			    # Send the move request
                LogVerbose "Sending request to move $($deleteIds.Count) items ($i remaining)"
                ApplyEWSOAuthCredentials
			    [void]$script:service.MoveItems( $moveIds, $targetFolder.Id )
			    $moveIds = [Activator]::CreateInstance($genericItemIdList)
		    }
	    }
	    if ($moveIds.Count -gt 0)
	    {
            LogVerbose "Sending final move request for $($moveIds.Count) items"
		    [void]$script:service.MoveItems( $moveIds, $targetFolder.Id )
	    }
    }

    if ( $Delete )
    {
        # Delete the items in the delete list
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        $i = $script:matches.Count
	    ForEach ($item in $script:matches.Values)
	    {
            Log "Deleting: $($item.Subject), $($item.Start) - $($item.End)" White
		    $deleteIds.Add($item.Id)
            $i--
		    if ($deleteIds.Count -ge $batchSize)
		    {
			    # Send the delete request
                LogVerbose "Sending request to delete $($deleteIds.Count) items ($i remaining).  SendCancellationsMode: $SendCancellationsMode"
                ApplyEWSOAuthCredentials
			    [void]$script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $SendCancellationsMode, $Null )
			    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
		    }
	    }
	    if ($deleteIds.Count -gt 0)
	    {
            LogVerbose "Sending final delete request for $($deleteIds.Count) items.  SendCancellationsMode: $SendCancellationsMode"
            ApplyEWSOAuthCredentials
		    [void]$script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $SendCancellationsMode, $Null )
	    }
    }
}

function ProcessMailbox()
{
    # Process the mailbox
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
	}
	
    $Folder = $Null
	if ($FolderPath)
	{
		$Folder = GetFolder($FolderPath)
		if (!$Folder)
		{
			Write-Host "Failed to find folder $FolderPath" -ForegroundColor Red
			return
		}
	}
    else
    {
        if ($PublicFolders)
        {
            Write-Host "You must specify folder path when searching public folders" -ForegroundColor Red
            return
        }
        else
        {
		    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
		    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $mbx )
	        $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
    }

    # Note that we can't use a PowerShell hash table to build a list of Item Ids, as the hash table is case-insensitive
    # We use a .Net Dictionary object instead
    $script:matches = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
	SearchForAppointments $Folder
    ProcessBatches
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