#
# RecoverDeletedItems.ps1
#
# By David Barrett, Microsoft Ltd. 2015-2023. Use at your own risk.  No warranties are given.
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

    [Parameter(Mandatory=$False,HelpMessage="Start date (if items are older than this, they will be ignored).")]
    [ValidateNotNullOrEmpty()]
    [datetime]$RestoreStart,
	
    [Parameter(Mandatory=$False,HelpMessage="End date (if items are newer than this, they will be ignored).")]
    [ValidateNotNullOrEmpty()]
    [datetime]$RestoreEnd,

    [Parameter(Mandatory=$False,HelpMessage="Policy tag of items to restore (only items with this tag will be restored).")]
    [ValidateNotNullOrEmpty()]
    [string]$RestorePolicyTag,

    [Parameter(Mandatory=$False,HelpMessage="Folder to restore from (if not specified, items are recovered from retention).  Use WellKnownFolderNames.DeletedItems to restore from Deleted Items folder.")]	
    [string]$RestoreFromFolder,

    [Parameter(Mandatory=$False,HelpMessage="Folder to restore to (if not specified, items are recovered based on where they were deleted from, or their item type).")]	
    [string]$RestoreToFolder,

    [Parameter(Mandatory=$False,HelpMessage="If this is specified and the restore folder needs to be created, the default item type for the created folder will be as defined here.  If missing, the default will be IPF.Note.")]	
    [string]$RestoreToFolderDefaultItemType = "IPF.Note",

    [Parameter(Mandatory=$False,HelpMessage="If this is specified then the item is copied back to the restore folder instead of being moved.")]
    [switch]$RestoreAsCopy,

    [Parameter(Mandatory=$False,HelpMessage="A list of message classes that will be recovered (any not listed will be ignored, unless the parameter is missing in which case all classes are restored).")]
    $RestoreMessageClasses,
    
    [Parameter(Mandatory=$False,HelpMessage="If specified, any emails sent from this address will be considered as sent from the mailbox owner (can help with Sent Item matching).")]
    [ValidateNotNullOrEmpty()]
    [string]$MyEmailAddress,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox).")]
    [switch]$Archive,

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox), and paths are from root (i.e. above Top of Information store).")]
    [switch]$ArchiveRoot,

    [Parameter(Mandatory=$False,HelpMessage="If accessing Exchange 2007, different logic is needed to restore, so this switch must be specified.")]
    [switch]$Exchange2007,

    [Parameter(Mandatory=$False,HelpMessage="If specified, and the PidLidSpamOriginalFolder property is set on the message, the script will attempt to restore to that folder.")]
    [switch]$UseJunkRestoreFolder,

    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS.")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
    [Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts) - this requires the ADAL dlls to be available.")]
    [switch]$OAuth,

    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
    [string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id of the tenant being accessed.")]
    [string]$OAuthTenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$OAuthSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="For debugging purposes.")]
    [switch]$OAuthDebug,

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox.")]
    [switch]$Impersonate,
	
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used).")]
    [string]$EwsUrl,
	
    [Parameter(Mandatory=$False,HelpMessage="If specified, requests are directed to Office 365 endpoint (this overrides -EwsUrl).")]
    [switch]$Office365,

    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed).")]
    [string]$EWSManagedApiPath = "",
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate).")]	
    [switch]$IgnoreSSLCertificate,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover.")]	
    [switch]$AllowInsecureRedirection,
	
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified.")]	
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="If selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled).")]
    [switch]$FastFileLogging,

    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file.")]	
    [string]$TraceFile,
	
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, actions that would be taken will be logged, but nothing will be changed.")]
    [switch]$WhatIf
	
)
$script:ScriptVersion = "1.2.0"

# Define our functions

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

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Host $Details -ForegroundColor $Colour

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

    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Verbose $Details
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
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

    $cca = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca = $cca.WithCertificate($OAuthCertificate)
    $cca = $cca.WithTenantId($OAuthTenantId)
    $cca = $cca.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    $script:Impersonate = $true
}

function GetTokenViaCode
{
    if ($script:oAuthToken -eq $null)
    {
        # We don't yet have a token, so need to acquire auth code
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
            $authcode = $authcode.Substring(0, $codeEnd)
        }
        Write-Verbose "Using auth code: $authcode"
        # Use the auth code to obtain our access and refresh token
        $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$authcode;redirect_uri=$OAuthRedirectUri}
    }
    else
    {
        # This is a renewal, so we use the refresh token previously acquired (no need for auth code)
        $body = @{grant_type="refresh_token";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;refresh_token=$script:oAuthToken.refresh_token}
    }

    # Acquire token
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Log "Failed to obtain OAuth token" Red
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

function GetOAuthCredentials
{
    # Obtain OAuth token for accessing mailbox
    param (
        [switch]$RenewToken
    )
    $exchangeCredentials = $null

    if ($script:oauthToken -ne $null -and -not $RenewToken)
    {
        # We already have a token
        if ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1))
        {
            # Token still valid, so return that
            $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
            return $exchangeCredentials
        }
    }
    # Token needs renewing

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
        GetTokenViaCode
    }

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $script:services -eq $null ) { return }
    if ( $script:services.Count -lt 1 ) { return }
    if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    LogVerbose("OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($exchangeCredentials -eq $null) { return }
    if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -le [DateTime]::Now )
    { 
        Log "OAuth Token renewal failed" Red
        exit # We no longer have access to the mailbox, so we stop here
    }

    Log "OAuth token successfully renewed; new expiry: $($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in))"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = $exchangeCredentials
        }
        LogVerbose "[ApplyEWSOAuthCredentials] Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
    }
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
    $ewsApiLocation = @()
    $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
    ReportError "LoadEWSManagedAPI"
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
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

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
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
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
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable=$False
        $Params.GenerateInMemory=$True
        $Params.IncludeDebugInformation=$False
	    $Params.ReferencedAssemblies.Add("System.dll") | Out-Null
        $Params.ReferencedAssemblies.Add($EWSManagedApiPath) | Out-Null

        $traceFileForCode = $traceFile.Replace("\", "\\")

        if (![String]::IsNullOrEmpty($TraceFile))
        {
            Log "Tracing to: $TraceFile"
        }

        $TraceListenerClass = @"
		    using System;
		    using System.Text;
		    using System.IO;
		    using System.Threading;
		    using Microsoft.Exchange.WebServices.Data;
		
            namespace TraceListener {
		        class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
		        {
			        private StreamWriter _traceStream = null;
                    private string _lastResponse = String.Empty;

			        public EWSTracer()
			        {
				        try
				        {
					        _traceStream = File.AppendText("$traceFileForCode");
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

				        lock (this)
				        {
					        try
					        {
						        _traceStream.WriteLine(traceMessage);
						        _traceStream.Flush();
					        }
					        catch { }
				        }
			        }

                    public string LastResponse
                    {
                        get { return _lastResponse; }
                    }
		        }
            }
"@

        $TraceCompilation=$Provider.CompileAssemblyFromSource($Params,$TraceListenerClass)
        $TraceAssembly=$TraceCompilation.CompiledAssembly
        $script:Tracer=$TraceAssembly.CreateInstance("TraceListener.EWSTracer")
    }

    # Attach the trace listener to the Exchange service
    $service.TraceListener = $script:Tracer
}

function CreateService($smtpAddress)
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
    if ($Exchange2007)
    {
        $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
    }
    else
    {
        $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    }

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

    if ($exchangeService.URL.AbsoluteUri.ToLower().Equals("https://outlook.office365.com/ews/exchange.asmx"))
    {
        # This is Office 365, so we'll add a small delay to try and avoid throttling
        if ($script:currentThrottlingDelay -lt 100)
        {
            $script:currentThrottlingDelay = 100
            LogVerbose "Office 365 mailbox, throttling delay set to $($script:currentThrottlingDelay)ms"
        }
    }
 
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpAddress)
        $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)
	}

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    CreateTraceListener $exchangeService
    $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $exchangeService.TraceEnabled = $True

    $script:services.Add($smtpAddress, $exchangeService)
    return $exchangeService
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
        IncreaseThrottlingDelay

        # Now back off for the time given by the server
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
    if ($exchangeService -eq $null)
    {
        $exchangeService = $script:service
    }

    try
    {
        ApplyEWSOAuthCredentials
        if ($propset -eq $null)
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
        }
        else
        {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
        }
        if (!($folder -eq $null))
        {
            LogVerbose "Successfully bound to folder $($folder.DisplayName)"
        }
        Start-Sleep -Milliseconds $script:currentThrottlingDelay
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
            ApplyEWSOAuthCredentials
            if ($propset -eq $null)
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
            }
            else
            {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
            }
            if (!($folder -eq $null))
            {
                LogVerbose "Successfully bound to folder $($folder.DisplayName)"
            }
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    LogVerbose "FAILED to bind to folder $folderId"
    return $null
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
        $script:folderPathCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)

    if (!$Folder.Id)
    {
        # This isn't a folder.  Assume it's an Id and try binding to the folder
        LogVerbose "Retrieving path for folder ID : $Folder"

        if ($script:folderPathCache.ContainsKey($Folder))
        {
            return $script:folderPathCache[$Folder]
        }
        $Folder = ThrottledFolderBind $Folder $propset $script:service
        $parentFolder = $Folder
    }
    else
    {
        LogVerbose "Retrieving path for folder: $($Folder.DisplayName)"
        if ($script:folderPathCache.ContainsKey($Folder.Id.UniqueId))
        {
            return $script:folderPathCache[$Folder.Id.UniqueId]
        }
        $parentFolder = ThrottledFolderBind $Folder.Id $propset $Folder.Service
    }

    
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
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset $Folder.Service
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    $script:folderPathCache.Add($Folder.Id, $folderPath)
    return $folderPath
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$RootFolder,
        [String]$FolderPath = $null,
        [bool]$Create,
        [String]$CreatedFolderType = "IPF.Note")
        	
    if ( [String]::IsNullOrEmpty($FolderPath) )
    {
        LogVerbose "GetFolder called with null folder path"
        return $null
    }

    if ($FolderPath.ToLower().StartsWith("wellknownfoldername"))
    {
        # Well known folder, so bind to it directly
        $wkf = $FolderPath.SubString(20)
        LogVerbose "Attempting to bind to well known folder: $wkf ($mbx)"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind($folderId)
        return $Folder
    }

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
                    Start-Sleep -Milliseconds $script:currentThrottlingDelay
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
					    $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($RootFolder.Service)
					    $subfolder.DisplayName = $PathElements[$i]
                        $subfolder.FolderClass = $CreatedFolderType
                        ApplyEWSOAuthCredentials
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

function ReadMailboxFolderHierarchy()
{
    # We read the mailbox folder tree to create a dictionary of folder Ids referenced to their PR_SOURCE_KEY (this enables original folder recovery)

    $PR_SOURCE_KEY = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $folderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, $PR_SOURCE_KEY, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
    $moreFolders = $true
    $folderView.Offset = 0

    LogVerbose "Building folder hierarchy"
    $script:FoldersBySourceKey = @{}

    while ($moreFolders)
    {
        ApplyEWSOAuthCredentials
        if ($Archive)
        {
            $findResults = $script:service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $folderView)
        }
        elseif ($ArchiveRoot)
        {
            $findResults = $script:service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot, $folderView)
        }
        else
        {
            $findResults = $script:service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
        }
        $folderView.Offset += 1000
        $moreFolders = $findResults.MoreAvailable

        foreach ($folder in $findResults.Folders)
        {
            $folderSourceKey = $null
            if ($folder.ExtendedProperties[0].PropertyDefinition -eq $PR_SOURCE_KEY)
            {
                # For some reason, the PowerShell hash table lookup doesn't work with binary values as a key, so convert to string
                $folderSourceKey = [System.BitConverter]::ToString($folder.ExtendedProperties[0].Value)
            }
            if ($folderSourceKey -ne $null)
            {
                LogVerbose "$($folder.DisplayName) = $($folderSourceKey): $($folder.Id)"
                $script:FoldersBySourceKey.Add($folderSourceKey, $folder.Id)
            }
        }
    }
}

function ConvertEntryId($entryId)
{
    # Use EWS ConvertId function to convert from EntryId to EWS Id

    $id = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
    $id.Mailbox = $mailbox
    $id.UniqueId = [System.BitConverter]::ToString($entryId) -replace "-", ""
    LogVerbose "EntryId as string: $($id.UniqueId)"
    $id.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId
    $ewsId = $Null
    ApplyEWSOAuthCredentials
    $ewsId = $script:service.ConvertId($id, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
    LogVerbose "EWS Id: $($ewsId.UniqueId)"
    return $ewsId.UniqueId
}

Function RecoverFromFolder()
{
	# Process all the items in the given folder and move them back to previous location
	
	if ($args -eq $null)
	{
		throw "No folder specified for RecoverFromFolder"
	}
	$Folder=$args[0]

    ReadMailboxFolderHierarchy
	
    $LastActiveParentID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x348A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    $PidLidSpamOriginalFolder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x859C,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    $PidTagPolicyTag = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    #$LastActiveParentID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

    if ($RestoreStart -or $RestoreEnd)
    {
        LogVerbose "RestoreStart: $RestoreStart"
        LogVerbose "RestoreEnd: $RestoreEnd"
    }

    if (-not [String]::IsNullOrEmpty($RestorePolicyTag))
    {
        $restorePolicyTagGuid = $null
        try
        {
            $restorePolicyTagGuid = [guid]::Parse($RestorePolicyTag)
        }
        catch
        {
            Log "RestorePolicyTag is not a valid Guid" Red
            exit
        }
        LogVerbose "RestorePolicyTag: $RestorePolicyTag"
    }

	$MoreItems=$true
    $skipped = 0
	
	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, $skipped, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                                   [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsFromMe, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $PidLidSpamOriginalFolder, $LastActiveParentID, $PidTagPolicyTag)

        if ($Exchange2007)
        {
            $View.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::SoftDeleted
            $FindResults=$service.FindItems($Folder.Id, $View)
        }
        else
        {
            ApplyEWSOAuthCredentials
		    $FindResults=$service.FindItems($Folder.Id, $View)
        }
		
		ForEach ($Item in $FindResults.Items)
		{
			LogVerbose "Item $($Item.ItemClass): $($Item.Id.UniqueId)"
            LogVerbose "Last modified time: $($Item.LastModifiedTime)"

            $itemShouldBeRestored = $True
            if ($restorePolicyTagGuid)
            {
                foreach ($prop in $Item.ExtendedProperties)
                {
                    if ($prop.PropertyDefinition -eq $PidTagPolicyTag)
                    {
                        $itemPolicyTagGuid = [System.Guid]::new($prop.Value)
                        LogVerbose "PidTagPolicyTag: $itemPolicyTagGuid vs $($restorePolicyTagGuid.ToString())"

                        if ($itemPolicyTagGuid -ne $restorePolicyTagGuid)
                        {
                            LogVerbose "PidTagPolicyTag does not match filter"
                            $itemShouldBeRestored = $false
                        }
                    }
                }
            }

            if ($RestoreStart)
            {    
                if ($Item.LastModifiedTime -lt $RestoreStart) { $itemShouldBeRestored = $False; LogVerbose "Item is not within restore time range (start check)" }
            }
            if ($RestoreEnd)
            {
                if ($Item.LastModifiedTime -gt $RestoreEnd) { $itemShouldBeRestored = $False; LogVerbose "Item is not within restore time range (end check)" }
            }

            if ($RestoreMessageClasses -and $itemShouldBeRestored)
            {
                $validMessageClass = $false
                foreach ($messageClass in $RestoreMessageClasses)
                {
                    if ( $messageClass.Equals($Item.ItemClass) )
                    {
                        $validMessageClass = $true
                        LogVerbose "Message class matches restore criteria: $($Item.ItemClass)"
                        break
                    }
                }
                $itemShouldBeRestored = $validMessageClass
                if (!$validMessageClass)
                {
                    LogVerbose "Item does not match message class being restored"
                }
            }
            $moveToFolder = $null
            $targetFolder = $null

            if ($itemShouldBeRestored)
            {
                if ($script:RestoreTargetFolder)
                {
                    # We're restoring all items to a specific folder
                    $targetFolder = $script:RestoreTargetFolder.Id
                    $moveToFolder = GetFolderPath($targetFolder)
                }
                else
                {
                    if ($UseJunkRestoreFolder)
                    {
                        # Check to see if we have $UseJunkRestoreFolder, as this will allow us to restore to the original folder (if it still exists)
                        foreach ($extendedProperty in $Item.ExtendedProperties)
                        {
                            if ($extendedProperty.PropertyDefinition -eq $PidLidSpamOriginalFolder)
                            {
                                # We've got an EntryId for the folder it was deleted from
                                $ewsId = $null
                                $ewsId = ConvertEntryId $extendedProperty.Value
                                if ($ewsId -ne $null)
                                {
                                    $moveToFolder = $ewsId
                                    $targetFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId($moveToFolder)
                                    LogVerbose "PidLidSpamOriginalFolder: $moveToFolder"
                                    break
                                }
                            }
                        }
                        if ( [String]::IsNullOrEmpty($moveToFolder) )
                        {
                            LogVerbose "No PidLidSpamOriginalFolder property found on item"
                        }
                    }

                    if ( [String]::IsNullOrEmpty($moveToFolder) )
                    {
                        # Check to see if we have $LastActiveParentEntryID, as this will allow us to restore to the original folder (if it still exists)
                        $lastIdFound = $false
                        foreach ($extendedProperty in $Item.ExtendedProperties)
                        {

                            if ($extendedProperty.PropertyDefinition -eq $LastActiveParentID)
                            {
                                # We have last active folder, so let's see if the Id is still valid
                                $propValue = [System.BitConverter]::ToString($extendedProperty.Value)
                                LogVerbose "LastActiveParentEntryID: $propValue"
                                $lastIdFound = $true

                                # Last active folder Id is the PidTagSourceKey value of the folder
                                if ($script:FoldersBySourceKey.Contains($propValue))
                                {
                                    $moveToFolder = $script:FoldersBySourceKey[$propValue]
                                    $targetFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId($moveToFolder)
                                    LogVerbose "LastActiveParentEntryID: $moveToFolder"
                                    break
                                }
                                else
                                {
                                    Log "LastActiveParentEntryID was not found in this mailbox" Red
                                }
                            }
                        }
                        if ( !$lastIdFound )
                        {
                            LogVerbose "No LastActiveParentEntryID property found on item"
                        }
                    }
                }

                if ([String]::IsNullOrEmpty($moveToFolder))
                {
			        switch -wildcard ($Item.ItemClass)
			        {
				        "IPM.Appointment*"
				        {
					        # Appointment, so move back to calendar
                            $moveToFolder = "Calendar"
				        }
				
				        "IPM.Note*"
				        {
					        # Message; need to determine if sent or not
					        Write-Verbose "Message is from me: $($Item.IsFromMe)"
                            Write-Verbose "Message sender: $($Item.Sender)"
                            $isFromMe = $Item.IsFromMe
                            if (![String]::IsNullOrEmpty($MyEmailAddress))
                            {
                                if ($MyEmailAddress.ToLower().Equals($Item.Sender.Address.ToLower())) { $isFromMe = $true }
                            }

					        if ($isFromMe)
					        {
						        # This is a sent message
                                $moveToFolder = "SentItems"
					        }
					        else
					        {
						        # This is a received message
                                $moveToFolder = "Inbox"
					        }
				        }
				
				        "IPM.StickyNote*"
				        {
					        # Sticky note, move back to Notes folder
                            $moveToFolder = "Notes"
				        }
				
				        "IPM.Contact*"
				        {
					        # Contact, so move back to Contacts folder
                            $moveToFolder = "Contacts"
				        }
				
				        "IPM.Task*"
				        {
					        # Task, so move back to Tasks folder
                            $moveToFolder = "Tasks"
				        }
				
				        default
				        {
					        Log "Item was not a class supported for recovery: $($Item.ItemClass)" Red
                            if (!$WhatIf) { $skipped++ }
				        }
			        }
                }

                if (![String]::IsNullOrEmpty($moveToFolder))
                {
                    if ([String]::IsNullOrEmpty($targetFolder))
                    {
                        $targetFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$moveToFolder
                    }
                    else
                    {
                        $moveToFolder = GetFolderPath($targetFolder)
                    }

                    if (!$WhatIf)
                    {
                        # Move the item
                        ApplyEWSOAuthCredentials
                        try
                        {
                            if ($RestoreAsCopy)
                            {
                                [void]$Item.Copy($targetFolder)
                                Log "Item $($Item.ItemClass): $($Item.Id.UniqueId) copied to $moveToFolder"
                            }
                            else
                            {
                                [void]$Item.Move($targetFolder)
                                Log "Item $($Item.ItemClass): $($Item.Id.UniqueId) moved to $moveToFolder"
                            }
                        }
                        catch
                        {
                            ReportError "RecoverItem"
                            Log "Item $($Item.ItemClass): $($Item.Id.UniqueId) FAILED to recover item to $moveToFolder" Red
                        }
                    }
                    else
                    {
                        Log "Item $($Item.ItemClass): $($Item.Id.UniqueId) would be moved to $moveToFolder"
                    }
                }
                if ($WhatIf) { $skipped++ }
            }
            else
            {
                Write-Verbose "Item $($Item.ItemClass): $($Item.Id.UniqueId) doesn't match restore criteria"
                $skipped++
            }
            

		}
		$MoreItems=$FindResults.MoreAvailable
	}
}

function ProcessMailbox()
{
    # Process the mailbox
    Write-Host "Processing $(if ($Archive -or $ArchiveRoot) { "archive "})mailbox $Mailbox" -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
        return
	}
	
	$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
    if ($Archive)
    {
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx )
    }
    elseif ($ArchiveRoot)
    {
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot, $mbx )
    }

    # Bind to root folder (fail if unsuccessful)
    $rootFolder = ThrottledFolderBind $rootFolderId $null $script:service
    if ($rootFolder -eq $null) { return }

    if (![String]::IsNullOrEmpty($RestoreToFolder))
    {
        # We have a folder specified for the restore, so ensure it exists (we'll try to create it if it doesn't)
        LogVerbose "Locating folder to restore items to: $RestoreToFolder"
        $folder = GetFolder $rootFolder $RestoreToFolder $True $RestoreToFolderDefaultItemType
        if (!$folder)
        {
            Log "Unable to find or create target folder for recovery: $RestoreToFolder" Red
        }
        $script:RestoreTargetFolder = $folder
    }

    if (![String]::IsNullOrEmpty($RestoreFromFolder))
    {
        # We are recovering from a specific folder
        $folder = GetFolder $rootFolder $RestoreFromFolder $False $RestoreToFolderDefaultItemType
        if ($folder)
        {
            RecoverFromFolder $folder
        }
    }
    else
    {
        # We are recovering from retention
        if ($Exchange2007)
        {
            $inbox = $Null
	        $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $mbx )
            if ($Archive)
            {
                Log "Archive recovery not supported for Exchange 2007"
                return
            }
            $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $FolderId)
            if ($inbox -eq $Null)
            {
                Log "Failed to open Inbox" Red
                return
            }
            RecoverFromFolder $inbox
        }
        else
        {
            $RecoverableItemsRoot = $Null
	        $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsDeletions, $mbx )
            if ($Archive)
            {
                $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRecoverableItemsDeletions, $mbx )
            }
	        $RecoverableItemsRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $FolderId)

            if ($RecoverableItemsRoot -eq $Null)
            {
                Log "Failed to open Recoverable Items folder" Red
                return
            }

            RecoverFromFolder $RecoverableItemsRoot
        }
    }
}


# The following is the main script

if ( [string]::IsNullOrEmpty($Mailbox) -and !$OAuth)
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

if ($Office365)
{
    $OAuth = $true
}

  

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