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
	[Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,

	[Parameter(Mandatory=$False,HelpMessage="Folder to search for NDRs - if omitted, the Inbox folder is assumed")]
	[string]$FolderPath,

	[Parameter(Mandatory=$False,HelpMessage="If set, messages will be saved to this folder instead of sent from the mailbox.  You can specify multiple Pickup folders using an array, and a round robin process will be followed")]
	$SaveToPickupFolder = $null,

	[Parameter(Mandatory=$False,HelpMessage="If set, any messages that can't be saved to Pickup folder will instead be saved to this folder (for debugging purposes)")]
	$FailPickupFolder = $null,

	[Parameter(Mandatory=$False,HelpMessage="If set, this return-path will be stamped on resent messages")]
	[string]$ReturnPath = "",

	[Parameter(Mandatory=$False,HelpMessage="If set, we'll forward all messages directly to target server (based on MX or specified SMTP server list)")]
	[switch]$SendUsingSMTP,

	[Parameter(Mandatory=$False,HelpMessage="A list of SMTP servers for specific target email addresses (or domains).  Any listed here will be used in preference to MX.")]
	$SMTPServerList,

	[Parameter(Mandatory=$False,HelpMessage="If set, messages will be written directly into the recipients' mailbox(es).  Requires the authenticating account to have ApplicationImpersonation rights on those mailboxes.")]
	[switch]$WriteDirectlyToRecipientMailbox,

	[Parameter(Mandatory=$False,HelpMessage="Folder to move processed items into")]
	[string]$MoveProcessedItemsToFolder = "",

	[Parameter(Mandatory=$False,HelpMessage="Folder to move failed items into (those we attempted to process but were unable to)")]
	[string]$MoveFailedItemsToFolder = "",

	[Parameter(Mandatory=$False,HelpMessage="Folder to move encrypted items into (we won't attempt to process them)")]
	[string]$MoveEncryptedItemsToFolder = "",

	[Parameter(Mandatory=$False,HelpMessage="If set, any items that are encrypted will have the encrypted content removed")]
	[switch]$RemoveEncryptedAttachments,

	[Parameter(Mandatory=$False,HelpMessage="If an item is processed, but couldn't be moved, then the Id will be added to this file so that it can be ignored on future runs")]	
	[string]$IgnoreIdsLog = "",

	[Parameter(Mandatory=$False,HelpMessage="If set, all items processed (or failed to process) will be logged to the ignore file (recommended if messages are not being moved once processed)")]	
    [switch]$AddAllItemsToIgnoreLog,

	[Parameter(Mandatory=$False,HelpMessage="Batch size for processing NDRs (the number of items queried from the Inbox at one time)")]
	[int]$BatchSize = -1,

	[Parameter(Mandatory=$False,HelpMessage="If specified, checks for the messageclass are done clientside so that no search is required on the server.")]
	[switch]$FilterNDRsClientside,

	[Parameter(Mandatory=$False,HelpMessage="If specified, only this number of items will be processed (script will stop when this number is reached)")]
	[int]$MaxItemsToProcess = -1,

	[Parameter(Mandatory=$False,HelpMessage="If specified, any messages larger than this will be failed (without being sent)")]
	[int]$MaxMessageSize = -1,

	[Parameter(Mandatory=$False,HelpMessage="If specified, message will only be resent to recipient(s) listed here")]
	$OnlyResendTo,

	[Parameter(Mandatory=$False,HelpMessage="If specified, specified recipient(s) will be added to the message")]
	$AddResendTo,

	[Parameter(Mandatory=$False,HelpMessage="If specified, any messages found that have a blank From: header will have this address applied as the sender")]	
	[string]$DefaultFromAddress = "",

	[Parameter(Mandatory=$False,HelpMessage="If specified, message will only be resent if the recipient specified in OnlyResendTo parameter was an original recipient of the email.  If this isn't specified, then all messages will be resent.")]
	[switch]$ConfirmResendAddress,

	[Parameter(Mandatory=$False,HelpMessage="If specified, a message will be sent to this recipient when the script has completed.")]
	[string]$SendCompletionEmailTo = "",

	[Parameter(Mandatory=$False,HelpMessage="If original message not included as attachment, attempt to find it in Sent Items.")]
	[switch]$SearchSentItems,

	[Parameter(Mandatory=$False,HelpMessage="Output statistics to the specified CSV file")]
	[string]$StatsCSV,

	[Parameter(Mandatory=$False,HelpMessage="When set, the script will not process any messages but will collect statistics from folder being processed")]
	[switch]$CollectStatsOnly,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
	[Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts) - this requires the ADAL dlls to be available.")]
	[switch]$OAuth,

	[Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
	[string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

	[Parameter(Mandatory=$False,HelpMessage="The tenant Id in which the application is registered.  If missing, application is assumed to be multi-tenant and the common log-in URL will be used.")]
	[string]$OAuthTenantId = "",

	[Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
	[string]$OAuthRedirectUri = "http://localhost/code",
				
	[Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
	[switch]$Impersonate,

	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
	[string]$EwsUrl,
	
	[Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
	[string]$EWSManagedApiPath = "",
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
	[switch]$IgnoreSSLCertificate,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
	[switch]$AllowInsecureRedirection,
	
	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",

	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile
)
$script:ScriptVersion = "1.2.4"

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

Function Log([string]$Details, [ConsoleColor]$Colour, [switch]$SuppressWriteToScreen)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    if (!$SuppressWriteToScreen) { Write-Host $Details -ForegroundColor $Colour }
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

function LoadADAL
{
    # First of all, we check if ADAL is already available
    # To do this, we simply try to instantiate an authentication context to the common log-on Url.  If we get an object back, we have ADAL

    LogDebug "Checking for ADAL"
    $authenticationContextCommon = $null
    try
    {
        $authenticationContextCommon = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.windows.net/common", $False)
    } catch {}
    if ($authenticationContextCommon -ne $null)
    {
        LogVerbose "ADAL already available, no need to load dlls."
        return $true
    }

    # Load the ADAL libraries
    $adalDllsLocation = @()
    return $(LoadLibraries $false @("Microsoft.IdentityModel.Clients.ActiveDirectory.dll") ([ref]$adalDllsLocation) )
}

function GetOAuthCredentials
{
    # Obtain OAuth token for accessing mailbox
    param (
        [switch]$RenewToken
    )
    $exchangeCredentials = $null

    if ( $(LoadADAL) -eq $false )
    {
        Log "Failed to load ADAL, which is required for OAuth" Red
        Exit
    }

    $script:authenticationResult = $null
    if ([String]::IsNullOrEmpty($OAuthTenantId))
    {
        $authenticationContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.windows.net/common", $False)
    }
    else
    {
        $authenticationContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/$OAuthTenantId", $False)
    }
    if ($RenewToken)
    {
        $platformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto)
    }
    else
    {
        $platformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::SelectAccount)
    }

    $redirectUri = New-Object Uri($OAuthRedirectUri)
    $script:authenticationResult = $authenticationContext.AcquireTokenAsync("https://outlook.office365.com", $OAuthClientId, $redirectUri, $platformParameters)

    if ( !$authenticationResult.IsFaulted )
    {
        $script:oAuthToken = $authenticationResult.Result
        $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthToken.AccessToken)
        $Mailbox = $authenticationResult.Result.UserInfo.UniqueId
        LogVerbose "OAuth completed for $($authenticationResult.Result.UserInfo.DisplayableId), access token expires $($script:oAuthToken.ExpiresOn)"
    }
    else
    {
        ReportError "GetOAuthCredentials"
    }

    return $exchangeCredentials
}

function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( $script:authenticationResult -eq $null ) { return }
    if ( $script:services -eq $null ) { return }
    if ( $script:services.Count -lt 1 ) { return }
    if ( $script:authenticationResult.Result.ExpiresOn -gt [DateTime]::Now ) { return }

    # The token has expired and needs refreshing
    LogVerbose("OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($exchangeCredentials -eq $null) { return }
    if ( $script:authenticationResult.Result.ExpiresOn -le [DateTime]::Now )
    { 
        Log "OAuth Token renewal failed"
        exit # We no longer have access to the mailbox, so we stop here
    }

    LogVerbose "OAuth token successfully renewed; new expiry: $($script:oAuthToken.ExpiresOn)"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:authenticationResult.Result.AccessToken)
        }
        LogVerbose "Updated OAuth token for $($script.services.Count) ExchangeService objects"
    }
}

Function LoadEWSManagedAPI
{
    # We try to instantiate a FolderId object to test if we already have the EWS API loaded
    try
    {
        $itemId = $null
        $itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
        if ($itemId -ne $null)
        {
            return $true
        }
    }
    catch {}
    $Error.Clear()

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

function CreateService()
{
    param (
        [String]$smtpAddress,
        [Switch]$ForceImpersonation
    )
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
        $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
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
 
    $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)
    if ($Impersonate -or $ForceImpersonation)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpAddress)
	}

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    if (![String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        CreateTraceListener $exchangeService
        $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
        $exchangeService.TraceEnabled = $True
    }

    $script:services.Add($smtpAddress, $exchangeService)
    LogVerbose "Currently caching $($script:services.Count) ExchangeService objects" $true
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

        # Increase our throttling delay to try and avoid throttling (we only increase to a maximum delay of 15 seconds between requests)
        if ( $script:throttlingDelay -lt 15000)
        {
            if ($script:throttlingDelay -lt 1)
            {
                $script:throttlingDelay = 2000
            }
            else
            {
                $script:throttlingDelay = $script:throttlingDelay * 2
            }
            if ( $script:throttlingDelay -gt 15000)
            {
                $script:throttlingDelay = 15000
            }
        }
        LogVerbose "Updated throttling delay to $($script:throttlingDelay)ms"

        # Now back off for the time given by the server
        Log "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
        Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text"
        Log "Throttling budget should now be reset, resuming operations" Gray
        return $true
    }

    Log "Last server response: $($script:Tracer.LastResponse)" Red
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
        LogVerbose "No exchange service passed to ThrottledFolderBind, so using default"
        $exchangeService = $script:service
    }

    try
    {
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
            LogVerbose "Successfully bound to folder $folderId"
        }
        Sleep -Milliseconds $script:currentThrottlingDelay
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
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
        [Microsoft.Exchange.WebServices.Data.Folder]$RootFolder = $null,
        [String]$FolderPath = "",
        [switch]$Create
    )
        	
    if ( $RootFolder -eq $null )
    {
        # If we don't have a root folder, we assume the root of the message store
        if ($script:msgFolderRoot -eq $null)
        {
            LogVerbose "[GetFolder] Attempting to locate root message folder"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox )
            $script:msgFolderRoot = ThrottledFolderBind $folderId $null $script:service
            if ($script:msgFolderRoot -eq $null)
            {
                Log "[GetFolder] Failed to bind to message root folder" Red
                return $null
            }
            LogVerbose "[GetFolder] Retrieved root message folder"
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
                LogDebug "[GetFolder] Finding folder $($PathElements[$i])"
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
                $FolderResults = $Null
                try
                {
				    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                    Sleep -Milliseconds $script:throttlingDelay
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
					Log "[GetFolder] Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red
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
                            LogVerbose "[GetFolder] Created folder $($PathElements[$i])"
                        }
                        catch
                        {
					        # Failed to create the subfolder
					        $Folder = $null
					        Log "[GetFolder] Failed to create folder $($PathElements[$i]) in path $FolderPath" Red
					        break
                        }
                        $Folder = $subfolder
                    }
                    else
                    {
					    # Folder doesn't exist
					    $Folder = $null
					    Log "[GetFolder] Folder $($PathElements[$i]) doesn't exist in path $FolderPath" Red
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
        [String]$Sender )

    if ($Sender -match "\<(.+)\>")
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
    if ($senderService -eq $null)
    {
        Log "Failed to open mailbox of sender: $senderEmail" Red
        $script:errorItems++
        return
    }

    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $senderEmail )
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $mbx )
    LogVerbose "Connecting to Sent Items folder of $senderEmail"
    $folder = ThrottledFolderBind $folderId $null $senderService
    if ($folder -eq $null)
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
        Sleep -Milliseconds $script:currentThrottlingDelay
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
    $content = ""
    if ($endOfHeaders -gt 0)
    {
        $content = $MIME.SubString($endOfHeaders+2)
        $headers = $MIME.SubString(0,$endOfHeaders)
    }
    $headerLines = $headers -split "`r`n|`r|`n"

    LogVerbose "[ExtractHeaderValue] Analysing header block (contains $($headerLines.Count) lines)"
    $i=0
    do {
        if ( $headerLines[$i].StartsWith("$($headerName): ", [System.StringComparison]::OrdinalIgnoreCase) )
        {
            #LogVerbose "$($headerLines[$i])"
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
            LogVerbose "[ExtractHeaderValue] Found header: $headerName"
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

    LogVerbose "[ReplaceMIMEHeader] Analysing header block"
    $i=0
    do {
        if ($headerLines[$i].StartsWith("$($HeaderName): ") )
        {
            # This is the header to replace
            $headerFound = $true
            if (![String]::IsNullOrEmpty($HeaderValue))
            {
                $updatedHeaders.AppendLine("$($HeaderName): $HeaderValue") | out-null
                LogVerbose "[ReplaceMIMEHeader] Found $HeaderName header, replaced: $($HeaderName): $HeaderValue"
            }
            else
            {
                LogVerbose "[ReplaceMIMEHeader] Removed $HeaderName header"
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
            LogVerbose "[ReplaceMIMEHeader] Added header: $($HeaderName): $HeaderValue"
        }
    }
    LogVerbose "[ReplaceMIMEHeader] Header block analysis complete; $i lines processed"

    # Now we just need to put the new header and the content back together
    return "$($updatedHeaders.ToString())$content"
}

function SendUsingSMTP()
{
   param (
        [String]$MIME,
        $recipients,
        [String]$sender
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

        if ($line -ne $null)
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
                        LogVerbose "[StripEncryptedAttachmentsFromMime] Found MIME boundary: $boundary"
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
                    LogVerbose "[StripEncryptedAttachmentsFromMime] Found MIME boundary"
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
                            LogVerbose "[StripEncryptedAttachmentsFromMime] Encrypted MIME part (length $($unwrittenData.Length)) removed"
                            $unwrittenData = [System.Text.StringBuilder]::new()
                            $encryptedAttachmentFoundInMIMEPart = $false
                        }
                        else
                        {
                            # Keep this MIME part
                            LogVerbose "[StripEncryptedAttachmentsFromMime] MIME part contained no encrypted data"
                            [void]$updatedMIME.Append($unwrittenData.ToString())
                            $unwrittenData = [System.Text.StringBuilder]::new()
                        }
                    }
                }
                if ($line.StartsWith("Content-Type: application/x-microsoft-rpmsg-message"))
                {
                    LogVerbose "[StripEncryptedAttachmentsFromMime] Encrypted MIME part found: $line"
                    $encryptedAttachmentFoundInMIMEPart = $true
                }
            }
        }

    } while ($line -ne $null)

    if ($unwrittenData.Length -gt 0)
    {
        # We still have some data that we need to deal with
        if ($encryptedAttachmentFoundInMIMEPart)
        {
            LogVerbose "[StripEncryptedAttachmentsFromMime] Encrypted MIME part (length $($unwrittenData.Length)) removed"
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
    if ( ($folderObject -ne $null) -or [String]::IsNullOrEmpty(($TargetFolder)) )
    {
        return $folderObject
    }

    LogVerbose "[ValidateFolderMoveParameter] Locating folder: $TargetFolder"
    $folderObject = GetFolder $null $TargetFolder -Create
    if ($folderObject -eq $null)
    {
        Log "[ValidateFolderMoveParameter] Unable to find/create target folder specified in parameters" Red
        exit
    }
    Log "[ValidateFolderMoveParameter] Folder located: $TargetFolder"
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
                    ReportError "[ResendMessages]"
                    Log "[ResendMessages] Failed to create ignore file:  $IgnoreIdsLog" Red
                    exit
                }
                else
                {
                    Log "[ResendMessages] No existing items to ignore, created ignore file:  $IgnoreIdsLog" Green
                }
            }
        }
    }

    $progressActivity = "Processing NDRs"
    LogVerbose "[ResendMessages] Processing $($NDRs.Count) items"
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
            LogVerbose "[ResendMessages] Ignoring item: $($NDR.Id.UniqueId)"
            $script:ignoredItems++
            continue
        }

        # Load the message body (we only need the text version)
        $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body,
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ToRecipients, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ParentFolderId, $PidTagBody)
        Log "[ResendMessages] Retrieving message Id: $($NDR.Id.UniqueId)" Gray
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
                    LogVerbose "[ResendMessages] Resending to $($resendToAddress) based on OnlyResendTo parameter"
                    $resendTo += $resendToAddress
                }
            }
            if ( $resendTo.Count -lt 1 )
            {
                # Parsing the OnlyResendTo recipients didn't return any recipients...
                Log "[ResendMessages] OnlyResendTo parameter invalid: $OnlyResendTo" Red
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
                    if ($recipient.Address -ne $null)
                    {
                        if ($recipient.Address.StartsWith("IMCEAINVALID"))
                        {
                            Log "[ResendMessages] Invalid To recipient: $($recipient.Address)" Red
                        }
                        else
                        {
                            $address = $recipient.Address
                            if ( $address.StartsWith("=SMTP:") ) { $address = $address.SubString(6) }
                            LogVerbose "[ResendMessages] Resending to $($address) based on message recipients"
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
                                LogVerbose "[ResendMessages] Resending to $($mailToMatch.Groups[1].Value) based on message content"
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
            Log "[ResendMessages] Could not read failed recipients from ndr" Red
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
            LogVerbose "[ResendMessages] Updated To header value: $toHeader"

            if ($NDR.Attachments.Count -eq 1)
            {
                # Attachment is most likely the original message, so resend that
                LogVerbose "[ResendMessages] Original message attached to NDR"
                $Error.Clear()
                try
                {
                    $itemAttachment = $null
                    $itemAttachment = [Microsoft.Exchange.WebServices.Data.ItemAttachment]$NDR.Attachments[0]
                    if ($MaxMessageSize -gt 0)
                    {
                        if ($itemAttachment.Size -gt $MaxMessageSize)
                        {
                            Log "[ResendMessages] Item too large ($($itemAttachment.Size))" Red
                            $itemAttachment = $null
                            $ndrProcessFail = $true
                        }
                    }

                    if ($itemAttachment -ne $null)
                    {
                        $itemAttachment.Load([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
                    }

                    if ($itemAttachment -ne $null)
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
                                        LogVerbose "[ResendMessages] Encrypted item found, encrypted attachment removed. Original MIME length: $($MIME.Length)  Updated MIME length: $($clearMIME.Length)"
                                        $MIME = $clearMIME
                                        $ndrEncrypted = $true
                                    }
                                    else
                                    {
                                        Log "[ResendMessages] Item with two attachments is not encrypted, processing as normal message. Original MIME length: $($MIME.Length)  Updated MIME length: $($clearMIME.Length)" Yellow
                                    }
                                }
                                else
                                {
                                    LogVerbose "[ResendMessages] Encrypted item detected, will be ignored"
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
                                    if ($targetService -ne $null)
                                    {
                                        try
                                        {
                                            LogVerbose "[ResendMessages] Writing message into mailbox: $targetMailbox"
                                            $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::new($targetService)
                                            $mail.MimeContent = $MIME
                                            $mail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
                                            Log "[ResendMessages] Message successfully saved to $targetMailbox" Green
                                        }
                                        catch
                                        {
                                            ReportError "[ResendMessages]"
                                        }
                                    }
                                }
                                $ndrProcessed = $true
                            }
                            else
                            {
                                Log "[ResendMessages] Cannot save directly to mailbox as recipients could not be read" Red
                            }
                        }
                        else
                        {
                            
                            $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "CC" -HeaderValue ""
                            if ( ![String]::IsNullOrEmpty($ReturnPath) )
                            {
                                $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "Return-Path" -HeaderValue $ReturnPath
                            }

                            if (![String]::IsNullOrEmpty($toHeader))
                            {
                                $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "To" -HeaderValue $toHeader
                            }
                            if (!$CollectStatsOnly)
                            {
                                # We don't resend the message if we are only collecting statistics
                                if ($SendUsingSMTP)
                                {
                                    LogVerbose "[ResendMessages] Resending message over SMTP"
                                    if ( SendUsingSMTP -Mime $MIME -recipients $resendTo -Sender $NDR.Sender.Address )
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
                                    if ( $SaveToPickupFolder -eq $null)
                                    {
                                        # Send message from the mailbox
                                        LogVerbose "[ResendMessages] Resending message"
                                        $EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($script:service)
                                        $EmailMessage.MimeContent = $MIME
                                        try
                                        {
                                            $EmailMessage.Send()
                                            $ndrProcessed = $true
                                        } catch
                                        {
                                            ReportError "[ResendMessages - send message]"
                                            $ndrProcessFail = $true
                                        }
                                    }
                                    else
                                    {
                                        # Save message to pickup folder
                                        $ndrProcessed = SaveMIMEToPickupFolder -mime $MIME -WasEncrypted $ndrEncrypted
                                        LogVerbose "[ResendMessages] Save to pickup folder success: $ndrProcessed"
                                        $ndrProcessFail = !$ndrProcessed
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    LogVerbose "[ResendMessages] Failed to read attached message: $Error[0]"
                    $ndrProcessFail = $true
                }
            }
            else
            {
                $ndrProcessFail = $true
                LogVerbose "[ResendMessages] Original message not attached to NDR"
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
                LogVerbose "[ResendMessages] Attempting to extract message Id and sender from NDR"
                $messageId = ExtractHeaderValue $ndrBody "Message-ID"
                $sender = ExtractHeaderValue $ndrBody "From"
                if (![String]::IsNullOrEmpty($messageId) -and ![String]::IsNullOrEmpty($sender))
                {
                    LogVerbose "[ResendMessages] Attempting to resend message $messageId from $sender"
                    FindAndResendMessage $messageId $sender
                }
                else
                {
                    LogVerbose "[ResendMessages] Unable to read required information from NDR for resending"
                }
            }
            else
            {
                LogVerbose "[ResendMessages] Failed to read body of NDR"
            }
        }

        if ($CollectStatsOnly)
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
            if ( ($script:moveProcessedItemsToFolderFolder -ne $null) -and (!$CollectStatsOnly) )
            {
                LogVerbose "[ResendMessages] Moving processed item"
                try
                {
                    [void]$NDR.Move($script:moveProcessedItemsToFolderFolder.Id)
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                }
                ReportError "[ResendMessages]"
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
            if ( ($script:moveEncryptedItemsToFolderFolder -ne $null) -and (!$CollectStatsOnly))
            {
                LogVerbose "[ResendMessages] Moving encrypted item"
                try
                {
                    [void]$NDR.Move($script:moveEncryptedItemsToFolderFolder.Id)
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                }
                ReportError "[ResendMessages]"
            }
        }

        if ($ndrProcessFail)
        {
            # We encountered an issue processing this NDR, so we move to error folder
            
            $script:errorItems++
            if ( ($script:moveErrorItemsToFolderFolder -ne $null) -and (!$CollectStatsOnly) )
            {
                LogVerbose "[ResendMessages] Moving error item"
                try
                {
                    $movedItem = $NDR.Move($script:moveErrorItemsToFolderFolder.Id)
                    Log "[ResendMessages] Failed item id (moved to error folder): $($movedItem.Id.UniqueId)" Red
                }
                catch
                {
                    # If we have an error on move, then we need to store the Id of the item so that we don't process it again in the future
                    $addItemToIgnoreList = $true
                    Log "[ResendMessages] Failed item id (move failed): $($NDR.Id.UniqueId)" Red
                }
                ReportError "[ResendMessages]"
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
            $sender = ExtractHeaderValue $MIME "From"
            $subject = ExtractHeaderValue $MIME "Subject"
            $sentTime = ExtractHeaderValue $MIME "Date"
            foreach ($targetAddress in $resendTo)
            {
                "`"$messageId`",`"$sender`",`"$subject`",`"$sentTime`",`"$targetAddress`",`"$ndrProcessed`",`"$ndrProcessFail`",`"$ndrEncrypted`"" | Out-File $StatsCSV -Append
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
        Log "[SaveMIMEToPickupFolder] $($pickUpFolderList.Length) pickup folder(s) being used (round robin)"
    }

    $messageIsValid = $true

    if ( ![String]::IsNullOrEmpty($ReturnPath) )
    {
        # We need to update Sender also, otherwise Exchange will overwrite Return-Path
        $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "Sender" -HeaderValue $ReturnPath
    }

    # Check that the message has a valid From header (Sender and Return-Path are optional, so we don't check these)
    $from = ExtractHeaderValue -headers $mime -HeaderName "From"
    if ([String]::IsNullOrEmpty($from))
    {
        if (![String]::IsNullOrEmpty($DefaultFromAddress))
        {
            # No from address found, but we have a default one to apply
            $MIME = ReplaceMIMEHeader -MIME $MIME -HeaderName "From" -HeaderValue "From: $DefaultFromAddress"
            Log "[SaveMIMEToPickupFolder] From header was empty, replaced with `"From: $DefaultFromAddress`"" Yellow
        }
        else
        {
            Log "[SaveMIMEToPickupFolder] From header was empty, message not saved to pickup folder: $from" Red
            $messageIsValid = $false
        }
    }
    LogVerbose "[SaveMIMEToPickupFolder] From header: $from"
    if (!$RemoveEncryptedAttachments -and $WasEncrypted)
    {
        $messageIsValid = $false
        Log "[SaveMIMEToPickupFolder] Encrypted message not processed as -RemoveEncryptedAttachments not specified (encrypted attachments must be removed for successful processing)" Yellow
        return $false # We don't save these messages to any Pickup debug folder as we know why it has failed
    }

    if ( $messageIsValid )
    {
        $filename = "$($script:pickUpFolderList[$script:pickupFolderIndex])\$([DateTime]::Now.Ticks).eml"
        $script:pickupFolderIndex++
        if ($script:pickupFolderIndex -ge $pickUpFolderList.Length) { $script:pickupFolderIndex = 0 }

        try
        {
            Log "[SaveMIMEToPickupFolder] Saving email to: $fileName" Gray
            [IO.File]::WriteAllText($fileName, $mime)
            return $true
        }
        catch
        {
            ReportError "[SaveMIMEToPickupFolder]"
            return $false # No point in debugging a write failure, as this will be an IO issue                                     
        }
    }

    # If we get to this point, the message failed validation
    if ( ![String]::IsNullOrEmpty($FailPickupFolder) )
    {
        $filename = "$FailPickupFolder\$([DateTime]::Now.Ticks).eml"
        try
        {
            Log "[SaveMIMEToPickupFolder] Saving debug email to: $fileName" Gray
            [IO.File]::WriteAllText($fileName, $mime)
        }
        catch
        {
            ReportError "[SaveMIMEToPickupFolder]"                                      
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

    $total = $script:processedItems + $script:errorItems
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
    $i = 0
	
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
            Sleep -Milliseconds $script:currentThrottlingDelay
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
                    Sleep -Seconds 360
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
	if ($script:service -eq $Null)
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

    if ($CollectStatsOnly)
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
    Write-Host "The API can be downloaded from the Microsoft Download Centre: http://www.microsoft.com/en-us/search/Results.aspx?q=exchange%20web%20services%20managed%20api&form=DLC"
    Write-Host "Use the latest version available"
	Exit
}

Add-Type -AssemblyName System.Web

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($Username -or $Password)
    {
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red
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

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}


if (![String]::IsNullOrEmpty($SendCompletionEmailTo))
{
    # Send email that the script has finished

}