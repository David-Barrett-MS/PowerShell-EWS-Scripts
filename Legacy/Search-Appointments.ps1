#
# Search-Appointments.ps1
#
# By David Barrett, Microsoft Ltd. 2015 - 2018. Use at your own risk.  No warranties are given.
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

    [Parameter(Mandatory=$False,HelpMessage="If specified, only appointments with at least this number of attachments will be returned")] 
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
	$SendCancellationsMode = [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone,
	
	[Parameter(Mandatory=$False,HelpMessage="Folder path to which matching appointments will be moved")]
	[string]$MoveToFolder,
	
	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,

	[Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts) - this requires the ADAL dlls to be available")]
	[switch]$OAuth,

	[Parameter(Mandatory=$False,HelpMessage="The client (application) Id that this script will identify as.  Must be registered in Azure AD.")]
	[string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

	[Parameter(Mandatory=$False,HelpMessage="The tenant Id in which the application is registered.  If missing, application is assumed to be multi-tenant and the common log-in URL will be used.")]
	[string]$OAuthTenantId = "",

	[Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
	[string]$OAuthRedirectUri = "http://localhost/code",
		
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
	
	[Parameter(Mandatory=$False,HelpMessage="If specified, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled)")]
	[switch]$FastFileLogging,

	[Parameter(Mandatory=$False,HelpMessage="CSV Export - appointments are exported to this file")]	
	[string]$ExportCSV = "",

	[Parameter(Mandatory=$False,HelpMessage="If this parameter is specified, exported times are in UTC")]	
	[switch]$ExportUTC,
	
	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile

)
$script:ScriptVersion = "1.1.3"

# Define our functions

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    $logInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
    if ($FastFileLogging)
    {
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            ReportError "Opening log file"
        }
        if ($script:logFileStream)
        {
            $streamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            $streamWriter.WriteLine($logInfo)
            $streamWriter.Dispose()
            if ( ErrorReported("Writing log file") )
            {
                $FastFileLogging = $false
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
    if ($VerbosePreference -eq "SilentlyContinue") { return }
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
    if ($DebugPreference -eq "SilentlyContinue") { return }
    LogToFile $Details
}

Function LogToCSV([string]$Details)
{
    # Write details to CSV (if specified, otherwise just to console)

	Write-Host $Details -ForegroundColor White
	if ( $ExportCSV -eq "" ) { return	}

    $FileExists = Test-Path $ExportCSV
    if (!$FileExists)
    {
        if ($script:CSVHeaders -ne $Null)
        {
            $script:CSVHeaders | Out-File $ExportCSV
        }
    }

	$Details | Out-File $ExportCSV -Append
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

function LoadLibraries()
{
    param (
        [bool]$searchProgramFiles,
        [bool]$searchLocalAppData = $false,
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

        if ($searchLocalAppData)
        {
            if ($dll -eq $null)
            {
	            $dll = Get-ChildItem -Recurse $env:LOCALAPPDATA -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
            }
        }
        $script:LastError = $Error[0] # We do this to suppress any errors encountered during the search above

        if ($dll -eq $null)
        {
            Log "Unable to load locate $dllName" Red
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
	# Find and load the managed API
    $ewsApiLocation = @()
    $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -searchLocalAppData $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
    ReportError "LoadEWSManagedAPI"

    if (!$ewsApiLoaded)
    {
        # Failed to load the EWS API, so try to install it from Nuget
        Write-Host "EWS Managed API was not found - attempt to automatically download and install from Nuget?" -ForegroundColor White
        if ($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").Character -ne 'y')
        {
            Exit # Can't do anything with the EWS API
        }

        $ewsapi = Find-Package "Exchange.WebServices.Managed.Api"
        if (!$ewsapi)
        {
            Register-PackageSource -Name NuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet
            $ewsapi = Find-Package "Exchange.WebServices.Managed.Api" -Source Nuget
        }
        if ($ewsapi.Entities.Name.Equals("Microsoft"))
        {
	        # We have found EWS API package, so install as current user (confirm with user first)
		    Install-Package $ewsapi -Scope CurrentUser -Force
            $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $false -searchLocalAppData $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
            ReportError "LoadEWSManagedAPI"
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
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.dll") | Out-Null

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
		        public class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
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


    $global:TraceObject = $script:Tracer
    # Attach the trace listener to the Exchange service
    try
    {
        $service.TraceListener = $script:Tracer
    }
    catch
    {
        ReportError "Setting TraceListener"
        $service.TraceListener = $null
    }
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

    if ($exchangeService.URL.AbsoluteUri.ToLower().Equals("https://outlook.office365.com/ews/exchange.asmx"))
    {
        # This is Office 365, so we'll add a small delay to try and avoid throttling
        if ($script:currentThrottlingDelay -lt 100)
        {
            $script:currentThrottlingDelay = 100
            LogVerbose "Office 365 mailbox, throttling delay set to $($script:currentThrottlingDelay)ms"
        }
    }
 
    $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpAddress)
	}

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    if (![String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        CreateTraceListener $exchangeService
        if ($exchangeService.TraceListener -ne $null)
        {
            $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
            $exchangeService.TraceEnabled = $True
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

function ProcessItem( $item )
{
	# We have found an item, so this function handles any processing
    LoadItem $item

    if ( $script:CSVHeaders -eq $Null )
    {
        $script:CSVHeaders = """Mailbox"",""Subject"",""IsRead"",""Sent"",""Received"",""Sender"",""Organizer"",""Start"",""End"",""IsAllDay"",""AppointmentType"""
    }

	LogToCSV "`"$Mailbox`",`"$($item.Subject)`",`"$($item.IsRead)`",`"$(ExportTime($item.DateTimeSent))`",`"$(ExportTime($item.DateTimeReceived))`",`"$($item.Sender)`",`"$($item.Organizer)`",`"$((ExportTime($item.Start)))`",`"$((ExportTime($item.End)))`",`"$($item.IsAllDayEvent)`",`"$($item.AppointmentType.ToString())`"" White

	# Add the item to our list of matches (for batch processing later)
    if ( $script:matches.ContainsKey($item.Id.UniqueId) )
    {
        LogVerbose "Item not added to match list as matching Id already present"
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
        #$PidLidIsRecurring = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x8223, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
        $PidLidIsRecurring = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8223, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidLidIsRecurring, $true)
        #$filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring, $true)

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
		    if ($moveIds.Count -ge 500)
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
		    if ($deleteIds.Count -ge 500)
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