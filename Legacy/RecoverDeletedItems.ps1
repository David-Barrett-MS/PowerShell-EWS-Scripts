#
# RecoverDeletedItems.ps1
#
# By David Barrett, Microsoft Ltd. 2015-2020. Use at your own risk.  No warranties are given.
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

	[Parameter(Position=1,Mandatory=$False,HelpMessage="Start date (if items are older than this, they will be ignored)")]
	[ValidateNotNullOrEmpty()]
	[datetime]$RestoreStart,
	
	[Parameter(Position=2,Mandatory=$False,HelpMessage="End date (if items are newer than this, they will be ignored)")]
	[ValidateNotNullOrEmpty()]
	[datetime]$RestoreEnd,

	[Parameter(Mandatory=$False,HelpMessage="Folder to restore from (if not specified, items are recovered from retention)")]	
	[string]$RestoreFromFolder,

	[Parameter(Mandatory=$False,HelpMessage="Folder to restore to (if not specified, items are recovered based on where they were deleted from, or their item type)")]	
	[string]$RestoreToFolder,

	[Parameter(Mandatory=$False,HelpMessage="If this is specified and the restore folder needs to be created, the default item type for the created folder will be as defined here.  If missing, the default will be IPF.Note.")]	
	[string]$RestoreToFolderDefaultItemType = "IPF.Note",

	[Parameter(Mandatory=$False,HelpMessage="If this is specified then the item is copied back to the mailbox instead of being moved.")]
	[switch]$RestoreAsCopy,

    [Parameter(Mandatory=$False,HelpMessage="A list of message classes that will be recovered (any not listed will be ignored, unless the parameter is missing in which case all classes are restored)")]
    $RestoreMessageClasses,
    
	[Parameter(Mandatory=$False,HelpMessage="If you specify this, any emails sent from this address will be considered as sent from the mailbox owner (can help with Sent Item matching)")]
	[ValidateNotNullOrEmpty()]
	[string]$MyEmailAddress,

	[Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox)")]
	[switch]$Archive,

	[Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox), and paths are from root (i.e. above Top of Information store)")]
	[switch]$ArchiveRoot,

	[Parameter(Mandatory=$False,HelpMessage="If we're accessing Exchange 2007, we need different logic")]
	[switch]$Exchange2007,


	[Parameter(Mandatory=$False,HelpMessage="If specified, and the PidLidSpamOriginalFolder property is set on the message, the script will attempt to restore to that folder")]
	[switch]$UseJunkRestoreFolder,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
	[Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts) - this requires the ADAL dlls to be available")]
	[switch]$OAuth,
	
	[Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
	[string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",
	
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

	[Parameter(Mandatory=$False,HelpMessage="If selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled)")]
	[switch]$FastFileLogging,

	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile,
	
	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, actions that would be taken will be logged, but nothing will be changed")]
	[switch]$WhatIf
	
)
$script:ScriptVersion = "1.1.6"

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

$script:fastLogWrite = $FastFileLogging
Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    $logInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
    if ($script:fastLogWrite)
    {
        Write-Host "Fast log write: $($script:fastLogWrite)" -ForegroundColor Yellow
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            try
            {
                $script:logFileStream = New-Object IO.FileStream($LogFile, @([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            } catch {}
            if ( $(ErrorReported "Opening log file") )
            {
                $script:fastLogWrite = $false
                Write-Host "Disabled fast log write: $($script:fastLogWrite)" -ForegroundColor Yellow
            }
        }
        if ($script:logFileStream)
        {
            Write-Host "Hello1" -ForegroundColor Cyan
            try
            {
                $streamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            } catch {}
            Write-Host "Hello5" -ForegroundColor Cyan
            if ( !$(ErrorReported "Opening log stream writer") )
            {
                Write-Host "Hello2" -ForegroundColor Cyan
                try
                {
                    $streamWriter.WriteLine($logInfo)
                    $streamWriter.Dispose()
                }
                catch {}
                if ( !$(ErrorReported "Writing log file") )
                {
                    return
                }
            }
            Write-Host "Hello3" -ForegroundColor Cyan
            $script:fastLogWrite = $false
            Write-Host "Disabled fast log write: $($script:fastLogWrite)" -ForegroundColor Yellow
        }
        Write-Host "Hello4" -ForegroundColor Cyan
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
		        LogVerbose "Loading $dllName v$($dll.VersionInfo.FileVersion) found at: $($dll.VersionInfo.FileName)"
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

    Log "OAuth token successfully renewed; new expiry: $($script:oAuthToken.ExpiresOn)"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($exchangeCredentials)
        }
        LogVerbose "Updated OAuth token for $($script.services.Count) ExchangeService objects"
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
        if ($script:folderPathCache.ContainsKey($Folder))
        {
            return $script:folderPathCache[$Folder]
        }
        $Folder = ThrottledFolderBind $Folder $propset $script:service
    }
    else
    {
        if ($script:folderPathCache.ContainsKey($Folder.Id.UniqueId))
        {
            return $script:folderPathCache[$Folder.Id.UniqueId]
        }
    }

    $parentFolder = ThrottledFolderBind $Folder.Id $propset $Folder.Service
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
    $folderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, $PR_SOURCE_KEY)
    $moreFolders = $true
    $folderView.Offset = 0

    LogVerbose "Building folder hierarchy"
    $script:FoldersBySourceKey = @{}

    while ($moreFolders)
    {
        ApplyEWSOAuthCredentials
        $findResults = $script:service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)
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
                LogVerbose "$($folderSourceKey): $($folder.Id)"
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
	# Process all the items in the given folder and move them back to mailbox
	
	if ($args -eq $null)
	{
		throw "No folder specified for RecoverFromFolder"
	}
	$Folder=$args[0]

    ReadMailboxFolderHierarchy
    if ($RestoreStart -or $RestoreEnd)
    {
        LogVerbose "RestoreStart: $RestoreStart"
        LogVerbose "RestoreEnd: $RestoreEnd"
    }
	
	# Set parameters - we will process in batches of 500 for the FindItems call
	$MoreItems=$true
    $skipped = 0
    $LastActiveParentID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x348A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    $PidLidSpamOriginalFolder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x859C,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    #$LastActiveParentID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65E0,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
	
	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, $skipped, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,
                                   [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsFromMe, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $PidLidSpamOriginalFolder, $LastActiveParentID)
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
                        foreach ($extendedProperty in $Item.ExtendedProperties)
                        {

                            if ($extendedProperty.PropertyDefinition -eq $LastActiveParentID)
                            {
                                # We have last active folder, so let's see if the Id is still valid
                                $propValue = [System.BitConverter]::ToString($extendedProperty.Value)
                                LogVerbose "LastActiveParentID: $propValue"

                                # Last active folder Id is the PidTagSourceKey value of the folder
                                if ($script:FoldersBySourceKey.Contains($propValue))
                                {
                                    $moveToFolder = $script:FoldersBySourceKey[$propValue]
                                    $targetFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId($moveToFolder)
                                    LogVerbose "lastActiveFolderId: $moveToFolder"
                                    break
                                }
                            }
                        }
                        if ( [String]::IsNullOrEmpty($moveToFolder) )
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
					        Log "Item was not a class enabled for recovery: $($Item.ItemClass)" Red
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
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
        return
	}
	
	$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )

    if ($Archive -or $ArchiveRoot)
    {
        if ([String]::IsNullOrEmpty($RestoreToFolder))
        {
            Log "When restoring from archive, -RestoreToFolder must be specified" Red
            return
        }
        if ($Archive)
        {
            $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx )
        }
        elseif ($ArchiveRoot)
        {
            $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot, $mbx )
        }
    }
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