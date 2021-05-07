#
# Update-FolderItems.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2019. Use at your own risk.  No warranties are given.
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

    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox will be accessed (instead of the main mailbox)")]
    [switch]$Archive,
		
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder")]
    [switch]$PublicFolders,

    [Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox message root folder is assumed.")]
    $FolderPath,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, subfolders will also be processed")]
    [switch]$ProcessSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="Adds the given property(ies) to the item(s) (must be supplied as hash table @{})")]
    $AddItemProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the given property(ies) from the item(s)")]
    $DeleteItemProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Marks the item(s) as read")]
    [switch]$MarkAsRead,
    
    [Parameter(Mandatory=$False,HelpMessage="Marks the item(s) as unread")]
    [switch]$MarkAsUnread,

    [Parameter(Mandatory=$False,HelpMessage="Actions will only apply to contact objects that have the given SMTP address as their email address.  Supports multiple SMTP addresses passed as an array.")]
    $MatchContactAddresses,

    [Parameter(Mandatory=$False,HelpMessage="If any matching contact object contains a contact photo, the photo is deleted")]
    [switch]$DeleteContactPhoto,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the item(s)")]
    [switch]$Delete,
    
    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given AQS filter will be processed `r`n(see https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-an-aqs-search-by-using-ews-in-exchange")]
    [string]$SearchFilter,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that have values in the given properties will be updated.")]
    $PropertiesMustExist,

    [Parameter(Mandatory=$False,HelpMessage="If specified, only items that match the given values in the given properties will be updated.  Properties must be supplied as a Dictionary @{""propId"" = ""value""}")]
    $PropertiesMustMatch,

    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credentials")]
    [System.Management.Automation.PSCredential]$Credential,
				
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
	
    [Parameter(Mandatory=$False,HelpMessage="If specified, only TLS 1.2 connections will be negotiated")]
    [switch]$ForceTLS12,
	
    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
    [string]$EWSManagedApiPath = "",
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
    [switch]$IgnoreSSLCertificate,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
    [switch]$AllowInsecureRedirection,
	
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
    [string]$TraceFile,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, no items will actually be deleted (but any processing that would occur will be logged)")]	
    [switch]$WhatIf
)
$script:ScriptVersion = "1.1.3"

if ($ForceTLS12)
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
else
{
    Write-Host "If having connection/auth issues for Exchange Online or hybrid, you may need -ForceTLS12 switch" -ForegroundColor Yellow
}

# Define our functions

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
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
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
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

function GetOAuthCredentials()
{
    # Obtain OAuth token for accessing mailbox
    $exchangeCredentials = $null

    if ( $(LoadADAL) -eq $false )
    {
        Log "Failed to load ADAL, which is required for OAuth" Red
        Exit
    }

    $authenticationContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.windows.net/common", $False)
    $platformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
    $redirectUri = New-Object Uri($OAuthRedirectUri)
    $authenticationResult = $authenticationContext.AcquireTokenAsync("https://outlook.office365.com", $OAuthClientId, $redirectUri, $platformParameters)

    if ( !$authenticationResult.IsFaulted )
    {
        $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($authenticationResult.Result.AccessToken)
        $Mailbox = $authenticationResult.Result.UserInfo.UniqueId
        LogVerbose "OAuth completed for $($authenticationResult.Result.UserInfo.DisplayableId)"
    }

    return $exchangeCredentials
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
        return $result.Properties["mail"]
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

    # Attach the trace listener to the Exchange service
    $service.TraceListener = $script:Tracer
}

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
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    return $null
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
        if ($Credential -ne $Null)
        {
            LogVerbose "Applying given credentials: $($Credential.UserName)"
            $exchangeService.Credentials = $Credential.GetNetworkCredential()
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
    if ($Impersonate)
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
        # Assume MAPI property
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
        $props = @{}

        foreach ($PropDef in $PropertiesMustMatch.Keys)
        {
            $EWSPropDef = GenerateEWSProp($PropDef)
            if ($EWSPropDef -ne $Null)
            {
                $props.Add($EWSPropDef, $PropertiesMustMatch[$PropDef])
            }

        }
        $script:propertiesMustMatchEws = $props
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
    $contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($item.Service, $item.Id, $propset)

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
            if (![String]::IsNullOrEmpty(($requiredProperty.PropertySetId)))
            {
                foreach ($itemProperty in $item.ExtendedProperties)
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
            }
            if (!$propMatches)
            {
                Write-Verbose "$requiredProperty does not match, ignoring item"
                return $false
            }
        }
    }
    return $true    
}

Function ProcessItem()
{
	# Apply updates to the given item

    $item = $args[0]
	if ($item -eq $null)
	{
		throw "No item specified"
	}
    
    if ( -not (ItemHasRequiredProperties($item)) -or -not (ItemPropertiesMatchRequirements($item)) ) { return }

    LogVerbose "Processing item: $($item.Subject)"

    # Check for delete first of all
    if ($Delete)
    {
        if (-not $WhatIf)
        {
            [void]$script:itemsToDelete.Add($item.Id)
            Log "$($item.Subject) added to list of items to be deleted" Gray
        }
        else
        {
            Log "$($item.Subject) would be deleted" Gray
        }
        $script:itemsAffected++
        return # If Delete is specified, any other parameter is irrelevant
    }

    $madeChanges = $false

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
                if ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed")
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
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	#$baseList = [System.Collections.Generic.List``1]
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $totalItems = $ItemsToDelete.Count
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    if ( $totalItems -gt 10000 )
    {
        if ( $script:throttlingDelay -lt 1000 )
        {
            $script:throttlingDelay = 1000
            LogVerbose "Large number of items will be processed, so throttling delay set to 1000ms"
        }
    }
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
			$results = $script:service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
            Start-Sleep -Milliseconds $script:throttlingDelay
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

        RemoveProcessedItemsFromList $deleteIds $results $ItemsToDelete

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
}


Function ProcessFolder()
{
	# Process all items within this folder

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

    $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,[Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead,[Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,[Microsoft.Exchange.WebServices.Data.ContactSchema]::HasPicture)
    if ($script:deleteItemPropsEws -ne $null) # We retrieve any properties that we want to delete
    {
        foreach ($deleteProperty in $script:deleteItemPropsEws)
        {
            $propSet.Add($deleteProperty)
        }
    }
    if ($script:propertiesMustExistEws -ne $null) # We retrieve any properties that must exist
    {
        foreach ($requiredProperty in $script:propertiesMustExistEws)
        {
            $propSet.Add($requiredProperty)
        }
    }
    if ($script:propertiesMustMatchEws -ne $null) # We retrieve any properties that we need to check the value of
    {
        foreach ($propMustMatch in $script:propertiesMustMatchEws.Keys)
        {
            $propSet.Add($propMustMatch)
        }
    }

    LogVerbose "Building list of items"
    if ($MatchContactAddresses)
    {
        $filters = @()
        foreach ($contactAddress in $MatchContactAddresses)
        {
            LogVerbose "Adding SMTP address search: $smtpAddress"
            $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress1, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
            $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress2, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)
            $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress3, $contactAddress, 
                [Microsoft.Exchange.WebServices.Data.ContainmentMode]::FullString, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)

        }
        $SearchFilter = $Null
        if ( $filters.Count -gt 0 )
        {
            $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
            foreach ($filter in $filters)
            {
                $SearchFilter.Add($filter)
            }
        }        
    }
    elseif (![String]::IsNullOrEmpty($SearchFilter))
    {
        LogVerbose "Search query being applied: $SearchFilter"
    }

    Write-Progress -Activity "$progressActivity reading items" -Status "0 items found" -PercentComplete -1
	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = $propSet

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
            Start-Sleep -Milliseconds $script:throttlingDelay
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
                $itemsToProcess += $item
		    }
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
    ForEach ($item in $itemsToProcess)
    {
        ProcessItem $item
        $i++
        if ($i%10 -eq 0)
        {
            Write-Progress -Activity "$progressActivity processing items" -Status "$i items processed" -PercentComplete (($i/$itemsToProcess.Count)*100)
        }
    }
    Write-Progress -Activity "$progressActivity processing items" -Status "Complete" -Completed

    if ($script:itemsToDelete.Count -gt 0)
    {
        ThrottledBatchDelete $script:itemsToDelete
    }
}

function ProcessMailbox()
{
    # Process the mailbox

    # Parse any properties that we want to delete - this is because we need to retrieve them first
    DeleteItemProperties $null

    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Write-Host "Failed to create ExchangeService" -ForegroundColor Red
	}

    $script:throttlingDelay = 0

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
                Write-Verbose "Attempting to bind to well known folder: $wkf"
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
                $Folder = ThrottledFolderBind($folderId)
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
    if ($WhatIf)
    {
        Log "$($Mailbox): $($script:itemsAffected) item(s) would be affected (but -WhatIf was specified so no action was taken)"
    }
    else
    {
        Log "$($Mailbox): $($script:itemsAffected) item(s) affected"
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

  

Write-Host ""
CreatePropLists

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