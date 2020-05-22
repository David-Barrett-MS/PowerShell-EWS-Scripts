#
# Search-MailboxForMessageClass.ps1
#
# By David Barrett, Microsoft Ltd. 2013-2020. Use at your own risk.  No warranties are given.
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
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,
	
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Specifies the message class of the items being searched")]
	[ValidateNotNullOrEmpty()]
	[string]$MessageClass,

	[Parameter(Mandatory=$False,HelpMessage="Specifies the new message class that will be applied to the items (note that you cannot change the base item class of an item)")]
	[ValidateNotNullOrEmpty()]
	[string]$NewMessageClass,

    [Parameter(Mandatory=$False,HelpMessage="Adds the given property(ies) to the list of those that will be retrieved for an item (must be supplied as hash table @{})")]
    $ViewProperties,
	
	[Parameter(Mandatory=$False,HelpMessage="If this switch is specified, items will be searched for in the archive mailbox (otherwise, the main mailbox is searched)")]
    [alias("SearchArchive")]
	[switch]$Archive,

	[Parameter(Mandatory=$False,HelpMessage="If this switch is specified, items will be deleted")]
	[switch]$DeleteItems,
	
	[Parameter(Mandatory=$False,HelpMessage="If this switch is specified, items will be hard-deleted (otherwise, they'll be moved to Deleted Items)")]
	[switch]$HardDelete,
	
	[Parameter(Mandatory=$False,HelpMessage="If this switch is specified, only associated items will be searched (these are hidden messages within the folder)")]
	[switch]$AssociatedItemsOnly,
	
    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.")]
    $IncludeFolderList,
    
	[Parameter(Mandatory=$False,HelpMessage="If this switch is specified, then subfolders of any specified folder will also be searched")]
	[switch]$ProcessSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be excluded (these folders will not be searched)")]
    $ExcludeFolderList,
    
	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
	[Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts) - this requires the ADAL dlls to be available")]
	[switch]$OAuth,
	
	[Parameter(Mandatory=$False,HelpMessage="The client (application) Id that this script will identify as.  Must be registered in Azure AD.")]
    [alias("OAuthAppId")]
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
	
	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile,

	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",

	[Parameter(Mandatory=$False,HelpMessage="Throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected")]	
	[int]$ThrottlingDelay = 0,

	[Parameter(Mandatory=$False,HelpMessage="Batch size (number of items batched into one EWS request) - this will be decreased if throttling is detected")]	
	[int]$BatchSize = 200
)
$script:ScriptVersion = "1.1.4"

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
            LogVerbose "Tracing to: $TraceFile"
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

    if (!$script:currentThrottlingDelay)
    {
        # No throttling delay currently set, so we'll explicitly set it to zero (to prevent any errors with Start-Sleep)
        $script:currentThrottlingDelay = 0
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
    $script:service = $exchangeService
    return $exchangeService
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

Function InitPropList()
{
    # We need to convert the properties to EWS extended properties
    if ($script:itemPropsEws -eq $Null)
    {
        Write-Verbose "Building list of properties to retrieve"
        $script:property = @()
        foreach ($property in $ViewProperties)
        {
            $propdef = $null

            if ($property.StartsWith("{"))
            {
                # Property definition starts with a GUID, so we expect one of these:
                # {GUID}/name/mapiType - named property
                # {GUID]/id/mapiType   - MAPI property (shouldn't be used when accessing named properties)

                $propElements = $property -Split "/"
                if ($propElements.Length -eq 2)
                {
                    # We expect three elements, but if there are two it most likely means that the MAPI property Id includes the Mapi type
                    if ($propElements[1].Length -eq 8)
                    {
                        $propElements += $propElements[1].Substring(4)
                        $propElements[1] = [Convert]::ToInt32($propElements[1].Substring(0,4),16)
                    }
                }
                $guid = New-Object Guid($propElements[0])
                $propType = EWSPropertyType($propElements[2])

                try
                {
                    $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($guid, $propElements[1], $propType)
                }
                catch {}
            }
            else
            {
                # Assume MAPI property
                if ($property.ToLower().StartsWith("0x"))
                {
                    $property = $deleteProperty.SubString(2)
                }
                $propId = [Convert]::ToInt32($deleteProperty.SubString(0,4),16)
                $propType = EWSPropertyType($deleteProperty.SubString(5))

                try
                {
                    $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
                }
                catch {}
            }

            if ($propdef -ne $Null)
            {
                $script:property += $propdef
                Write-Verbose "Added property $property to list of those to retrieve"
            }
            else
            {
                Log "Failed to parse (or convert) property $property" Red
            }
        }
    }
}

Function GetMaxCount( [String]$DictionaryType )
{
    if ($DictionaryType.EndsWith("PhoneNumberDictionary")) { return 18 }
    return 2
}

$script:excludedProperties = @("Schema","Service","IsDirty","IsAttachment","IsNew")
Function StoreFriendlyData( $item )
{
    # Process this object so that the data is presented in friendly format (when piped to Export-CSV, for example)

    $prettyItem = New-Object PsObject
    $item.PsObject.Properties | foreach {
        if ( !$script:excludedProperties.Contains($_.Name) )
        {
            $value = $_.Value
            if ($value -ne $null)
            {
                try
                {
                    $objectType = $_.Value.GetType().BaseType.ToString()
                    if ($objectType.Equals("System.Array"))
                    {
                        # This is an array
                        $value = ""
                        for ($i=0; $i -le $_.Value.Length; $i++)
                        {
                            if ($i -gt 0) { if (![String]::IsNullOrEmpty($value)) { $value += ";" } }
                            if (![String]::IsNullOrEmpty($_.Value[$i])) { $value += $_.Value[$i] }
                        }
                    }
                    elseif ( $objectType.Equals("Microsoft.Exchange.WebServices.Data.AttachmentCollection") -or $objectType.Equals("Microsoft.Exchange.WebServices.Data.ComplexPropertyCollection`1[Microsoft.Exchange.WebServices.Data.Attachment]") )
                    {
                        # List the attachments
                        $value = ""
                        for ($i=0; $i -lt $_.Value.Count; $i++)
                        {
                            $attach = $_.Value[$i]
                            if (![String]::IsNullOrEmpty($attach.Name))
                            {
                                if (![String]::IsNullOrEmpty($value)) { $value +=";" }
                                $value += $attach.Name
                            }
                        }
                    }
                    elseif ( $objectType.Contains("Dictionary") -or $objectType.StartsWith("Microsoft.Exchange.WebServices.Data.ComplexPropertyCollection"))
                    {
                        # Generic handling for EWS Dictionary objects
                        $i = 0
                        $value = ""
                        $maxCount = GetMaxCount($_.Value.ToString())
                        for ($i=0; $i -le $maxCount; $i++)
                        {
                            if (![String]::IsNullOrEmpty($_.Value[$i]))
                            {
                                if (![String]::IsNullOrEmpty($value)) { $value +=";" }
                                $value += $_.Value[$i]
                            }
                        }
                    }
                    else
                    {
                        #Write-Host "$($_.Name) : $objectType" -ForegroundColor Gray
                    }
                }
                catch {}
                $prettyItem | Add-Member -MemberType NoteProperty -Name $_.Name -Value $value
            }
        }
    }
    $prettyItem
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
                if ( ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed") -or ($results[$i].ErrorCode -eq "ErrorInvalidOperation") -or ($results[$i].ErrorCode -eq "ErrorItemNotFound") )
                {
                    # This is a permanent error, so we remove the item from the list
                    [void]$Items.Remove($requestedItems[$i])
                    if (!$suppressErrors)
                    {
                        Log "Permanent error $($results[$i].ErrorCode) ($($results[$i].MessageText)) reported for item: $($requestedItems[$i].UniqueId)" Red
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
                            Log "Permanent error $($results[$i].ErrorCode) ($($results[$i].MessageText)) reported for item: $($requestedItems[$i].UniqueId)" Red
                        }
                    }
                }
                $failed++
            } 
        }
    }
    if ( ($failed -gt 0) -and !$suppressErrors )
    {
        Log "$failed items reported error during batch request (if throttled, some failures are expected)" Yellow
    }
}

Function ThrottledBatchDelete()
{
    # Send request to move/copy items, allowing for throttling
    param (
        $ItemsToDelete,
        $BatchSize = 200,
        $SuppressNotFoundErrors = $false
    )

    if ($script:MaxBatchSize -gt 0)
    {
        # If we've had to reduce the batch size previously, we'll start with the last size that was successful
        $BatchSize = $script:MaxBatchSize
    }

    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
    if ($HardDelete)
    {
        $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
    }

    $progressActivity = "Deleting items"
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $totalItems = $ItemsToDelete.Count
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    if ( $totalItems -gt 10000 )
    {
        if ( $script:currentThrottlingDelay -lt 1000 )
        {
            $script:currentThrottlingDelay = 1000
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

			$results = $script:service.DeleteItems( $deleteIds, $deleteMode, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
            Start-Sleep -Milliseconds $script:currentThrottlingDelay
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

Function ThrottledBatchUpdate()
{
    # Send request to move/copy items, allowing for throttling
    param (
        $UpdatedItems,
        $BatchSize = 200
    )

    if ($script:MaxBatchSize -gt 0)
    {
        # If we've had to reduce the batch size previously, we'll start with the last size that was successful
        $BatchSize = $script:MaxBatchSize
    }

    $progressActivity = "Updating items"  
    $genericItemList = [System.Collections.Generic.List``1].MakeGenericType([Microsoft.Exchange.WebServices.Data.Item])
      
    $finished = $false
    $totalItems = $UpdatedItems.Count
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    if ( $totalItems -gt 10000 )
    {
        if ( $script:currentThrottlingDelay -lt 1000 )
        {
            $script:currentThrottlingDelay = 1000
            LogVerbose "Large number of items will be processed, so throttling delay set to 1000ms"
        }
    }
    $consecutiveErrors = 0

    while ( !$finished )
    {
	    $updateBatch = [Activator]::CreateInstance($genericItemList)
        
        for ([int]$i=0; $i -lt $BatchSize; $i++)
        {
            if ($UpdatedItems[$i] -ne $null)
            {
                $updateBatch.Add($UpdatedItems[$i])
            }
            if ($i -ge $UpdatedItems.Count)
                { break }
        }

        $results = $null
        try
        {
            LogVerbose "Sending batch request to update $($updateBatch.Count) items ($($UpdatedItems.Count) remaining)"

			$results = $script:service.UpdateItems( $updateBatch, $null, [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, [Microsoft.Exchange.WebServices.Data.MessageDisposition]::SaveOnly, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone )
            Start-Sleep -Milliseconds $script:currentThrottlingDelay
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

        RemoveProcessedItemsFromList $updateBatch $results $false $UpdatedItems

        $percentComplete = ( ($totalItems - $UpdatedItems.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($UpdatedItems.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
}

Function InitLists()
{
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId])
    $script:ItemsToDelete = [Activator]::CreateInstance($genericItemIdList)

    $genericItemList = [System.Collections.Generic.List``1].MakeGenericType([Microsoft.Exchange.WebServices.Data.Item])
    $script:ItemsToUpdate = [Activator]::CreateInstance($genericItemList)
}

Function ProcessItem( $item )
{
	# We have found an item, so this function handles any processing
	# In this case, we are simply going to log a few details

    # We now add this item to our collection of found items (so that we can export)
    if ($script:itemPropsEws)
    {
        $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        if ($script:itemPropsEws.Length -gt 0)
        {
            # We have additional properties to retrieve, so we reload the item asking for first class properties and all the additional ones
            $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, $script:itemPropsEws)
        }
        $item.Load($propset)
    }

    StoreFriendlyData $item

	if ($DeleteItems)
	{
		LogVerbose "Adding item to delete list: $($item.Subject)"
        $script:ItemsToDelete.Add($item.Id)
        return # If we are deleting an item, then no other updates are relevant
	}

    if ( ![String]::IsNullOrEmpty($NewMessageClass) )
    {
        # We need to update the message class
        if ($item.ItemClass -ne $NewMessageClass)
        {
		    LogVerbose "Updating item class from $($item.ItemClass) to $($NewMessageClass): $($item.Subject)"
            $item.ItemClass = $NewMessageClass
            $script:ItemsToUpdate.Add($item)
        }
    }

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

    if ($FolderPath.ToLower().StartsWith("wellknownfoldername"))
    {
        # Well known folder, so bind to it directly
        $wkf = $FolderPath.SubString(20)
        LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind($folderId)
        return $Folder
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
                    Start-Sleep -Milliseconds $script:currentThrottlingDelay
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

Function SearchMailbox()
{
    Log ([string]::Format("Processing mailbox {0}", $Mailbox)) Gray

	$script:service = CreateService($Mailbox)
    if ( $script:service -eq $null) { return }

	# Set our root folder
    $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
    if ($Archive)
    {
        $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot
    }

    InitPropList

    if (!($IncludeFolderList))
    {
        # No folders specified to search, so we start with mailbox root
        $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId( $rootFolderId )	
	    SearchFolder $FolderId
    }
    else
    {
        # We are searching specific folders
        $rootFolder = ThrottledFolderBind $rootFolderId
        foreach ($includedFolder in $IncludeFolderList)
        {
            $folder = $null
            $folder = GetFolder($rootFolder, $includedFolder, $false)

            if ($folder)
            {
                $folderPath = GetFolderPath($folder)
                Log "Starting search in: $folderPath"                
	            SearchFolder $folder.Id
            }
        }
    }
}

Function IncreaseThrottlingDelay()
{
    # Increase our throttling delay to try and avoid throttling (we only increase to a maximum delay of 10 seconds between requests)
    $maxThrottlingDelay = 5000
    if ( $script:currentThrottlingDelay -eq $maxThrottlingDelay) { return }

    if ( $script:currentThrottlingDelay -lt $maxThrottlingDelay)
    {
        if ($script:currentThrottlingDelay -lt 1)
        {
            $script:currentThrottlingDelay = $ThrottlingDelay
            if ($script:currentThrottlingDelay -lt 1) 
            {
                # In case throttling delay parameter is set to 0, or a silly value
                $script:currentThrottlingDelay = 100
            }
        }
        else
        {
            $script:currentThrottlingDelay = $script:currentThrottlingDelay * 2
        }
        if ( $script:currentThrottlingDelay -gt $maxThrottlingDelay)
        {
            $script:currentThrottlingDelay = $maxThrottlingDelay
        }
    }
    LogVerbose "Updated throttling delay to $($script:currentThrottlingDelay)ms"
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
    }

    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
    $parentFolder = ThrottledFolderBind $Folder.Id $propset $Folder.Service
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
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset $Folder.Service
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function GetWellKnownFolderPath($WellKnownFolder)
{
    if (!$script:wellKnownFolderCache)
    {
        $script:wellKnownFolderCache = @{}
    }

    if ($script:wellKnownFolderCache.ContainsKey($WellKnownFolder))
    {
        return $script:wellKnownFolderCache[$WellKnownFolder]
    }

    $folder = $null
    $folderPath = $null
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service,$WellKnownFolder)
    if ($folder)
    {
        $folderPath = GetFolderPath($folder)
        LogVerbose "GetWellKnownFolderPath: Path for $($WellKnownFolder): $folderPath"
    }
    $script:wellKnownFolderCache.Add($WellKnownFolder, $folderPath)
    return $folderPath
}

Function IsFolderExcluded()
{
    # Return $true if folder is in the excluded list

    param ($folderPath)

    # To support localisation, we need to handle WellKnownFolderName enumeration
    # We do this by putting all our excluded folders into a hash table, and checking that we have the full path for any well known folders (which we retrieve from the mailbox)
    if ($script:excludedFolders -eq $null)
    {
        # Create and build our hash table
        $script:excludedFolders = @{}

        if ($ExcludeFolderList)
        {
            LogVerbose "Building folder exclusion list"#: $($ExcludeFolderList -join ',')"
            ForEach ($excludedFolder in $ExcludeFolderList)
            {
                $excludedFolder = $excludedFolder.ToLower()
                $wkfStart = $excludedFolder.IndexOf("wellknownfoldername")
                LogVerbose "Excluded folder: $excludedFolder"
                if ($wkfStart -ge 0)
                {
                    # Replace the well known folder name with its full path
                    $wkfEnd = $excludedFolder.IndexOf("\", $wkfStart)-1
                    if ($wkfEnd -lt 0) { $wkfEnd = $excludedFolder.Length }
                    $wkf = $null
                    $wkf = $excludedFolder.SubString($wkfStart+20, $wkfEnd - $wkfStart - 19)
                    
                    $wellKnownFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf
                    $wellKnownFolderPath = GetWellKnownFolderPath($wellKnownFolder)

                    $excludedFolder = $excludedFolder.Substring(0, $wkfStart) + $wellKnownFolderPath + $excludedFolder.Substring($wkfEnd+1)
                    LogVerbose "Path of excluded folder: $excludedFolder"
                }
                $script:excludedFolders.Add($excludedFolder, $null)
            }
        }
    }

    return $script:excludedFolders.ContainsKey($folderPath.ToLower())
}

Function SearchFolder( $FolderId )
{
	# Bind to the folder and show which one we are processing
    $folder = $null
    try
    {
	    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service,$FolderId)
    }
    catch {}
    ReportError "SearchFolder"
    if ($folder -eq $null) { return }

    $folderPath = GetFolderPath($folder)

    if (IsFolderExcluded($folderPath))
    {
        LogVerbose "Folder excluded: $folderPath"
        return
    }
	Log "Processing folder: $folderPath"
    InitLists

	# Search the folder for any matching items
	$pageSize = 1000 # We will get details for up to 1000 items at a time
	$offset = 0
	$moreItems = $true
	
	# Perform the search and display the results
	while ($moreItems)
	{
		$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
		
		$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
		$view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
		$view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow

        if ($AssociatedItemsOnly)
        {
            LogVerbose "Searching for associated items only"
            $view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        }
		
		$results = $service.FindItems( $FolderId, $searchFilter, $view )
		
		ForEach ($item in $results.Items)
		{
			ProcessItem $item
		}
		
		$moreItems = $results.MoreAvailable
		$offset += $pageSize
	}

    if ($script:ItemsToDelete.Count -gt 0)
    {
        # Delete the items we found in this folder
        ThrottledBatchDelete $script:ItemsToDelete -SuppressNotFoundErrors $true
    }

    if ($script:ItemsToUpdate.Count -gt 0)
    {
        # Delete the items we found in this folder
        ThrottledBatchUpdate $script:ItemsToUpdate
    }
	
	# Now search subfolders
    if ($ProcessSubfolders)
    {
	    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
        $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
	    ForEach ($subFolder in $folder.FindFolders($view))
	    {
		    SearchFolder $subFolder.Id $folderPath
	    }
    }
}



# The following is the main script


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

$script:searchResults = @()
  
# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
	$csv = Import-CSV $Mailbox
	foreach ($entry in $csv)
	{
		$Mailbox = $entry.PrimarySmtpAddress
		if ( ![string]::IsNullOrEmpty($Mailbox) )
		{
			if ($Mailbox.Contains("@")) { SearchMailbox }
		}
	}
}
Else
{
	# Process as single mailbox
	SearchMailbox
}
