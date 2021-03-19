#
# Update-Folders.ps1
#
# By David Barrett, Microsoft Ltd. 2016. Use at your own risk.  No warranties are given.
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

	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, public folders will be processed")]
	[switch]$PublicFolders,

	[Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox is accessed (instead of the main mailbox)")]
	[switch]$Archive,
		
	[Parameter(Mandatory=$False,HelpMessage="If specified, folder or list of folders to process.  If omitted, root message folder is assumed.")]
	$FolderPaths = "",

    [Parameter(Mandatory=$False,HelpMessage="Specifies the folder(s) to be excluded")]
    [ValidateNotNullOrEmpty()]
    $ExcludedFolderPaths,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is set, only those folders that specifically match the exclusion list will be excluded (subfolders will still be processed)")]
    [ValidateNotNullOrEmpty()]
    [switch]$DoNotExcludeSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="If this switch is set, search folders will also be processed (by default they are excluded)")]
    [ValidateNotNullOrEmpty()]
    [switch]$IncludeSearchFolders,

    [Parameter(Mandatory=$False,HelpMessage="Changes the display name of the folder(s)")]
    $NewDisplayName,
    
    [Parameter(Mandatory=$False,HelpMessage="Any characters defined here will be removed from any folder names use an array of characters, e.g. @('.')")]
    $RemoveCharactersFromDisplayName,
    
    [Parameter(Mandatory=$False,HelpMessage="Sets the class of the folder(s) to that specified (e.g. IPF.Note)")]
    $FolderClass,
    
    [Parameter(Mandatory=$False,HelpMessage="If specified, any folders that do not have an item class defined (i.e. it is empty) will have the item class set.  If -FolderClass is specified, all blank folder are set to that - otherwise, the class is set to the same as the parent folder.")]
    [switch]$RepairFolderClass,
    
	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, subfolders will also be processed")]
	[switch]$ProcessSubfolders,
	
    [Parameter(Mandatory=$False,HelpMessage="Adds the given properties (must be supplied as hash table @{}) to the folder(s)")]
    $AddFolderProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the given properties from the folder(s)")]
    $DeleteFolderProperties,
    
    [Parameter(Mandatory=$False,HelpMessage="Deletes the folder(s)")]
    [switch]$Delete,
    
    [Parameter(Mandatory=$False,HelpMessage="Purges (empties) the folder(s).  This parameter is required if you want to delete folders that have messages in them.")]
    [switch]$Purge,
    
    [Parameter(Mandatory=$False,HelpMessage="Purges (empties) the folder(s).  This switch will force a hard-delete of any items found in the folder (otherwise soft-delete is used).  Can only be used with -Purge (both switches are required)")]
    [switch]$HardPurge,
    
    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
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

	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]
	[string]$TraceFile,

	[Parameter(Mandatory=$False,HelpMessage="If specified, changes will be made to the mailbox (but actions that would be taken will be logged)")]	
	[switch]$WhatIf
)
$script:ScriptVersion = "1.0.3"

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

Function LogFolderAction($Folder, [String]$Details, [ConsoleColor]$Colour)
{
    $folderPath = GetFolderPath $Folder
    if ($WhatIf)
    {
        Log "[WHATIF]$($folderPath): $Details" $Colour
    }
    else
    {
        Log "$($folderPath): $Details" $Colour
    }
}

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
        return $mail
    }
    return $null
}

Function TrustAllCerts()
{
    # Trust all SSL certificates
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
        # Back off for the time given by the server
        Log "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
        Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text"
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
    if ($propset -eq $null)
    {
        $propset = $script:requiredFolderProperties
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

Function GetFolder()
{
	# Return a reference to a folder specified by path
    param (
        $RootFolder,
        [string]$FolderPath,
        [bool]$Create
    )	
	
    if ( ($RootFolder -eq $null) -and (!$FolderPath.ToLower().StartsWith("wellknownfoldername")) )
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
            if ( ![String]::IsNullOrEmpty($FolderPath) )
            {
                LogVerbose "[GetFolder]Remainder of path to match: $FolderPath"
            }
            else
            {
                return $RootFolder
            }
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

Function HexStringToByteArray {
    param(
        [String] $HexString
    )

    $Bytes = [byte[]]::new($HexString.Length / 2)

    For($i=0; $i -lt $HexString.Length; $i+=2){
        $Bytes[$i/2] = [convert]::ToByte($HexString.Substring($i, 2), 16)
    }

    return $Bytes
}

Function AddFolderProperties($folder)
{
    # Add the specified properties to the folder

    # First of all ensure we have some properties to add...
    if ($AddFolderProperties -eq $Null) { return }

    # We need to convert the properties to EWS extended properties
    if ($script:addFolderPropsEws -eq $Null)
    {
        Write-Verbose "Building list of properties to add"
        $script:addFolderPropsEws = @{}
        foreach ($addProperty in $AddFolderProperties.Keys)
        {
            $value = $AddFolderProperties[$addProperty]
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
                if ($propType -eq [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
                {
                    [byte[]]$bytes = HexStringToByteArray $value
                    $script:addFolderPropsEws.Add($propdef, $bytes)
                }
                else
                {
                    $script:addFolderPropsEws.Add($propdef, $value)
                }
                Write-Verbose "Added property $addProperty to add list; value is $($script:addFolderPropsEws[$propdef])"
            }
            else
            {
                Log "Failed to parse (or convert) property $addProperty" Red
            }
        }
    }

    # Now we add the properties to the folder
    foreach ($addProperty in $script:addFolderPropsEws.Keys)
    {

        $folder.SetExtendedProperty($addProperty, $script:addFolderPropsEws[$addProperty])
    }

    # Now update the folder
    try
    {
        if (!$WhatIf)
        {
            $folder.Update()
        }
        LogFolderAction $folder "added properties" Green
    }
    catch
    {
        LogFolderAction "Error during update: $($Error[0])" Red
        $folder.Load() # We do this to clear the invalid update from the folder object
    }
}

Function DeleteFolderProperties($folder)
{
    # Delete any properties specified

    # Check if we are deleting any properties
    if ($DeleteFolderProperties -eq $Null) { return }

    # We need to convert the properties to EWS extended properties
    if ($script:deleteFolderPropsEws -eq $Null)
    {
        Write-Verbose "Building list of properties to delete"
        $script:deleteFolderPropsEws = @()
        foreach ($deleteProperty in $DeleteFolderProperties)
        {
            if ($deleteProperty.ToLower().StartsWith("0x"))
            {
                $deleteProperty = $deleteProperty.SubString(2)
            }
            $propId = [Convert]::ToInt32($deleteProperty.SubString(0,4),16)
            $propType = EWSPropertyType($deleteProperty.SubString(5))

            $propdef = $Null
            try
            {
                $propdef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
            }
            catch {}
            if ($propdef -ne $Null)
            {
                $script:deleteFolderPropsEws += $propdef
                Write-Verbose "Added property $deleteProperty to delete list"
            }
            else
            {
                Log "Failed to parse (or convert) property $deleteProperty" Red
            }
        }
    }

    $loadPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet($script:deleteFolderPropsEws)
    $loadPropSet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
    $folder.Load($loadPropSet)

    Write-Verbose "Deleting properties from $($folder.DisplayName)"

    # Process each of the properties and add delete request
    $foundProperties = @()
    foreach ($deleteProperty in $script:deleteFolderPropsEws)
    {
        $found = $False
        foreach ($extendedProperty in $folder.ExtendedProperties)
        {
            if ( ($deleteProperty.Tag -eq $extendedProperty.PropertyDefinition.Tag) -and ($deleteProperty.MapiType -eq $extendedProperty.PropertyDefinition.MapiType) )
            {
                $foundProperties += $deleteProperty
                $found = $True
            }
        }
        
        if (!$found)
        {
            $propTag = [String]::Format("0x{0}", [Convert]::ToString($deleteProperty.Tag, 16))
            LogFolderAction $folder "Property $propTag was not found" Red
        }
    }

    if ($foundProperties.Count -gt 0)
    {
        Write-Verbose "Applying changes"
        foreach ($foundProperty in $foundProperties)
        {
            $propTag = [String]::Format("0x{0}", [Convert]::ToString($foundProperty.Tag, 16))
            if ( $folder.RemoveExtendedProperty($foundProperty) )
            {
                LogFolderAction $folder "Property $propTag was found and added to delete request" Gray
            }
            else
            {
                LogFolderAction $folder "Property $propTag was found, but cannot be removed" Red
            }
        }

        # Now update the folder
        try
        {
            if (!$WhatIf)
            {
                $folder.Update()
            }
            LogFolderAction $folder "properties deleted" Green
        }
        catch
        {
            LogFolderAction $folder "Error while updating: $($Error[0])" Red
            $folder.Load($loadPropSet) # We do this to clear the invalid update from the folder object
        }
    }
    else
    {
        LogFolderAction $folder "No properties to delete, folder not updated" Green
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
                if ( ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed") -or ($results[$i].ErrorCode -eq "ErrorInvalidOperation") -or ($results[$i].ErrorCode -eq "ErrorCannotDeleteObject") )
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
    # Send request to delete items, allowing for throttling
    param (
        $ItemsToDelete
    )

    if ($script:currentBatchSize -lt 1) { $script:currentBatchSize = $BatchSize }

	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType(@($itemIdType))
    
    $finished = $false
    $progressActivity = "Deleting items"
    $totalItems = $ItemsToDelete.Count
    LogVerbose "Batch deleting $totalItems items"
    Write-Progress -Activity $progressActivity -Status "0% complete" -PercentComplete 0

    $deleteType = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
    if ($HardPurge)
    {
        LogVerbose "Setting delete type to HardDelete"
        $deleteType = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
    }

    if ( $totalItems -gt 10000 )
    {
        if ( $script:currentThrottlingDelay -lt 5000 )
        {
            $script:currentThrottlingDelay = 5000
            LogVerbose "Large number of items will be processed, so throttling delay set to 5000ms"
        }
    }

    while ( !$finished )
    {
	    $deleteIds = [Activator]::CreateInstance($genericItemIdList)
        
        for ([int]$i=0; $i -lt $script:currentBatchSize; $i++)
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
            if (!$WhatIf)
            {
			    $results = $script:service.DeleteItems( $deleteIds, $deleteType, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly)
                Sleep -Milliseconds $script:currentThrottlingDelay
            }
            else
            {
                # As we are running in WhatIf mode, we remove all items from the delete list
                RemoveProcessedItemsFromList $deleteIds $deleteIds $ItemsToDelete
            }
        }
        catch
        {
            if ( Throttled )
            {
                # We've been throttled, nothing to do here as the check for throttling will also have backed off
            }
            elseif ($Error[0].Exception.InnerException.ToString().Contains("The operation has timed out"))
            {
                # We've probably been throttled, so we'll reduce the batch size and try again
                if ($script:currentBatchSize -gt 50)
                {
                    LogVerbose "Timeout error received"
                }
                else
                {
                    $finished = $true
                }
            }
            else
            {
                $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
                $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
                $responseXml = [xml]$lastResponse
	            if ($responseXml.Trace.Envelope.Body.Fault.detail.ResponseCode.Value -eq "ErrorNoRespondingCASInDestinationSite")
                {
                    # We get this error if the destination CAS (request was proxied) hasn't returned any data within the timeout (usually 60 seconds)
                    # Reducing the batch size should help here, and we want to reduce it quite aggressively
                    if ($script:currentBatchSize -gt 50)
                    {
                        LogVerbose "ErrorNoRespondingCASInDestinationSite error received"
                    }
                    else
                    {
                        $finished = $true
                    }
                }
                else
                {
                    Log "Unexpected error: $($Error[0].Exception.InnerException.ToString())" Red
                    $finished = $true
                }
            }
        }

        RemoveProcessedItemsFromList $deleteIds $results $ItemsToDelete
        LogVerbose "$($totalItems - $ItemsToDelete.Count) items deleted"

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
            Write-Progress -Activity $progressActivity -Status "100% complete" -Completed
        }
    }
}

Function PurgeFolder($Folder)
{
    LogVerbose "Purging $($Folder.DisplayName)"

    if (!($script:service.RequestedServerVersion.ToString().Contains("2007")) -and ($script:service.RequestedServerVersion -ne [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010))
    {
        # Schemas later than 2010 support folder empty, so we try that first
        try
        {
            if (!$WhatIf)
            {
                $Folder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $False)
            }
            LogFolderAction $Folder "emptied" Green
            $Folder.Load()
            return
        }
	    catch
	    {
            if (ErrorReported)
            {
		        LogFolderAction $Folder "unable to Empty(), attempting manual deletion" Yellow
            }
	    }
    }

    # Delete all items from the folder

	$Offset = 0
	$PageSize = 1000 # We're only querying Ids, so 1000 items at a time is reasonable
	$MoreItems = $true

    # We create a list of all the items we need to move, and then batch move them later (much faster than doing it one at a time)
    $itemsToDelete = New-Object System.Collections.ArrayList
    $i = 0
	
    $progressActivity = "Reading items in folder $(GetFolderPath($Folder))"
    LogVerbose "Building list of items to delete from $($Folder.DisplayName)"
    Write-Progress -Activity $progressActivity -Status "0 items found" -PercentComplete -1

	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)

        $FindResults = $null
        try
        {
		    $FindResults = $Folder.FindItems($View)
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
                $MoreItems = $false
            }
        }
		
        if ($FindResults)
        {
		    ForEach ($item in $FindResults.Items)
		    {
                [void]$itemsToDelete.Add($item.Id)
		    }
		    $MoreItems = $FindResults.MoreAvailable
            if ($MoreItems)
            {
                LogVerbose "$($itemsToDelete.Count) items found so far, more available"
            }
		    $Offset += $PageSize
        }
        Write-Progress -Activity $progressActivity -Status "$($itemsToDelete.Count) items found" -PercentComplete -1
	}
    Write-Progress -Activity $progressActivity -Status "$($itemsToDelete.Count) items found" -Completed

    if ( $itemsToDelete.Count -gt 0 )
    {
        Log "$($itemsToDelete.Count) item(s) found; attempting to delete" Green
        ThrottledBatchDelete $itemsToDelete
    }
    else
    {
        Log "No items found in $($Folder.DisplayName)"
    }
}

function RemoveCharactersFromDisplayName($Folder)
{
    # Check for any characters in the folder name that we need to remove

    if (!$RemoveCharactersFromDisplayName)
        { return }

    [string]$displayName = $Folder.DisplayName
    foreach ($char in $RemoveCharactersFromDisplayName)
    {
        if ($displayName.Contains($char))
        {
            $displayName = $displayName.Replace($char, '')
        }
    }
    if ($displayName.Equals($Folder.DisplayName))
    {
        # No change
        LogVerbose "No characters found to replace in DisplayName"
        return
    }
    LogFolderAction $Folder "DisplayName changed to $displayName"
    if (!$WhatIf)
    {
        $Folder.DisplayName = $displayName
        $Folder.Update()
    }
}

Function IsFolderExcluded()
{
    # Return $true if folder is in the excluded list

    param ($folder)

    $folderPath = GetFolderPath($folder)

    # Check PR_FOLDER_TYPE (0x36010003) to see if this is a search folder (FOLDER_SEARCH=2)
    if (!$IncludeSearchFolders)
    {
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
    }

    if ($ExcludedFolderPaths)
    {
        LogVerbose "Checking for exclusions: $($ExcludeFolderList -join ',')"
        ForEach ($excludedFolder in $ExcludedFolderPaths)
        {
            if ($folderPath.ToLower().Equals($excludedFolder.ToLower()))
            {
                Log "[IsFolderExcluded]Excluded folder being skipped: $folderPath"
                return $true
            }
        }
    }
    return $false
}

Function ProcessFolder()
{
	# Apply updates to the given folder

    param (
        $Folder,
        $defaultFolderClass
    )
	if ($Folder -eq $null)
	{
		throw "No folder specified"
	}
	
    if ( [String]::IsNullOrEmpty($defaultFolderClass) ) { $defaultFolderClass = "IPF.Note" }

	# Process any subfolders
    if ( !$(IsFolderExcluded $Folder) -or $DoNotExcludeSubfolders )
    {
	    if ($ProcessSubFolders)
	    {
		    if ($Folder.ChildFolderCount -gt 0)
		    {
                Log "Processing subfolders of: $($Folder.DisplayName)" Gray
                # We read the list of all folders first, so that we have the complete list before any processing
                $subfolders = @()
                $moreFolders = $True
			    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
                $FolderView.PropertySet = $script:requiredFolderProperties
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
                        if ( !$(IsFolderExcluded $subFolder) )
                        {
				            ProcessFolder $subFolder $defaultFolderClass
                        }
			        }
                }
		    }
            else
            {
                LogVerbose "Folder has no subfolders to process: $($Folder.DisplayName)"
            }
	    }
    }
    if ( $(IsFolderExcluded $Folder) )
    {
        Log "Folder excluded: $($Folder.DisplayName)" Gray
        return
    }

    Log "Processing folder: $($Folder.DisplayName)" Gray

    # Check for folder purge
    if ($Purge)
    {
        PurgeFolder $Folder
    }

    # If delete parameter is set, check if the source folder is empty (and if so, delete it)
    if ($Delete)
    {
	    if (($Folder.TotalCount -eq 0) -And ($Folder.ChildFolderCount -eq 0))
	    {
		    # Folder is empty, so can be safely deleted
		    try
		    {
			    $Folder.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete)
			    LogFolderAction $Folder "deleted" Green
		    }
		    catch {}
            ReportError "deleting $($Folder.DisplayName)"
	    }
	    else
	    {
		    # Folder is not empty
		    LogFolderAction $Folder "could not be deleted as it is not empty." Red
	    }
        return # If Delete is specified, any other parameter is irrelevant
    }

    DeleteFolderProperties($Folder)
    AddFolderProperties($Folder)

    if (![String]::IsNullOrEmpty($NewDisplayName))
    {
        # Update display name of the folder
        try
        {
            if (!$WhatIf)
            {
                $Folder.DisplayName = $NewDisplayName
                $folder.Update()
            }
            LogFolderAction $folder "Display name updated to $NewDisplayName" Green
        }
        catch {}
        ReportError "Updating display name"
    }

    RemoveCharactersFromDisplayName($Folder)

    if ( $RepairFolderClass )
    {
        if ( [String]::IsNullOrEmpty($Folder.FolderClass) )
        {
            # Empty folder class, so set it
            try
            {
                if (!$WhatIf)
                {
                    $Folder.FolderClass = $defaultFolderClass
                    $folder.Update()
                }
                LogFolderAction $Folder "Folder class updated to $defaultFolderClass" Green
            }
            catch {}
            ReportError "Updating folder class"
        }
        else
        {
            LogFolderAction $Folder "Folder class is $($Folder.FolderClass)" Green
        }
    }
    elseif (![String]::IsNullOrEmpty($FolderClass))
    {
        # Update default item class of the folder
        try
        {
            if (!$WhatIf)
            {
                $Folder.FolderClass = $FolderClass
                $folder.Update()
            }
            LogFolderAction $Folder "Folder class updated to $FolderClass" Green
        }
        catch {}
        ReportError "Updating folder class"
    }
}

function ValidateWellKnownFolderDefaultItemType()
{
    param (
        $FolderName,
        $FolderClass
    )

    $folder = $null
    $folder = GetFolder $null "WellKnownFolderName.$FolderName" $False
    if ( !($folder -eq $null) )
    {
        if ($folder.FolderClass -ne $FolderClass)
        {
            if ( [String]::IsNullOrEmpty($folder.FolderClass) )
            {
                # FolderClass is currently blank, so we should be able to set it
                try
                {
                    $folder.FolderClass = $FolderClass
                    $folder.Update()
                    LogFolderAction $folder "Missing folder class updated to $FolderClass" Green
                }
                catch {}
                ReportError "Updating folder class [$($folder.DisplayName)]"
            }
            else
            {
                # FolderClass is set, so we can't change it but will report that it is not the expected type
                LogFolderAction $folder "Unexpected FolderClass ($($folder.FolderClass); should be $FolderClass" Red
            }
        }
        else
        {
            LogFolderAction $folder "FolderClass is already $($folder.FolderClass)" Green
        }
    }
    else
    {
        Log "Failed to check folder type for WellKnownFolderName::$FolderName" Red
    }
}

function ValidateWellKnownFolderDefaultItemTypes()
{
    # Go through each of the well known folders and check the item class

    ValidateWellKnownFolderDefaultItemType "Calendar" "IPF.Appointment"
    ValidateWellKnownFolderDefaultItemType "Contacts" "IPF.Contact"
    ValidateWellKnownFolderDefaultItemType "Journal" "IPF.Journal"
    ValidateWellKnownFolderDefaultItemType "Notes" "IPF.StickyNote"
    ValidateWellKnownFolderDefaultItemType "Tasks" "IPF.Task"

    # All other folders have a default item class of IPF.Note, so we don't need to explicitly check them
}

function ProcessMailbox()
{
    # Process the mailbox
    Log([string]::Format("Accessing mailbox {0}", $Mailbox))  Gray
    $script:currentThrottlingDelay = 0
    $script:currentBatchSize = 500
	$script:service = CreateService($Mailbox)
	if ($script:service -eq $Null)
	{
		Log "Failed to create ExchangeService, cannot continue" Red
        return
	}

    # Define the properties we need to retrieve
    $script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount, $script:PR_FOLDER_TYPE)
	
	# Get our root folder
    $rootFolder = $Null
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
    if ($PublicFolders)
    {
        Write-Host "Processing public folders" Gray
        $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
    }
    else
    {
        if ($Archive)
        {
            LogVerbose "Processing archive mailbox"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx )
            $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
        else
        {
            LogVerbose "Processing primary mailbox"
            $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
            $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)
        }
    }

    # Check we could get the root folder (if not, there's nothing else we can do)
    if ($rootFolder -eq $Null)
    {
        Log "Unable to bind to root folder.  No further processing possible" Red
        return
    }

    if ($RepairFolderClass)
    {
        # When validating the folder class of folders, we need to ensure that the well known folders are set correctly, otherwise the repair logic will set everything to IPF.Note
        ValidateWellKnownFolderDefaultItemTypes
    }


    # FolderPath could be an array...
    if ([String]::IsNullOrEmpty($FolderPaths))
    {
        ProcessFolder $rootFolder "IPF.Note"
    }
    else
    {
        foreach ($fPath in $FolderPaths)
        {
            # Now get our specific folder, if we have one
            $folder = $Null
            if (![String]::IsNullOrEmpty(($fPath)))
            {
                $folder = GetFolder $rootFolder $fPath $False
            }

            # Now start processing the folder
            if ($folder -ne $Null)
            {
                ProcessFolder $folder
            }
            else
            {
                Log "Unable to locate folder: $fPath" Red
            }
        }
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

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($OAuth)
    {
        Write-Host "Please specify *either* -Credentials *or* -OAuth" Red
        Exit
    }
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