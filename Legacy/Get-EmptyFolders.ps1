#
# Get-EmptyFolders.ps1
#
# By David Barrett, Microsoft Ltd. 2015. Use at your own risk.  No warranties are given.
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

	[Parameter(Mandatory=$False,HelpMessage="Folder to search from - if omitted, the mailbox message root folder is assumed.  All folders beneath this folder will be scanned.")]
	[string]$FolderPath,

	[Parameter(Mandatory=$False,HelpMessage="List of folders to ignore")]
	$IgnoreList = @("\Conversation History", "\Contacts\Skype for Business Contacts"),

	[Parameter(Position=1,Mandatory=$False,HelpMessage="If specified, export CSV of empty folders to this file ({0} will be replaced by SMTP address of the mailbox being reported on).  e.g. c:\Temp\EmptyFolderReport-{0}.csv")]
	[ValidateNotNullOrEmpty()]
	[string]$ReportToFile,

	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, empty folders will be deleted (user will be prompted to confirm unless -Force is also specified)")]
	[switch]$Delete,

	[Parameter(Mandatory=$False,HelpMessage="If this switch is present, empty folders will be deleted without confirmation")]
	[switch]$Force,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [System.Management.Automation.PSCredential]$Credentials,
				
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$Username,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$Password,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$Domain,
	
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

# Define our functions

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
	Write-Host $Details -ForegroundColor $Colour
	if ( [String]::IsNullOrEmpty($LogFile) ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details

    if ($VerbosePreference -eq "SilentlyContinue") { return } # We only log verbose messages to the log-file if verbose messages are shown in the console
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( Test-Path $EWSManagedApiPath )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) Yellow
	}
	
	$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
	if ($a)	
	{
		# Load EWS Managed API
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
        $script:EWSManagedApiPath = $a.VersionInfo.FileName
		return $true
	}
	return $false
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
    return $false
}

function ThrottledFolderBind()
{
    param (
        $folderId,
        $propset = $null)

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
        Sleep -Milliseconds $script:throttlingDelay
        if (-not ($folder -eq $null))
        {
            LogVerbose "Successfully bound to $($folderId): $($folder.DisplayName)" White
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

function CreateService($targetMailbox)
{
    # Creates and returns an ExchangeService object to be used to access mailboxes

    # First of all check to see if we have a service object for this mailbox already
    if ($script:services -eq $null)
    {
        $script:services = @{}
    }
    if ($script:services.ContainsKey($targetMailbox))
    {
        return $script:services[$targetMailbox]
    }

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

    # Set credentials if specified, or use logged on user.
    if ($Credentials -ne $Null)
    {
        LogVerbose "Applying given credentials"
        $exchangeService.Credentials = $Credentials.GetNetworkCredential()
    }
    elseif ($Username -and $Password)
    {
	    LogVerbose "Applying given credentials for $Username"
	    if ($Domain)
	    {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
	    } else {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
	    }
    }
    else
    {
	    LogVerbose "Using default credentials"
        $exchangeService.UseDefaultCredentials = $true
    }

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
    	$exchangeService.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
		    LogVerbose "Performing autodiscover for $targetMailbox"
		    if ( $AllowInsecureRedirection )
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox, {$True})
		    }
		    else
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox)
		    }
		    if ([string]::IsNullOrEmpty($exchangeService.Url))
		    {
			    Log "$targetMailbox : autodiscover failed" Red
			    return $null
		    }
		    LogVerbose "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
            Log "$targetMailbox : error occurred during autodiscover: $($Error[0])" Red
            return $null
    	}
    }
 
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox)
	}

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    CreateTraceListener $exchangeService
    $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $exchangeService.TraceEnabled = $True

    $script:services.Add($targetMailbox, $exchangeService)
    return $exchangeService
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

    if ($Folder -eq "\")
    {
        # Special handling for root folder
        if  ($script:folderCache.ContainsKey("\"))
        {
            return $script:folderCache["\"]
        }
        $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
        $rootFolder = ThrottledFolderBind $folderId $propset
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
        $parentFolder = ThrottledFolderBind $Folder.Id $propset
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
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propset
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = "$($parentFolder.DisplayName)\$folderPath"
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function GetWellKnownFolderIds()
{
    # Get the Ids of all the well known folders (so that we can exclude them from our search)
    $wellKnownFolders = [System.Enum]::GetNames([Microsoft.Exchange.WebServices.Data.WellKnownFolderName])
    $script:wellKnownFolderIds = @()

    foreach ($wellKnownFolder in $wellKnownFolders)
    {
        try
        {
            $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            $folder = ThrottledFolderBind $wellKnownFolder $propset
            $script:wellKnownFolderIds += $folder.Id
        }
        catch
        {
            # We ignore any errors, as it just means that the particular folder we've queried doesn't exist
        }
    }
}

function IsHidden()
{
    param (
        $folder
    )

    # Returns true if the folder is hidden
    if ($folder.ExtendedProperties.Count -lt 1)
    {
        return $false
    }
    foreach ($prop in $folder.ExtendedProperties)
    {
        if ($prop.PropertyDefinition -eq $script:PidTagAttributeHidden)
        {
            #Write-Host "Hidden" -ForegroundColor Red
            return $prop.Value
        }
    }
    return $false
}

function SearchEmptyFolders()
{
    param (
        $folder
    )

    $folder.Load($script:FolderPropSet)
    if (IsHidden $folder)
    {
        LogVerbose "Ignoring hidden folder: $($Folder.DisplayName)"
        return
    }
    LogVerbose "Processing: $($Folder.DisplayName)"

	# Recurse into any subfolders first
	$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $FolderView.PropertySet = $script:FolderPropSet
	$FindFolderResults = $folder.FindFolders($FolderView)
    Sleep -Milliseconds $script:throttlingDelay
	ForEach ($subFolder in $FindFolderResults.Folders)
	{
        SearchEmptyFolders $subFolder
	}
    
    # Now we load the properties for this folder to see if it is empty (i.e. no subfolders, no items)
    $folder.Load($script:FolderPropSet)
    if ( ($folder.TotalCount -eq 0) -and ($folder.ChildFolderCount -eq 0) )
    {
        # This folder is empty
        $folderPath = GetFolderPath $folder
        $isSpecialFolder = $false
        if ($script:wellKnownFolderIds.Contains($folder.Id))
        {
            $isSpecialFolder = $true
            Log "$folderPath is empty, but is well known folder" Gray
        }
        if ($IgnoreList)
        {
            if ($IgnoreList.Contains("$folderPath".ToLower()))
            {
                $isSpecialFolder = $true
                Log "$folderPath is empty, but is on ignore list" Gray
            }
        }

        if (-not $isSpecialFolder)
        {
            Log "$folderPath is empty" Green
            $script:emptyFolders += $folderPath
            if ($Delete)
            {
                # Attempt to delete the folder
                try
                {
                    $deleteThisFolder = $true
                    if (-not $Force)
                    {
                        # Need to ask the user whether we should delete this folder
                        $response = Read-Host -Prompt "Confirm delete (YyNn)? $folderPath"
                        if (-not $response.ToLower().Equals("y"))
                        {
                            $deleteThisFolder = $false
                        }
                    }
                    if ($deleteThisFolder)
                    {
                        $folder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                        Log "$folderPath has been deleted" Yellow
                    }
                    else
                    {
                        Log "$folderPath was not deleted" Gray
                    }
                }
                catch
                {
                    if ($Error[0].Exception.Message.Contains("Distinguished folders cannot be deleted."))
                    {
                        # We shouldn't encounter this error as we exclude WellKnownFolders
                        Log "$folderPath was NOT deleted as it is a distinguished folder" Gray
                    }
                    else
                    {
                        Log "$folderPath was NOT deleted due to error: $($Error[0].Exception.Message)" Red
                    }
                }
            }
        }
    }
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
        $FolderPath = "wellknownfoldername.MsgFolderRoot"
    }
    if ($FolderPath.ToLower().StartsWith("wellknownfoldername."))
    {
        # Well known folder specified (could be different name depending on language, so we bind to it using WellKnownFolderName enumeration)
        $wkf = $FolderPath.SubString(20)
        LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind($folderId)
    }
    else
    {
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
        $Folder = ThrottledFolderBind($folderId)
        if ($Folder -and ($FolderPath -ne "\"))
        {
	        $Folder = GetFolder($Folder, $FolderPath, $false)
        }
    }

	if (!$Folder)
	{
		Log "Failed to find folder $FolderPath" Red
		return
	}
	
    # Now we search for empty folders below our root folder

    # Declare the property set that we need to retrieve for each folder
    $script:PidTagAttributeHidden = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x10F4, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
    $script:FolderPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount, [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PidTagAttributeHidden)

    GetWellKnownFolderIds # We use these to exclude any default folders (e.g. Inbox)

    # Go through the ignore list and add the root folder to the path
    if ($IgnoreList.Count -gt 0)
    {
        $rootFolderPath = GetFolderPath("\")
        for ($i=0; $i -lt $IgnoreList.Count; $i++)
        {
            if ($IgnoreList[$i].StartsWith("\"))
            {
                $IgnoreList[$i] = "$rootFolderPath$($IgnoreList[$i])"
            }
            else
            {
                $IgnoreList[$i] = "$rootFolderPath\$($IgnoreList[$i])"
            }
            $IgnoreList[$i] = $IgnoreList[$i].ToLower()
        }
    }

    # Now search the folder heirarchy for empty folders, and report
    $script:emptyFolders = @() # Collect the paths of any empty folders to export to CSV
    SearchEmptyFolders $Folder
    if (![String]::IsNullOrEmpty($ReportToFile))
    {
        if ($script:emptyFolders.Count -gt 0)
        {
            $export = @()
            foreach ($emptyFolder in $script:emptyFolders)
            {
                $info = @{ "Mailbox" = $Mailbox; "Empty Folder" = $emptyFolder }
                $export += New-Object PSObject -Property $info
            }
            $exportFile = [String]::Format($ReportToFile, $Mailbox)
            $export | Select-Object Mailbox,"Empty Folder" | Sort-Object Mailbox,"Empty Folder" | Export-CSV $exportFile -NoTypeInformation
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
    Write-Host "The API can be downloaded from the Microsoft Download Centre: http://www.microsoft.com/en-us/search/Results.aspx?q=exchange%20web%20services%20managed%20api&form=DLC"
    Write-Host "Use the latest version available"
	Exit
}

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($Username -or $Password)
    {
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red
        Exit
    }
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

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}