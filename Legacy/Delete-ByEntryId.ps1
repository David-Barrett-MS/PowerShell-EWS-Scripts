#
# Delete-ByEntryId.ps1
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
	[Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed.  If not present, then the first EntryId supplied must be the PrimarySmtpAddress of the mailbox.")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,

	[Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox will be accessed (instead of the main mailbox)")]
	[switch]$Archive,
		
	[Parameter(Mandatory=$False,HelpMessage="List of EntryIds (optionally with named properties) to delete (or path to file containing this list)")]
    [ValidateNotNullOrEmpty()]
	$EntryIds,

	[Parameter(Mandatory=$False,HelpMessage="If this switch is set, the whole item will be deleted (otherwise, just the named property will be)")]
	[switch]$DeleteItems,

	[Parameter(Mandatory=$False,HelpMessage="If specified, items are deleted in batches.  This is a lot quicker, but doesn't retrieve the item details prior to deletion.")]	
    [switch]$Batch,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [System.Management.Automation.PSCredential]$Credentials,
				
	[Parameter(Mandatory=$False,HelpMessage="If set, ApplicationImpersonation will not be used to access the mailbox (FullAccess rights will be required for the authenticating account)")]	
	[switch]$DoNotImpersonate,

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
$script:ScriptVersion = "1.0.3"

# Define our functions

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
	if ( $LogFile -eq "" ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
	if ( $LogFile -eq "" ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
	if ( $LogFile -eq "" ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

$script:LastError = $Error[0]
Function ErrorReported($Context)
{
    # We check for errors here, using $Error variable, as try...catch isn't reliable when remoting
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

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	# Check if we've been given the path to the managed API
    if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( Test-Path $EWSManagedApiPath )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) Yellow
	}
	
    # Search for the managed API
    $a = Get-ChildItem -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
    if (!$a)
    {
	    $a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction Ignore | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
    }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction Ignore | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
    # If we've found it, we can load the managed API now
	if ($a)	
	{
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
        $script:EWSManagedApiPath = $a.VersionInfo.FileName
		return $true
	}
	return $false
}

Function TrustAllCerts()
{
    # Implement call-back to override certificate handling (and accept all)
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    ErrorReported "CreateCompiler" | Out-Null
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
    ErrorReported "CompileAssembly" | Out-Null
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    ErrorReported "CreateInstance" | Out-Null
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
    ErrorReported "Assign Policy" | Out-Null
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

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

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

    LogVerbose "Creating ExchangeService for: $smtpAddress"

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
        LogVerbose "Using EWS Url: $EwsUrl"
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
 
    if (!$DoNotImpersonate)
    {
        # We are using ApplicationImpersonation to access the mailbox
	    $exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpAddress)
    }
    $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    CreateTraceListener $exchangeService
    $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $exchangeService.TraceEnabled = $True

    # To test we have access to the mailbox, we bind to the Inbox.  Any error here and we fail.
    $testBindFolder = $null
    try
    {
        if ($Archive)
        {
            $testBindFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        }
        else
        {
            $testBindFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        }
    } catch {}
    ReportError "Bind to folder in mailbox"
    if ($testBindFolder -eq $null) { return $null }

    return $exchangeService
}

function ConvertId($entryId)
{
    # Use EWS ConvertId function to convert from EntryId to EWS Id

    $id = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
    $id.Mailbox = $Mailbox
    $id.UniqueId = $entryId
    $id.IsArchive = $Archive
    $id.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId
    $ewsId = $Null
    try
    {
        $ewsId = $script:service.ConvertId($id, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
    }
    catch {}
    ErrorReported | out-null
    LogVerbose "EWS Id: $($ewsId.UniqueId)"
    return $ewsId
}

Function RemoveProcessedItemsFromList()
{
    # Process the results of a batch move/copy and remove any items that were successfully moved from our list of items to move
    param (
        $requestedItems,
        $results,
        $Items
    )

    $remainingItems = @()
    if ($results -ne $null)
    {
        $failed = 0
        for ($i = 0; $i -lt $requestedItems.Count; $i++)
        {
            if ($results[$i].ErrorCode -eq "NoError")
            {
                LogVerbose "Item successfully processed: $($requestedItems[$i])"
            }
            else
            {
                if ( ($results[$i].ErrorCode -eq "ErrorMoveCopyFailed") -or ($results[$i].ErrorCode -eq "ErrorInvalidOperation") -or ($results[$i].ErrorCode -eq "ErrorItemNotFound") )
                {
                    # This is a permanent error, so we remove the item from the list
                    
                }
                else
                {
                    $remainingItems += $requestedItems[$i]
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
    return $remainingItems
}

Function BatchDelete()
{
    # Send request to move/copy items, allowing for throttling (which in this case is likely to manifest as time-out errors)
    param (
        $ItemsToDelete,
        $BatchSize = 200
    )

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
			$results = $script:Service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
        }
        catch
        {
            try
            {
                Log "Unexpected error: $($Error[0].Exception.InnerException.ToString())" Red
            }
            catch
            {
                Log "Unexpected error: $($Error[1])" Red
            }
        }

        $ItemsToDelete = RemoveProcessedItemsFromList $deleteIds $results $ItemsToDelete

        $percentComplete = ( ($totalItems - $ItemsToDelete.Count) / $totalItems ) * 100
        Write-Progress -Activity $progressActivity -Status "$percentComplete% complete" -PercentComplete $percentComplete

        if ($ItemsToDelete.Count -eq 0)
        {
            $finished = $True
        }
    }
    Write-Progress -Activity $progressActivity -Status "Complete" -Completed
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
    return $null
}

function GetExtendedPropertyDefinition($guid, $name, $mapiType)
{
    # Return an EWS ExtendedPropertyDefinition for the given MAPI property

    return new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($( new-object System.Guid($guid) ), $name, $( EWSPropertyType $mapiType ))
}

function ProcessMailbox()
{
    # Process the mailbox
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray

    # Bind to the mailbox
	$script:Service = CreateService($Mailbox)
    if ( $script:Service -eq $null )
    {
        Write-Host "Failed to create ExchangeService" -ForegroundColor Red
        exit
    }
    $entryIdsBatch = @()
    $activity = "Deleting named properties from items"
    if ($DeleteItems)
    {
        if ($Batch)
        {
            $activity = "Collating items to delete"
        }
        else
        {
            $activity = "Deleting items"
        }
    }
    $itemsProcessed = 0

    # Now we process our list of EntryIds and delete the items
    ForEach ($entryIdElement in $EntryIds)
    {
        Write-Progress -Activity $activity -Status "Processing item" -PercentComplete (($itemsProcessed++/$($EntryIds.Count))*100)
        if (!$entryIdElement.Contains("@")) # A real EntryId does not contain @
        {
            $entryId = $null
            $namedProps = $null
            if ($entryIdElement.Contains(";"))
            {
                # EntryId has named property information
                $entryId, $namedProps = $entryIdElement -split ";"
            }
            else
            {
                # Seems to be just an EntryId
                $entryId = $entryIdElement
            }

            if ($entryId)
            {
                Write-Progress -Activity $activity -Status "Processing item $entryId" -PercentComplete (($itemsProcessed/$($EntryIds.Count))*100)
                LogVerbose "Converting EntryId to EwsId: $entryId"
                $ewsId = ConvertId($entryId)
                $basePropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)

                if ($ewsId)
                {
                    if ($DeleteItems)
                    {
                        # We are deleting the items (not just the named properties)
                        if ($Batch)
                        {
                            # We're batching the items to delete
                            $entryIdsBatch += $ewsId.UniqueId
                            Log "Adding item to delete list: $($ewsId.UniqueId)"
                        }
                        else
                        {
                            LogVerbose "Binding to item: $($ewsId.UniqueId)"
                            $item = $null
                            $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($script:Service, $ewsId.UniqueId, $basePropset)
                            if ($item)
                            {
                                Log "Deleting item: $($item.Subject) ($($ewsId.UniqueId))"
                                try
                                {
                                    if ($item.ItemClass.StartsWith("IPM.Appointment"))
                                    {
                                        # This is an appointment, so we hard delete and suppress notifications to attendees
                                        [Microsoft.Exchange.WebServices.Data.Appointment]$apt = $item
                                        $apt.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone)
                                    }
                                    else
                                    {
                                        $item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $true)
                                    }
                                }
                                catch {}
                                ReportError "Deleting item"
                            }
                        }
                    }
                    else
                    {
                        # We are just deleting the named properties (we do this one item at a time, not in batches [yet])
                        $item = $null
                        $extendedProperties = @()
                        $propset = $basePropset
                        foreach ($namedProp in $namedProps)
                        {
                            try
                            {
                                LogVerbose "Parsing property: $namedProp"
                                $namedPropElements = $namedProp -split "/"
                                if ($namedPropElements.Count -eq 3)
                                {
                                    # Named prop should be in format guid/name/type
                                    $ewsPropDef = $( GetExtendedPropertyDefinition $namedPropElements[0] $namedPropElements[1] $namedPropElements[2] )
                                    $extendedProperties += $ewsPropDef
                                    $propset.Add($ewsPropDef)
                                }
                            }
                            catch {}
                            ReportError "Parsing property"
                        }

                        if ( $extendedProperties.Count -gt 0 )
                        {
                            LogVerbose "Binding to item: $($ewsId.UniqueId)"
                            $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($script:Service, $ewsId.UniqueId, $propset)

                            if ($item)
                            {
                                Log "Deleting properties from item: $($item.Subject) ($($ewsId.UniqueId))"
                                try
                                {
                                    foreach ($extendedProperty in $extendedProperties)
                                    {
                                        [void]$item.RemoveExtendedProperty($extendedProperty)
                                    }
                                    [void]$item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, $true)
                                }
                                catch {}
                                ReportError "Delete properties"
                            }
                        }
                    }
                }
                else
                {
                    Log "Failed to convert EntryId to EWS Id: $entryId" Red
                }
            }
            else
            {
                LogVerbose "Invalid EntryId ignored"
            }
        }
    }
    Write-Progress -Activity $activity -Completed

    if ($Batch)
    {
        if ($Force)
        {
            Log "Delete list contains $($entryIdsBatch.Count) items.  Starting batch delete." Yellow
            BatchDelete $entryIdsBatch
        }
        else
        {
            Log "Delete list contains $($entryIdsBatch.Count) items.  Batch delete not processed as -Force not specified" Green
        }
    }
}

# The following is the main script

# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Log "WARNING: Ignoring any SSL certificate errors" Yellow
    TrustAllCerts
    ErrorReported | out-null
}
 
# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Log "Failed to locate EWS Managed API, cannot continue" Red
	Exit
}
  
Write-Host ""

# Check whether we have a file as input...
$FileExists = Test-Path $EntryIds
If ( $FileExists )
{
	# Import the EntryIds from a text file
    LogVerbose "Reading EntryIDs from file: $EntryIds"
	$EntryIds = Get-Content -Path $EntryIds
}

# Check we have a valid mailbox (either specified, or first entry of the EntryIds list)
if ($EntryIds[0])
{
    if ($EntryIds[0].Contains("@"))
    {
        if ( [String]::IsNullOrEmpty($Mailbox) -or ($Mailbox.ToLower().Equals($EntryIds[0].ToLower())) )
        {
            # Mailbox is first EntryId
            LogVerbose "Mailbox specified in file: $($EntryIds[0])"
            $Mailbox = $EntryIds[0]
            if ($EntryIds.Count -lt 2)
            {
                Log "EntryIds file found, but no EntryIds were listed" Yellow
                Exit
            }
        }
        else
        {
            # The mailbox specified in -Mailbox parameter does not match that in the EntryIds list, so we fail here
            Log "Mailbox specified by parameter: $Mailbox" White
            Log "Mailbox specified in EntryIds list: $($EntryIds[0])" White
            Log "Mailbox mismatch between parameter and item list" Red
            Exit
        }
    }
}
else
{
    if ($Error[0])
    {
        Log "Error: $($Error[0])"
    }
    Log "No EntryIds to process"
}

if ( [string]::IsNullOrEmpty($Mailbox) )
{
	Log "Mailbox is required (but not specified)." -ForegroundColor Red
	Exit
}

# Process as single mailbox
ProcessMailbox


if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}