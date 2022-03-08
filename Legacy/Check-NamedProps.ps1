#
# Check-NamedProps.ps1
#
# By David Barrett, Microsoft Ltd. 2019. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

$version = "1.1.6"

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

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
	if ( $LogFile -eq "" ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
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

Function CmdletsAvailable()
{
    param (
        $RequiredCmdlets,
        $Silent = $False,
        $PSSession = $null
    )

    $cmdletsAvailable = $True
    foreach ($cmdlet in $RequiredCmdlets)
    {
        $cmdletExists = $false
        if ($PSSession)
        {
            $cmdletExists = $(Invoke-Command -Session $PSSession -ScriptBlock { Get-Command $Using:cmdlet -ErrorAction Ignore })
        }
        else
        {
            $cmdletExists = $(Get-Command $cmdlet -ErrorAction Ignore)
        }
        if (!$cmdletExists)
        {
            if (!$Silent) { Log "Required cmdlet $cmdlet is not available" Red }
            $cmdletsAvailable = $False
        }
    }

    return $cmdletsAvailable
}

Function CheckEnvironment($PowerShellUrl)
{
    # Now check that we have the required Exchange cmdlets available
    if ( !$(CmdletsAvailable @("Get-Mailbox") $True ) )
    {
        if ( ![String]::IsNullOrEmpty($PowerShellUrl) )
        {
            # Try to connect and import a session
            Log "Connecting to Exchange using PowerShell Url: $PowerShellUrl"
            $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PowerShellUrl -Credential $Credentials #-Authentication Kerberos  -WarningAction 'SilentlyContinue'
            ReportError "New-PSSession"
            $script:ExchangeSession | Export-Clixml -Path "$($FilePath)exchange.session.xml"

            # If we don't have all the cmdlets available, we can't go any further
            if ( !$(CmdletsAvailable @("Get-Mailbox") $False $script:ExchangeSession) )
            {
                Log "Required Exchange cmdlet(s) are missing, cannot continue" Red
                exit
            }

            # Import the session
            Import-PSSession $script:ExchangeSession -AllowClobber -WarningAction 'SilentlyContinue' -CommandType All -DisableNameChecking
            ReportError "Import-PSSession"
        }
    }
}

Function GetMailboxSmtpAddress()
{
    param ($id, $archive, $organisation)

    if ( ![String]::IsNullOrEmpty($script:MailboxId) )
    {
        return $script:MailboxId
    }
    $mb = $null

    if ( ![String]::IsNullOrEmpty($organisation) )
    {
        LogVerbose "Get-Mailbox $id -Organization $organisation"
        $mb = Get-Mailbox $id -Organization $organisation
    }
    else
    {
        LogVerbose "Get-Mailbox $id"
        $mb = Get-Mailbox $id
    }
    if ($mb -eq $null)
    {
        $script:MailboxId = "Unknown: $id"
    }
    else
    {
        if ( ![String]::IsNullOrEmpty($organisation) )
        {
            $script:MailboxId = $mb.PrimarySmtpAddress.Address
        }
        else
        {
            $script:MailboxId = $mb.PrimarySmtpAddress
        }
        if ($archive)
        {
            $script:MailboxId = "[ARCHIVE]$($script:MailboxId)"
        }
    }
    LogVerbose "Mailbox Id: $($script:MailboxId)"
}

$script:dumpEntryIdsCounter = 0
Function DumpEntryIdList()
{
    param (
        $dumpEntryIdsFileName,
        $propIds,
        $force = $false
    )
    
    # Dump all found EntryIds to file - we only do this once every $SaveRestartInfoInterval times requested, as dumping to text file is very expensive in the context of this script
    # (if a message has lots of named properties on it, we don't really want to dump the whole file for every named prop)
    if (!$force)
    {
        $script:dumpEntryIdsCounter++
        if ($script:dumpEntryIdsCounter -lt $SaveRestartInfoInterval) { return }
        $script:dumpEntryIdsCounter = 0
    }

    LogVerbose "Dumping complete entry Id list for $($script:MailboxId)"

    $mailboxSmtpAddress = $script:MailboxId
    if ( ![String]::IsNullOrEmpty($script:MailboxId.Local) )
    {
        # Our mailbox Id is not a string, so we'll piece together the SMTP address
        $mailboxSmtpAddress = "$($script:MailboxId.Local)@$($script:MailboxId.Domain)"
    }
    $mailboxSmtpAddress | out-file $dumpEntryIdsFileName
    foreach ($entryId in $propIds.Keys)
    {
        "$entryId;$($propIds[$entryId])" | out-file $dumpEntryIdsFileName -Append
    }
}

Function Check-NamedProps
{
    param (
	    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox server (only databases from this server will be queried)")]
	    [ValidateNotNullOrEmpty()]
	    $MailboxServer,

	    [Parameter(Position=1,Mandatory=$False,HelpMessage="Specifies the mailbox database to analyse")]
	    [ValidateNotNullOrEmpty()]
	    $MailboxDatabase,

	    [Parameter(Position=2,Mandatory=$False,HelpMessage="Specifies the mailbox to analyse")]
	    [ValidateNotNullOrEmpty()]
	    $Mailbox,

	    [Parameter(Mandatory=$False,HelpMessage="Specifies the organization to which the mailbox belong (Microsoft internal use only)")]
	    [ValidateNotNullOrEmpty()]
	    $Organization,

	    [Parameter(Mandatory=$False,HelpMessage="Processes the archive mailbox instead of the primary")]
	    [ValidateNotNullOrEmpty()]
	    [switch]$Archive,

	    [Parameter(Mandatory=$False,HelpMessage="If set, the script keeps track of where it is so that if interrupted it continues where it left off")]
        [alias("Restartable")]
	    [switch]$Restart,

	    [Parameter(Mandatory=$False,HelpMessage="If set, the script keeps track of where it is so that if interrupted it continues where it left off")]
	    [ValidateNotNullOrEmpty()]
	    [int]$SaveRestartInfoInterval = 50,

	    [Parameter(Mandatory=$False,HelpMessage="If a mailbox has more than the specified number of properties, then the properties are dumped to a file.  Setting this to -1 disables property dumps.")]
	    [ValidateNotNullOrEmpty()]
	    $DumpPropsIfTotalExceeds = 0,

	    [Parameter(Mandatory=$False,HelpMessage="If specifed, only properties that have this name will be searched.")]
	    [ValidateNotNullOrEmpty()]
	    $SearchNamedProp,

	    [Parameter(Mandatory=$False,HelpMessage="If specifed, only properties that have this GUID will be searched.")]
	    [ValidateNotNullOrEmpty()]
	    $SearchGuid,

	    [Parameter(Mandatory=$False,HelpMessage="Type of named property being searched (defaults to PT_STRING).  Only required for a message search of found props.")]
	    [ValidateNotNullOrEmpty()]
	    $NamedPropType = "001f",

	    [Parameter(Mandatory=$False,HelpMessage="Where to dump the properties when needed.  If missing, current directory will be used.")]
	    [ValidateNotNullOrEmpty()]
	    [String]$DumpPropsPath = "",

	    [Parameter(Mandatory=$False,HelpMessage="If specified, the entry Ids of messages with matched named properties will be retrieved and saved (requires -SearchNamedProp)")]
	    [ValidateNotNullOrEmpty()]
	    [switch]$DumpEntryIds,

        [Parameter(Mandatory=$false,HelpMessage="PowerShell Url")]	
        [string]$PowerShellUrl,

	    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	    [string]$LogFile = ""	
    )

    Begin
    {
        # Check we have Exchange cmdlets available
        CheckEnvironment $PowerShellUrl

        # Need to load ManagedStoreDiagnostics.ps1 for Get-StoreQuery
        $originalPath = $(Get-Location).Path
        if ($(Test-Path -Path "$($env:exchangeinstallpath)Scripts\ManagedStoreDiagnosticFunctions.ps1"))
        {
            cd "$($env:exchangeinstallpath)Scripts"
        }
        else
        {
            # ManagedStoreDiagnosticFunctions not found.
            if (!$(Test-Path -Path "ManagedStoreDiagnosticFunctions.ps1"))
            {
                # We can't continue without this
                Log "Could not locate ManagedStoreDiagnosticFunctions.ps1 - this is available from your Exchange server Scripts folder, and should be placed in the same location as this script." Red
                exit
            }
        }
        . .\ManagedStoreDiagnosticFunctions.ps1
        cd $originalPath
    }

    Process
    {
        Log "$($MyInvocation.MyCommand.Name) version $version"

        # Check parameters
        if (![String]::IsNullOrEmpty($DumpPropsPath))
        {
            if (!$DumpPropsPath.EndsWith("\"))
                { $DumpPropsPath = "$DumpPropsPath\" }
        }

        # As this process can take a very long time, we show a progress bar
        Write-Progress -Activity "Reading mailboxes" -Status "Reading database(s)" -PercentComplete 0

        # Get our list of databases
        $script:MailboxGuid = ""
        if ($MailboxServer)
        {
            $databases = Get-MailboxDatabase -Server $MailboxServer
        }
        elseif ($MailboxDatabase)
        {
            $databases = Get-MailboxDatabase -Identity $MailboxDatabase
        }
        elseif ($Mailbox)
        {
            # We are checking a single mailbox, so we locate the database first
            $mb = $null
            if ( ![String]::IsNullOrEmpty($Organization) )
            {
                LogVerbose "Get-Mailbox $Mailbox -Organization $Organization"
                $mb = Get-Mailbox $Mailbox -Organization $Organization
            }
            else
            {
                LogVerbose "Get-Mailbox $Mailbox"
                $mb = Get-Mailbox $Mailbox
            }
            if ( $mb -eq $null )
            {
                Log "Failed to retrieve mailbox $Mailbox" Red
                Exit
            }
            if ( ![String]::IsNullOrEmpty($Organization) )
            {
                $script:MailboxId = $mb.PrimarySmtpAddress.Address
            }
            else
            {
                $script:MailboxId = $mb.PrimarySmtpAddress
            }
            if ($Archive)
            {
                Log "Checking archive mailbox $Mailbox, which is located in database $($mb.ArchiveDatabase)"
                $databases = Get-MailboxDatabase -Identity $mb.ArchiveDatabase
                $script:mailboxGuid = $mb.ArchiveGuid
            }
            else
            {
                Log "Checking mailbox $Mailbox, which is located in database $($mb.Database)"
                $databases = Get-MailboxDatabase -Identity $mb.Database
                $script:mailboxGuid = $mb.ExchangeGuid
            }
        }
        else
        {
            $databases = Get-MailboxDatabase
        }
        $currentDb = 0
        $dbCount = $databases.Count
        if (!$dbCount) { $dbCount = 1 }

        $wildcardSearch = $false
        if ( ![String]::IsNullOrEmpty($SearchNamedProp) )
        {
            $wildcardSearch = $SearchNamedProp.Contains("*")
        }
        LogVerbose "Wildcard search: $wildcardSearch"

        foreach ($database in $databases)
        {
            Log "Processing database $($database.Name) ($($database.AdminDisplayVersion))"
            Write-Progress -Activity "Reading mailboxes" -Status "Reading mailboxes in database $database.Name ($currentDb of $dbCount)" -PercentComplete (($currentDb++/$dbCount)*100)
            # Get list of mailboxes and their numbers
            $mailboxQuery = "SELECT MailboxGuid,MailboxInstanceGuid,DisplayName,MailboxNumber FROM Mailbox"
            if ( ![String]::IsNullOrEmpty($MailboxGuid) )
            {
                $mailboxQuery = "$mailboxQuery WHERE MailboxGuid=`"$($MailboxGuid)`""
            }
            LogVerbose "Mailbox query: $mailboxQuery"
            $mbxs = Get-StoreQuery -Database $database.Name -Query $mailboxQuery -Unlimited
            $mbxCount = $mbxs.Count
            if ( ($mbxCount -eq $null) -and $mbxs)
            {
                # We don't have Count property returned when we only have one mailbox
                $mbxCount = 1
            }

            # For each mailbox, get the named prop count
            $currentMbx = 0
            if ($mbxCount -lt 1)
            {
                # We didn't find any mailboxes
                Log "No mailboxes to check"
            }
            else
            {
                foreach ($mbx in $mbxs)
                {
                    GetMailboxSmtpAddress $mbx.MailboxGuid.ToString() $Archive $Organization
                    $entryIdsDumpFile = "$DumpPropsPath$($mbx.MailboxGuid).EntryIds.txt"
                    $propIds = @{}
                    $alreadySearchedProps = @()
                    if ( $Restart )
                    {
                        # In restartable mode, so we check for an already existing file, and if it is present, load it
                        if ( $(Test-Path $entryIdsDumpFile) )
                        {
	                        # Import the EntryIds from a text file
                            Log "EntryId export already exists for mailbox, reading existing export from file: $entryIdsDumpFile"
	                        $exportedEntryIds = Get-Content -Path $entryIdsDumpFile
                            if ( $exportedEntryIds )
                            {
                                if ($exportedEntryIds[0] -ne $script:MailboxId)
                                {
                                    Log "Existing mailbox export does not match, restart not possible" Red
                                    $entryIdsDumpFile = "$DumpPropsPath$($mbx.MailboxGuid).EntryIds.New.txt"
                                    Log "Will export to file: $entryIdsDumpFile"
                                }
                                else
                                {
                                    for ( $i=1; $i -lt $exportedEntryIds.Length; $i++ )
                                    {
                                        $importedEntryId, $exportedNamedProps = $exportedEntryIds[$i] -split ";"
                                        foreach ($exportedNamedProp in $exportedNamedProps)
                                        {
                                            $alreadySearchedProps += $exportedNamedProp
                                        }
                                        $propIds.Add( $importedEntryId, $($exportedNamedProps -join ";") )

                                    }
                                }
                            }
                        }
                        else
                        {
                            # EntryIds export file doesn't exist, so create one with the mailbox Id as the first line
                            $mailboxSmtpAddress = $script:MailboxId
                            if ( ![String]::IsNullOrEmpty($script:MailboxId.Local) )
                            {
                                # Our mailbox Id is not a string, so we'll piece together the SMTP address
                                $mailboxSmtpAddress = "$($script:MailboxId.Local)@$($script:MailboxId.Domain)"
                            }
                            $mailboxSmtpAddress | out-file $entryIdsDumpFile
                        }
                    }

                    $currentMbx++
                    Write-Progress -Activity "Retrieving named props" -Status "Retrieving named properties in mailbox $($mbx.DisplayName)" -PercentComplete (($currentMbx/$mbxCount)*100)
                    $skip = $false
                    # We ignore system and health mailboxes
                    if ($mbx.DisplayName.StartsWith("SystemMailbox") -or $mbx.DisplayName.Contains("HealthMailbox") -or $mbx.DisplayName.Equals("Microsoft Exchange"))
                        { $skip = $true }

                    if (!$skip)
                    {
                        LogVerbose "Retrieving properties from mailbox $($mbx.DisplayName)"
                        $propQuery = "SELECT PropNumber,PropGuid,PropName,PropDispId FROM ExtendedPropertyNameMapping WHERE MailboxPartitionNumber=$($mbx.MailboxNumber)"
                        if (![String]::IsNullOrEmpty($SearchNamedProp))
                        {
                            # If the name contains wildcard, we don't limit our query (we filter in the script instead)
                            if ( !$WildCardSearch )
                            {
                                $propQuery = "$propQuery AND PropName=`"$SearchNamedProp`""
                            }
                        }
                        if (![String]::IsNullOrEmpty($SearchGuid))
                        {
                            $propQuery = "$propQuery AND PropGuid=`"$SearchGuid`""
                        }
                        LogVerbose "Prop query: $propQuery"

                        $namedProps = Get-StoreQuery -Database $database.Name -Query $propQuery -Unlimited
                        $namedPropCount = $namedProps.Count
                        if (!$namedPropCount) { $namedPropCount = 0 }
                        LogVerbose "Number of properties returned: $($namedPropCount)"

                        # We create a new PsObject to hold this data as then we can pipe to other cmdlets (e.g. Export-CSV)
                        $output = New-Object PsObject
                        if ( ![String]::IsNullOrEmpty($script:MailboxGuid) )
                        {
                            $output | Add-Member -MemberType NoteProperty -Name "MailboxGuid" -Value $script:MailboxGuid
                        }
                        else
                        {
                            $output | Add-Member -MemberType NoteProperty -Name "MailboxGuid" -Value $mbx.MailboxGuid
                        }
                        $output | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mbx.DisplayName
                        $output | Add-Member -MemberType NoteProperty -Name "NamedPropCount" -Value $namedPropCount
                        $output

                        if ( ($namedPropCount -gt 0) -and ($DumpPropsIfTotalExceeds -ge 0) )
                        {
                            if ( ($DumpPropsIfTotalExceeds -le $namedPropCount) -or $SearchNamedProp -or $SearchGuid )
                            {
                                # This mailbox exceeds our property limit (or contains the specific named property we are looking for), so we dump all the properties to a file for further investigation
                                $dumpPropsFileName = "$DumpPropsPath$($mbx.MailboxGuid).namedprops.xml"
                                LogVerbose "Dumping property list to: $dumpPropsFileName"
                                $namedProps | Export-Clixml $dumpPropsFileName

                                if ($DumpEntryIds)
                                {
                                    # Now we perform a search for all messages with these properties
                                    $mailboxInstanceGuid = [BitConverter]::ToString($mbx.MailboxInstanceGuid.ToByteArray()).Replace("-","") # Required to create EntryId
                                    $dumpEntryIdsDebugFileName = "$DumpPropsPath$($mbx.MailboxGuid).EntryIds.Debug.{n}.txt"
                                    $dumpEntryIdsDebugIndex = 1
                                    $currentProp = 1
                                    $namedPropCount = $namedProps.Count
                                    $smtpWritten = $false                                 

                                    DumpEntryIdList $entryIdsDumpFile $propIds # This ensures we have any restartable info, and stamps the mailbox name at the top of the file
                                    ForEach ($namedProp in $namedProps)
                                    {
                                        $ewsPropId = "$($namedProp.PropGuid)/$($namedProp.PropName)/$NamedPropType"
                                        $nameMatch = !$wildcardSearch
                                        if ( $wildcardSearch )
                                        {
                                            if ( $namedProp.PropName -like $SearchNamedProp )
                                            {
                                                LogVerbose "$($namedProp.PropName) matches $SearchNamedProp"
                                                $nameMatch = $true
                                            }
                                            else
                                            {
                                                LogVerbose "$($namedProp.PropName) does NOT match $SearchNamedProp"
                                            }
                                        }

                                        if ( $Restart )
                                        {
                                            # If we've restarted, check we haven't already searched for this prop
                                            if ( $alreadySearchedProps.Contains($ewsPropId) )
                                            {
                                                $nameMatch = $false
                                                LogVerbose "Already searched for property $ewsPropId, skipping"
                                            }
                                        }

                                        if ( $nameMatch )
                                        {
                                            Write-Progress -Activity "Searching for properties" -Status "Searching for property $($namedProp.PropNumber)" -PercentComplete (($currentProp++/$namedPropCount)*100)
                                            $propId = "p$('{0:x}' -f $namedProp.PropNumber)$NamedPropType"
                                            
                                            $messageQuery = "SELECT MessageId, FolderId FROM Message WHERE MailboxPartitionNumber=$($mbx.MailboxNumber) AND $propId != null"
                                            LogVerbose "Message query: $messageQuery"
                                            $messages = Get-StoreQuery -Database $database.Name -Query $messageQuery -Unlimited

                                            # Get a valid message count.  Store query returns a blank record rather than no records when no data is found
                                            $messageCount = $messages.Count
                                            if (!$messageCount)
                                            {
                                                $messageCount = 0
                                                if (![String]::IsNullOrEmpty($messages.FolderId) -and ![String]::IsNullOrEmpty($messages.MessageId))
                                                {
                                                    $messageCount = 1
                                                }
                                            }
                                            LogVerbose "Number of messages returned: $($messageCount)"

                                            if ($messageCount -gt 0)
                                            {

                                                $eidType = "0700" # eitLTPrivateMessage
                                                ForEach ($message in $messages)
                                                {
                                                    if ( ($message.FolderId.Length -lt 50) -or ($message.MessageId.Length -lt 50) )
                                                    {
                                                        Log "Invalid message details returned: FolderId = $($message.FolderId), MessageId = $($message.MessageId)" Red
                                                    }
                                                    else
                                                    {
                                                        $entryId = "00000000$($mailboxInstanceGuid)$eidType$($message.FolderId.Substring(2,48))$($message.MessageId.Substring(2,48))"
                                                        Log "$propId found on item: $entryId"
                                                        if (!$propIds.ContainsKey($entryId)) # Deduplication (for messages with more than one property)
                                                        {
                                                            $propIds.Add($entryId, $ewsPropId)
                                                            if ($Restart)
                                                            {
                                                                "$entryId;$ewsPropId" | out-file $entryIdsDumpFile -Append
                                                            }
                                                        }
                                                        else
                                                        {
                                                            # Here we have more than one named prop on a single item, so we need to dump the whole
                                                            # EntryId list again
                                                            $newProps = "$($propIds[$entryId]);$ewsPropId"
                                                            $propIds.Remove($entryId)
                                                            $propIds.Add($entryId, $newProps)
                                                            if ($Restart)
                                                            {
                                                                DumpEntryIdList $entryIdsDumpFile $propIds
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    Write-Progress -Activity "Searching for properties" -Completed
                                    DumpEntryIdList $entryIdsDumpFile $propIds $true
                                }
                            }
                        }
                    }
                    else
                    {
                        LogVerbose "Skipping $mbx.DisplayName"
                    }
                }
                Write-Progress -Activity "Retrieving named props" -Completed
            }
        }
        Write-Progress -Activity "Reading mailboxes" -Completed
        Log "Mailbox processing finished" Green
    }

    End
    {
        # Just to ensure that we don't change the path... I hate it when a script changes a path and doesn't set it back!
        cd $originalPath
    }
}