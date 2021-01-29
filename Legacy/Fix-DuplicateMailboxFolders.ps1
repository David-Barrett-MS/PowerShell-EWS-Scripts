#
# Fix-DuplicateMailboxFolders.ps1
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

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS (not required if -WhatIf is specified, or if default credentials are to be used).  These credentials will also be used to import an Exchange PowerShell session, if necessary.")]
    [System.Management.Automation.PSCredential]$Credentials,
				
	[Parameter(Mandatory=$False,HelpMessage="This parameter can be used to control whether ApplicationImpersonation rights are needed to access the mailbox (default is TRUE)")]	
    [bool]$Impersonate = $True,

	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if blank, then autodiscover is used; if not specified then default Office 365 Url is used)")]	
	[string]$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx",
	
	[Parameter(Mandatory=$False,HelpMessage="PowerShell Url (default is Office 365 Url: https://ps.outlook.com/powershell/)")]
    [String]$PowerShellUrl = "https://ps.outlook.com/powershell/",
    
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",

	[Parameter(Mandatory=$False,HelpMessage="If this parameter is present, then Merge-MailboxFolder.ps1 script will be called to attempt to eliminate the duplicate folder (by moving all items into the primary folder and then deleting the duplicate)")]	
    [switch]$Repair
)

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function LogVerbose([string]$Details)
{
    if ($VerbosePreference -eq "SilentlyContinue") { return }
    Write-Verbose $Details
	if ( $LogFile -eq "" ) { return	}
	$Details | Out-File $LogFile -Append
}

Function CmdletsAvailable()
{
    param (
        $RequiredCmdlets,
        $Silent = $False
    )

    $cmdletsAvailable = $True
    foreach ($cmdlet in $RequiredCmdlets)
    {
        if (Get-Command $cmdlet -ErrorAction SilentlyContinue)
        {
        }
        else
        {
            if (!$Silent) { Write-Host "Required cmdlet $cmdlet is not available" -ForegroundColor Red }
            $cmdletsAvailable = $False
            break
        }
    }

    return $cmdletsAvailable
}

Function ImportExchangeManagementSession()
{
    param (
        $RequiredCmdlets = "Get-Mailbox"
    )

    # Check we have Exchange Management Session available.  If not, we attempt to connect to and import one.
    if ( CmdletsAvailable $RequiredCmdlets $True )
    {
        # Cmdlets we need are available, so no need to import any session
        return
    }

    if ([String]::IsNullOrEmpty($PowerShellUrl))
    {
        Write-Host "PowerShell Url not specified and Exchange PowerShell session not available.  Cannot continue." -ForegroundColor Red
        exit
    }

    Write-Host "Attempting to connect to and import Exchange Management session" -ForegroundColor Gray
    $global:session = $null
    if ($Credentials -eq $null)
    {
        # No credentials specified, so we attempt to connect without specifying them (which will attempt to authenticate as the logged on user)
        $global:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PowerShellUrl -AllowRedirection 
    }
    else
    {
        # We have credentials, so we use them - we only use basic auth if the Url is https
        if (!$PowerShellUrl.ToLower().StartsWith("https"))
        {
            $global:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PowerShellUrl -Credential $Credentials -AllowRedirection 
        }
        else
        {
            # With HTTPS we use basic auth, as this is required for Office 365
            $global:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PowerShellUrl -Credential $Credentials -Authentication Basic -AllowRedirection 
        }
    }

    if ($global:session -eq $null)
    {
        Write-Host "Failed to open Exchange Administration session, cannot continue" -ForegroundColor Red
        exit
    }
    Write-Host "Exchange PowerShell session successfully established" -ForegroundColor Green
    Import-PSSession $global:session

    # Now check that we have the cmdlets we need available
    if ( CmdletsAvailable($RequiredCmdlets) )
    {
        return
    }

    exit
}

Function ConvertFolderIdToEntryId($folderId)
{
    # Get-MailboxFolderStatistics returns a modified EntryId as the FolderId
    # We need to decode it, and remove the first and last bytes to convert it to 
    # standard EntryId

    # Convert the id to a byte array
    $id = [System.Convert]::FromBase64String($folderId)
    Write-Host $id -ForegroundColor Gray

    # Create the real EntryId from the FolderId (i.e. copy everything except first and last bytes)
    [byte[]]$entryId = @()
    for ($i = 1; $i -lt $id.Length-1; $i++)
    {
        $entryId = $entryId + $id[$i]
    }

    # The id is now a standard EntryId, so just Base64 encode it again
    return [System.Convert]::ToBase64String($entryId)
}

Function ProcessFolder($realFolder)
{
    # Fix the given folder.  We find any other folders with the same name, move any contents to this folder, and then delete the duplicates

    # First of all check whether we have been passed a group of folders
    if ($realFolder.Count -gt 1)
    {
        foreach ($f in $realFolder)
        {
            ProcessFolder($f)
        }
        return
    }

    # Search the folder list for duplicates of this folder
    LogVerbose "Searching for duplicates: $($realFolder.Name)" -ForegroundColor Gray

    foreach ($folder in $script:folders)
    {
        if (($folder.Name -eq $realFolder.Name) -and ($folder.FolderPath -eq $realFolder.FolderPath) -and ($folder.FolderId -ne $realFolder.FolderId))
        {
            # This is a duplicate folder, so we want to merge it into the main folder, then delete it
            Log "Duplicate folder $($folder.Name) found: $($folder.FolderId)" Yellow
            $script:duplicateFolderFound = $true
            if ($Repair)
            {
                $targetId = ConvertFolderIdToEntryId($realFolder.FolderId)
                $sourceId = ConvertFolderIdToEntryId($folder.Id)
                if ($Impersonate)
                {
                    .\Merge-MailboxFolder.ps1 -SourceMailbox $Mailbox -MergeFolderList @{ $targetId = $sourceId } -ByEntryId -ProcessSubfolders -CreateTargetFolder -Delete -Impersonate -Credentials $Credentials -EwsUrl $EWSUrl -LogFile $LogFile
                }
                else
                {
                    .\Merge-MailboxFolder.ps1 -SourceMailbox $Mailbox -MergeFolderList @{ $targetId = $sourceId } -ByEntryId -ProcessSubfolders -CreateTargetFolder -Delete -Credentials $Credentials -EwsUrl $EWSUrl -LogFile $LogFile
                }
            }
        }
    }
}

Function ProcessMailbox($mbx)
{
    if ([String]::IsNullOrEmpty($mbx))
    {
        if (![String]::IsNullOrEmpty($Mailbox))
        {
            Log "Processing $($Mailbox)"
            $mbx = Get-Mailbox $Mailbox
        }
        else
        {
            return
        }
    }

    if ($mbx -eq $Null)
    {
        Log "Invalid mailbox" Red
        exit
    }

    # Retrieve the list of all mailbox folders (this is so we can identify the duplicates)
    $script:folders = $Null
    $script:folders = Get-MailboxFolderStatistics -Identity $mbx.Identity
    if ($script:folders -eq $Null)
    {
        Log "Failed to read mailbox folders for $mbx.Identity" Red
        exit
    }

    # Now process each folder and remove any duplicates
    LogVerbose "Searching mailbox $($mbx.PrimarySmtpAddress) for duplicate folders"
    $script:duplicateFolderFound = $false
    ForEach ($folder in $script:folders)
    {
        ProcessFolder($folder)
    }
    if (!$script:duplicateFolderFound)
    {
        Log "No duplicate folders found for $($mbx.PrimarySmtpAddress)" Green
    }
}

ImportExchangeManagementSession( @( "Get-Mailbox", "Get-MailboxFolderStatistics") )

if ([String]::IsNullOrEmpty($Mailbox))
{
    # No mailbox specified, so run a report against all of them
    $mbxs = Get-Mailbox -ResultSize Unlimited
    ForEach ($mbx in $mbxs) {
        Log "Processing $($mbx.PrimarySmtpAddress)"
        ProcessMailbox $mbx
    }
}
else
{
    ProcessMailbox ""
}