#
# Create-MailboxBatches.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2021. Use at your own risk.  No warranties are given.
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
	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with Exchange (PowerShell)")]
    [alias("Credentials")]
    [System.Management.Automation.PSCredential]$Credential,
				
	[Parameter(Mandatory=$False,HelpMessage="PowerShell Url (default is Office 365 Url: https://ps.outlook.com/powershell/)")]
    [String]$PowerShellUrl = "https://ps.outlook.com/powershell/",
    
    [Parameter(Mandatory=$False,HelpMessage="Same as Get-Mailbox -Filter parameter, use for filtering")]	
	$Filter = "*",

    [Parameter(Mandatory=$False,HelpMessage="Same as Get-Mailbox -OrganizationalUnit parameter, use for filtering")]	
	$OrganizationalUnit,

    [Parameter(Mandatory=$True,HelpMessage="Where the mailbox batch files will be created")]	
	[String]$ExportBatchPath,

    [Parameter(Mandatory=$False,HelpMessage="Maximum number of mailboxes per batch")]	
	[int]$BatchSize = 25
)


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

# Check export path
if ( !(Test-Path -Path $ExportBatchPath -PathType Container) )
{
    # Doesn't exist, we'll try to create it
    New-Item -ItemType Directory $ExportBatchPath -Force
    if ( !(Test-Path -Path $ExportBatchPath -PathType Container) )
    {
        Write-Host "Invalid export path: $ExportBatchPath" -ForegroundColor Red
        Exit
    }
}
if ( !($ExportBatchPath.EndsWith("\")) )
    { $ExportBatchPath = "$ExportBatchPath\" }

# Validate the availability of Get-Mailbox
ImportExchangeManagementSession( @( "Get-Mailbox") )

# Retrieve all mailboxes
if (![String]::IsNullOrEmpty($OrganizationalUnit))
{
    $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter $Filter -OrganizationalUnit $OrganizationalUnit
}
else
{
    $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter $Filter
}

# Now export the primary SMTP addresses of each mailbox to a file
$fileNum = 1
$userCount = 0
foreach ($mailbox in $mailboxes) {
    $primarySmtpAddress = $mailbox.PrimarySmtpAddress.Address
    if ([String]::IsNullOrEmpty($primarySmtpAddress))
    {
        $primarySmtpAddress = $mailbox.WindowsEmailAddress
    }
    Write-Verbose $primarySmtpAddress
    $primarySmtpAddress | Out-File "$ExportBatchPath\$fileNum.txt" -Append
    $userCount++
    if ($userCount -ge $BatchSize)
    {
        $userCount = 0
        $fileNum++
    }
}