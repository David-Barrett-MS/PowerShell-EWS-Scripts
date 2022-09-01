#
# Delete-ByInternetMessageId.ps1
#
# By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
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
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox (or mailboxes) to be searched.")]
	[ValidateNotNullOrEmpty()]
	$Mailbox,

	[Parameter(Mandatory=$True,HelpMessage="InternetMessageId being searched for")]
	[string]$InternetMessageId,
	
	[Parameter(Mandatory=$False,HelpMessage="If this switch is set, any matching items will be deleted (otherwise, no changes are made)")]
	[switch]$Delete,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")]
    [System.Management.Automation.PSCredential]$Credentials,

    [Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for MFA enabled accounts).")]
    [switch]$OAuth,

    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
    [string]$OAuthClientId = "",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id in which the application is registered.  If missing, application is assumed to be multi-tenant and the common log-in URL will be used.")]
    [string]$OAuthTenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$OAuthSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.  Please note that certificate auth requires the MSAL dll to be available.")]
    $OAuthCertificate = $null,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox.")]
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
	
	[Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file")]	
	[string]$TraceFile
)
$script:ScriptVersion = "1.0.2"


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
        Write-Verbose "Searching for DLL: $dllName"
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
            Write-Host "Unable to load locate $dll" -ForegroundColor Red
            return $false
        }
        else
        {
            try
            {
		        Write-Verbose ([string]::Format("Loading {2} v{0} found at: {1}", $dll.VersionInfo.FileVersion, $dll.VersionInfo.FileName, $dllName))
		        Add-Type -Path $dll.VersionInfo.FileName
                if ($dllLocations)
                {
                    $dllLocations.value += $dll.VersionInfo.FileName
                }
            }
            catch
            {
                return $false
            }
        }
    }
    return $true
}

function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!script:msalApiLoaded)
    {
        $msalLocation = @()
        $script:msalApiLoaded = $(LoadLibraries -searchProgramFiles $false -dllNames @("Microsoft.Identity.Client.dll") -dllLocations ([ref]$msalLocation))
        if (!$script:msalApiLoaded)
        {
            Write-Host "Failed to load MSAL.  Cannot continue with certificate authentication." -ForegroundColor Red
            exit
        }
    }   

    $cca1 = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca2 = $cca1.WithCertificate($OAuthCertificate)
    $cca3 = $cca2.WithTenantId($OAuthTenantId)
    $cca = $cca3.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$OAuthRedirectUri&response_mode=query&scope=openid%20profile%20email%20offline_access%20https://outlook.office365.com/.default"
    Write-Host "Please complete log-in via the web browser, and then paste the redirect URL (including auth code) here to continue" -ForegroundColor Green
    Start-Process $authUrl

    $authcode = Read-Host "Auth code"
    $codeStart = $authcode.IndexOf("?code=")
    if ($codeStart -gt 0)
    {
        $authcode = $authcode.Substring($codeStart+6)
    }
    $codeEnd = $authcode.IndexOf("&session_state=")
    if ($codeEnd -gt 0)
    {
        $authcode = $authcode.Substring(0, $codeEnd)
    }
    Write-Verbose "Using auth code: $authcode"

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$authcode;redirect_uri=$OAuthRedirectUri}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}

function GetTokenWithKey
{
    $Body = @{
      "grant_type"    = "client_credentials";
      "client_id"     = "$OAuthClientId";
      "client_secret" = "$OAuthSecretKey";
      "scope"         = "https://outlook.office365.com/.default"
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token: $Error" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}

function GetOAuthCredentials
{
    # Obtain OAuth token for accessing mailbox
    param (
        [switch]$RenewToken
    )
    $exchangeCredentials = $null

    if ($script:oauthToken -ne $null)
    {
        # We already have a token
        if ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1))
        {
            # Token still valid, so return that
            $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
            return $exchangeCredentials
        }

        # Token needs renewing

    }

    if (![String]::IsNullOrEmpty($OAuthSecretKey))
    {
        GetTokenWithKey
    }
    elseif ($OAuthCertificate -ne $null)
    {
        GetTokenWithCertificate
    }
    else
    {
        GetTokenViaCode
    }

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $script:services -eq $null ) { return }
    if ( $script:services.Count -lt 1 ) { return }
    if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    Write-Verbose "OAuth access token invalid, attempting to renew"
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($exchangeCredentials -eq $null) { return }
    if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -le [DateTime]::Now )
    { 
        Write-Host "OAuth Token renewal failed" -ForegroundColor Red
        exit # We no longer have access to the mailbox, so we stop here
    }

    Write-Host "OAuth token successfully renewed; new expiry: $($script:oAuthToken.ExpiresOn)" -ForegroundColor Green
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($exchangeCredentials)
        }
        Write-Verbose "Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
    }
}

Function LoadEWSManagedAPI
{
	# Find and load the managed API
    $ewsApiLocation = @()
    $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))

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
	        }
        }
    }

    if ($ewsApiLoaded)
    {
        if ($ewsApiLocation[0])
        {
            Write-Host "Using EWS Managed API found at: $($ewsApiLocation[0])" -ForegroundColor Gray
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

Function CreateTraceListener($service)
{
    # Create trace listener to capture EWS conversation (useful for debugging)

    if ([String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        Write-Host "Managed API path missing; unable to create tracer" -ForegroundColor Red
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
            Write-Host "Tracing to: $TraceFile"
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

        $script:Tracer = $null
        $TraceCompilation=$Provider.CompileAssemblyFromSource($Params,$TraceListenerClass)
        If ($TraceCompilation)
        {
            $TraceAssembly=$TraceCompilation.CompiledAssembly
            $script:Tracer=$TraceAssembly.CreateInstance("TraceListener.EWSTracer")
            # Attach the trace listener to the Exchange service
            $service.TraceListener = $script:Tracer
        }
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
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

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
            Write-Verbose "Applying given credentials: $($Credentials.UserName)"
            $exchangeService.Credentials = $Credentials.GetNetworkCredential()
        }
        else
        {
	        Write-Verbose "Using default credentials"
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
		    Write-Verbose "Performing autodiscover for $smtpAddress"
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
			    Write-Host "$smtpAddress : autodiscover failed" -ForegroundColor Red
			    return $Null
		    }
		    Write-Verbose "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
            Write-Host "$smtpAddress : error occurred during autodiscover: $($Error[0])" -ForegroundColor Red
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
        if ($script:Tracer)
        {
            $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
            $exchangeService.TraceEnabled = $True
        }
        else
        {
            Write-Host "Failed to create EWS trace listener.  Throttling back-off time won't be detected." -ForegroundColor Yellow
        }
    }

    $script:services.Add($smtpAddress, $exchangeService)
    Write-Verbose "Currently caching $($script:services.Count) ExchangeService objects"
    return $exchangeService
}

Function ProcessFolder()
{
	# Process this folder

    $Folder = $args[0]
	if ($Folder -eq $null)
	{
		throw "No folder specified"
	}
	
    Write-Verbose "Reading subfolders of: $($Folder.DisplayName)"

	# Process any subfolders
	if ($Folder.ChildFolderCount -gt 0)
	{
        # We read the list of all folders first, so that we have the complete list before any processing
        $subfolders = @()
        $moreFolders = $True
		$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
        while ($moreFolders)
        {
            ApplyEWSOAuthCredentials
			$FindFoldersResults = $Folder.FindFolders($FolderView)
            $subfolders += $FindFoldersResults.Folders
            $moreFolders = $FindFoldersResults.MoreAvailable
            $FolderView.Offset += 500
        }
        # Process the subfolders
        if ($subfolders.Count -gt 0)
        {
            Write-Verbose "$($Folder.DisplayName) contains $($subfolders.Count) subfolders"
			ForEach ($subFolder in $subfolders)
			{
				ProcessFolder $subFolder
			}
        }
	}

    # Search for the item with given InternetMessageId
    Write-Verbose "Searching for message in: $($Folder.DisplayName)"
	$Offset=0
	$PageSize=1000
	$MoreItems=$true

    $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId, $InternetMessageId)
    $deletedFromFolder = 0

	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning)
        $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly

        try
        {
            $FindResults=$Folder.FindItems($SearchFilter, $View)
            $Offset+=$PageSize
            $MoreItems = $FindResults.MoreAvailable
            if ($FindResults)
            {
                foreach ($item in $FindResults.Items)
                {
                    if ($Delete)
                    {
                        $item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                    }
                    $deletedFromFolder++
                }
            }
        }
        catch
        {
            $MoreItems = $false
        }
		
	}
    if ($deletedFromFolder -gt 0)
    {
        if ($Delete)
        {
            Write-Host "$deletedFromFolder items matched and were deleted from folder $($Folder.DisplayName)" -ForegroundColor Green
        }
        else
        {
            Write-Host "$deletedFromFolder items matched in folder $($Folder.DisplayName)" -ForegroundColor Green
        }
    }
    else
    {
        Write-Verbose "0 items matched in folder $($Folder.DisplayName)"
    }
    $script:itemsDeleted += $deletedFromFolder
}

function ProcessMailbox($TargetMailbox)
{
    # Process the mailbox

    Write-Host ([string]::Format("Processing mailbox {0}", $TargetMailbox)) -ForegroundColor Gray

	$script:Service = CreateService($TargetMailbox)
	if ($script:Service -eq $Null)
	{
		Write-Host "Failed to connect to mailbox $TargetMailbox" -ForegroundColor Red
        return
	}
    
    # Bind to mailbox root folder
    $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $TargetMailbox )
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx )
    $mailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:service, $folderId)

    if ( $mailboxRoot -eq $null )
    {
        Write-Host "Failed to open message store ($TargetMailbox)" -ForegroundColor Red
        if ($Impersonate)
        {
            Write-Host "Please check that you have impersonation permissions" -ForegroundColor Yellow
        }
        return
    }

    # To search the whole mailbox, we need to recurse each folder and perform the search.  We recurse from the root folder to do this.
    $script:itemsDeleted = 0
    ProcessFolder $mailboxRoot

    if ($Delete)
    {
        Write-Host "$($script:itemsDeleted) items matched and were deleted from $TargetMailbox" -ForegroundColor Green
    }
    else
    {
        Write-Host "$($script:itemsDeleted) items matched (but were not deleted as -Delete not set) from $TargetMailbox" -ForegroundColor Yellow
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
  

# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
    LogVerbose "Reading mailboxes from CSV file"
	$csv = Import-CSV $SourceMailbox -Header "PrimarySmtpAddress"
	foreach ($entry in $csv)
	{
        LogVerbose $entry.PrimarySmtpAddress
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress))
        {
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress"))
            {
		        $Mailbox = $entry.PrimarySmtpAddress
			    ProcessMailbox $Mailbox
            }
        }
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox $Mailbox
}

if ($script:Tracer -ne $null)
{
    $script:Tracer.Close()
}
