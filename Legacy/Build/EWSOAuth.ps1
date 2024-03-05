#
# EWSOAuth.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2022. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

param (

#>** EWS/OAUTH PARAMETERS START **#
    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS.")]
    [alias("Credential")]
    [System.Management.Automation.PSCredential]$Credentials,
	
    [Parameter(Mandatory=$False,HelpMessage="If set, then we will use OAuth to access the mailbox (required for Office 365)")]
    [switch]$OAuth,

    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in Azure AD.")]
    [string]$OAuthClientId = "8799ab60-ace5-4bda-b31f-621c9f6668db",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id (application must be registered in the same tenant being accessed).")]
    [string]$OAuthTenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$OAuthSecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.  Certificate auth requires MSAL libraries to be available.")]
    $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="If set, OAuth tokens will be stored in global variables for access in other scripts/console.  These global variable will be checked by later scripts using delegate auth to prevent additional log-in prompts.")]	
    [switch]$GlobalTokenStorage,

    [Parameter(Mandatory=$False,HelpMessage="For debugging purposes.")]
    [switch]$OAuthDebug,

    [Parameter(Mandatory=$False,HelpMessage="A value greater than 0 enables token debugging (specify total number of token renewals to debug).")]	
    $DebugTokenRenewal = 0,

    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox.")]
    [switch]$Impersonate,
	
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used).")]	
    [string]$EwsUrl,

    [Parameter(Mandatory=$False,HelpMessage="If specified, requests are directed to Office 365 endpoint (this overrides -EwsUrl).")]
    [switch]$Office365,
	
    [Parameter(Mandatory=$False,HelpMessage="If specified, only TLS 1.2 connections will be negotiated.")]
    [switch]$ForceTLS12,
	
    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed).")]	
    [string]$EWSManagedApiPath = "",
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate).")]	
    [switch]$IgnoreSSLCertificate,
	
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing AutoDiscover.")]	
    [switch]$AllowInsecureRedirection,

    [Parameter(Mandatory=$False,HelpMessage="User-Agent header that will be set on ExchangeService and AutodiscoverService objects.")]	
    [string]$UserAgent = "https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts",

    [Parameter(Mandatory=$False,HelpMessage="Trace file - if specified, EWS tracing information is written to this file.")]	
    [string]$TraceFile,
#>** EWS/OAUTH PARAMETERS END **#

    [Parameter(Mandatory=$False,HelpMessage="This is a dummy parameter to avoid syntax error.")]
    $Dummy = $null
)




#>** EWS/OAUTH FUNCTIONS START **#

# These functions are common for all my EWS scripts and are injected as part of the build/publish process.  Changes should be made to EWSOAuth.ps1 code snippet, not the script being run.
# EWS/OAuth library version: 1.0.5

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
            if ($null -eq $dll)
            {
                Log "$dllName not found in current directory - searching Program Files folders" Yellow
	            $dll = Get-ChildItem -Recurse "C:\Program Files (x86)" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            if (!$dll)
	            {
		            $dll = Get-ChildItem -Recurse "C:\Program Files" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            }
            }
        }

        if ($null -eq $dll)
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

function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!$script:msalApiLoaded)
    {
        $msalLocation = @()
        $script:msalApiLoaded = $(LoadLibraries -searchProgramFiles $false -dllNames @("Microsoft.Identity.Client.dll") -dllLocations ([ref]$msalLocation))
        if (!$script:msalApiLoaded)
        {
            Log "Failed to load MSAL.  Cannot continue with certificate authentication." Red
            exit
        }
    }   

    $cca = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca = $cca.WithCertificate($OAuthCertificate)
    $cca = $cca.WithTenantId($OAuthTenantId)
    $cca = $cca.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    LogVerbose "Requesting token using certificate auth"
    
    try
    {
        $script:oauthToken = $acquire.ExecuteAsync().Result
    }
    catch
    {
        Log "Failed to obtain OAuth token: $Error" Red
        exit # Failed to obtain a token
    }

    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    if ($null -ne $script:oAuthAccessToken)
    {
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
        $script:Impersonate = $true
        return
    }

    # If we get here, we don't have a token so can't continue
    Log "Failed to obtain OAuth token (no error thrown)" Red
    exit
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$OAuthRedirectUri&response_mode=query&prompt=select_account&scope=openid%20profile%20email%20offline_access%20https://outlook.office365.com/.default"
    Write-Host "Please complete log-in via the web browser, and then copy the redirect URL (including auth code) to the clipboard to continue" -ForegroundColor Green
    Set-Clipboard -Value "Waiting for auth code"
    Start-Process $authUrl

    do
    {
        $authcode = Get-Clipboard
        Start-Sleep -Milliseconds 250
    } while ($authCode -eq "Waiting for auth code")

    $codeStart = $authcode.IndexOf("?code=")
    if ($codeStart -gt 0)
    {
        $authcode = $authcode.Substring($codeStart+6)
        $codeEnd = $authcode.IndexOf("&session_state=")
        if ($codeEnd -gt 0)
        {
            $authcode = $authcode.Substring(0, $codeEnd)
        }
        LogVerbose "Using auth code: $authcode"
        Write-Host "Auth code acquired, attempting to obtain access token" -ForegroundColor Green
    }
    else
    {
        throw "Failed to obtain Auth code from clipboard"
    }

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$authcode;redirect_uri=$OAuthRedirectUri}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
        return
    }
    catch {}

    throw "Failed to obtain OAuth token"
}

function RenewOAuthToken
{
    # Renew the delegate token (original token was obtained by auth code, but we can now renew using the access token)
    if (!$script:oAuthToken)
    {
        # We don't have a token, so we can't renew
        GetTokenViaCode
        return
    }

    $body = @{grant_type="refresh_token";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;refresh_token=$script:oauthToken.refresh_token}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Write-Host "Failed to renew OAuth token (auth code grant)" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}

function GetTokenWithKey
{
    $Body = @{
      "grant_type"    = "client_credentials";
      "client_id"     = "$OAuthClientId";
      "scope"         = "https://outlook.office365.com/.default"
    }

    if ($null -ne $script:oAuthToken)
    {
        # If we have a refresh token, add that to our request body and change grant type
        if (![String]::IsNullOrEmpty($script:oAuthToken.refresh_token))
        {
            $Body.Add("refresh_token", $script:oAuthToken.refresh_token)
            $Body["grant_type"] = "refresh_token"
        }
    }
    if ($Body["grant_type"] -eq "client_credentials")
    {
        # To obtain our first access token we need to use the secret key
        $Body.Add("client_secret", $OAuthSecretKey)
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        $script:oauthTokenAcquireTime = [DateTime]::UtcNow
    }
    catch
    {
        Log "Failed to obtain OAuth token: $Error" Red
        exit # Failed to obtain a token
    }
    $script:Impersonate = $true
}

function JWTToPSObject
{
    param([Parameter(Mandatory=$true)][string]$token)

    $tokenheader = $token.Split(".")[0].Replace('-', '+').Replace('_', '/')
    while ($tokenheader.Length % 4) { $tokenheader = "$tokenheader=" }    
    #$tokenHeaderObject = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    while ($tokenPayload.Length % 4) { $tokenPayload = "$tokenPayload=" }
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)
    $tokenObject = $tokenArray | ConvertFrom-Json
    return $tokenObject
}

function LogOAuthTokenInfo
{
    if ($null -eq $global:OAuthAccessToken)
    {
        Log "No OAuth token obtained." Red
        return
    }

    $idToken = $null
    if (-not [String]::IsNullOrEmpty($global:OAuthAccessToken.id_token))
    {
        $idToken = $global:OAuthAccessToken.id_token
    }
    elseif (-not [String]::IsNullOrEmpty($global:OAuthAccessToken.IdToken))
    {
        $idToken = $global:OAuthAccessToken.IdToken
    }

    if ([String]::IsNullOrEmpty($idToken))
    {
        Log "OAuth ID token not present" Yellow
    }
    else
    {
        $global:idTokenDecoded = JWTToPSObject($idToken)
        Log "OAuth ID Token (`$idTokenDecoded):" Yellow
        Log $global:idTokenDecoded Yellow
    }

    if (-not [String]::IsNullOrEmpty($global:OAuthAccessToken))
    {
        $global:accessTokenDecoded = JWTToPSObject($global:OAuthAccessToken)
        Log "OAuth Access Token (`$accessTokenDecoded):" Yellow
        Log $global:accessTokenDecoded Yellow
    }
    else
    {
        Log "OAuth access token not present" Red
    }
}

function GetOAuthCredentials
{
    # Obtain OAuth token for accessing mailbox
    param (
        [switch]$RenewToken
    )
    $exchangeCredentials = $null

    if ($null -ne $script:oauthToken)
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
    elseif ($null -ne $OAuthCertificate)
    {
        GetTokenWithCertificate
    }
    else
    {
        if ($RenewToken)
        {
            RenewOAuthToken
        }
        else
        {
            if ($GlobalTokenStorage -and $null -eq $script:oauthToken)
            {
                # Check if we have token variable set globally
                if ($global:oAuthPersistAppId -eq $OAuthClientId)
                {
                    $script:oAuthToken = $global:oAuthPersistToken
                    $script:oauthTokenAcquireTime = $global:oAuthPersistTokenAcquireTime
                }
                RenewOAuthToken
            }
            else
            {
                GetTokenViaCode
            }
        }
    }

    if ($GlobalTokenStorage -or $OAuthDebug)
    {
        # Store the OAuth in a global variable for later access
        $global:oAuthPersistToken = $script:oAuthToken
        $global:oAuthPersistAppId = $OAuthClientId
        $global:oAuthPersistTokenAcquireTime = $script:oauthTokenAcquireTime
    } 

    if ($OAuthDebug)
    {
        LogVerbose "`$oAuthPersistToken contains token response"
        $global:OAuthAccessToken = $script:oAuthAccessToken
        LogVerbose "`$OAuthAccessToken: `r`n$($global:OAuthAccessToken)"
        LogOAuthTokenInfo
    }

   

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

$script:oAuthDebugStop = $false
$script:oAuthDebugStopCount = 0
function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $null -eq $script:services ) { return }

    
    if ($DebugTokenRenewal -gt 0 -and $script:oauthToken)
    {
        # When debugging tokens, we stop after on every other EWS call and wait for the token to expire
        if ($script:oAuthDebugStop)
        {
            # Wait until token expires (we do this after every call when debugging OAuth)
            # Access tokens can't be revoked, but a policy can be assigned to reduce lifetime to 10 minutes: https://learn.microsoft.com/en-us/graph/api/resources/tokenlifetimepolicy?view=graph-rest-1.0
            if ( $null -ne $OAuthCertificate)
            {
                $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
            }
            else
            {
                $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
            }
            $timeUntilExpiry = $tokenExpire.Subtract([DateTime]::UtcNow).TotalSeconds
            if ($timeUntilExpiry -gt 0)
            {
                Write-Host "Waiting until token has expired: $tokenExpire (UTC)" -ForegroundColor Cyan
                Start-Sleep -Seconds $tokenExpire.Subtract([DateTime]::UtcNow).TotalSeconds
            }
            Write-Host "Token expired, continuing..." -ForegroundColor Cyan
            $oAuthDebugStop = $false
            $script:oAuthDebugStopCount++
        }
        else
        {
            if ($DebugTokenRenewal-$script:oAuthDebugStopCount -gt 0)
            {
                $script:oAuthDebugStop = $true
            }
        }
    }
    
    if ($null -ne $OAuthCertificate)
    {
        if ( [DateTime]::UtcNow -lt $script:oauthToken.ExpiresOn.UtcDateTime) { return }
    }
    elseif ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    LogVerbose("OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($null -eq $exchangeCredentials) { return }

    if ($null -ne $OAuthCertificate)
    {
        $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
        if ( [DateTime]::UtcNow -ge $tokenExpire)
        {
            Log "OAuth Token renewal failed (certificate auth)"
            exit # We no longer have access to the mailbox, so we stop here
        }
    }
    else
    {
        if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -lt [DateTime]::UtcNow )
        { 
            Log "OAuth Token renewal failed"
            exit # We no longer have access to the mailbox, so we stop here
        }
        $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
    }

    Log "OAuth token successfully renewed; new expiry: $tokenExpire"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = $exchangeCredentials
        }
        LogVerbose "Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
    }
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
        # We can only set any EWS properties once the API is loaded

        $script:PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6698, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) # PR_REPLICA_LIST is required to perform AutoDiscover requests for public folders
        $script:PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3601, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    }

    return $ewsApiLoaded
}

Function CurrentUserPrimarySmtpAddress()
{
    # Attempt to retrieve the current user's primary SMTP address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $result = $searcher.FindOne()

    if ($null -ne $result)
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

    Add-Type -TypeDefinition $TASource -ReferencedAssemblies "System.DLL"

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=[Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll]::new()
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

Function SetClientRequestId($exchangeService)
{
    # Apply unique client-request-id to the EWS service object

    $exchangeService.ClientRequestId = (New-Guid).ToString()
}

Function CreateTraceListener($exchangeService)
{
    # Create trace listener to capture EWS conversation (useful for debugging)

    if ([String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        Log "Managed API path missing; unable to create tracer" Red
        Exit
    }

    if ($null -eq $script:Tracer)
    {
        $traceFileForCode = ""

        if (![String]::IsNullOrEmpty($TraceFile))
        {
            $traceFileForCode = $traceFile.Replace("\", "\\")
        }

        $TraceListenerClass = @"
		    using System;
		    using System.Text;
		    using System.IO;
		    using System.Threading;
		    using Microsoft.Exchange.WebServices.Data;
		
		    public class EWSTracer: Microsoft.Exchange.WebServices.Data.ITraceListener
		    {
			    private StreamWriter _traceStream = null;
                private string _lastResponse = String.Empty;
                private string _traceFileFullPath = String.Empty;

			    public EWSTracer(string traceFileName = "$traceFileForCode" )
			    {
				    try
				    {
                        if (!String.IsNullOrEmpty(traceFileName))
					        _traceStream = File.AppendText(traceFileName);
                        FileInfo fi = new FileInfo(traceFileName);
                        _traceFileFullPath = fi.Directory.FullName + "\\" + fi.Name;
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

					try
					{
						_traceStream.WriteLine(traceMessage);
						_traceStream.Flush();
					}
					catch { }
			    }

                public string LastResponse
                {
                    get { return _lastResponse; }
                }

                public string TraceFileFullPath
                {
                    get { return _traceFileFullPath; }
                }
		    }
"@

        if ("EWSTracer" -as [type]) {} else {
            Add-Type -TypeDefinition $TraceListenerClass -ReferencedAssemblies $EWSManagedApiPath
        }
        $script:Tracer=[EWSTracer]::new($traceFileForCode)

        # Attach the trace listener to the Exchange service
        $exchangeService.TraceListener = $script:Tracer
        if (![String]::IsNullOrEmpty($TraceFile))
        {
            Log "Tracing to: $($script:Tracer.TraceFileFullPath)"
        }
    }
}

function CreateService($smtpAddress, $impersonatedAddress = "")
{
    # Creates and returns an ExchangeService object to be used to access mailboxes

    # First of all check to see if we have a service object for this mailbox already
    if ($null -eq $script:services)
    {
        $script:services = @{}
    }
    if ($script:services.ContainsKey($smtpAddress))
    {
        return $script:services[$smtpAddress]
    }

    # Create new service
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    $exchangeService.UserAgent = $UserAgent

    # Do we need to use OAuth?
    if ($Office365) { $OAuth = $true }
    if ($OAuth)
    {
        $exchangeService.Credentials = GetOAuthCredentials
        if ($null -eq $exchangeService.Credentials)
        {
            # OAuth failed
            return $null
        }
    }
    else
    {
        # Set credentials if specified, or use logged on user.
        if ($null -ne $Credentials)
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
 
    if ([String]::IsNullOrEmpty($impersonatedAddress))
    {
        $impersonatedAddress = $smtpAddress
    }
    $exchangeService.HttpHeaders.Add("X-AnchorMailbox", $smtpAddress)
    if ($Impersonate)
    {
        $exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $impersonatedAddress)
    }

    # We enable tracing so that we can retrieve the last response (and read any throttling information from it - this isn't exposed in the EWS Managed API)
    if (![String]::IsNullOrEmpty($EWSManagedApiPath))
    {
        CreateTraceListener $exchangeService
        if ($script:Tracer)
        {
            $exchangeService.TraceListener = $script:Tracer
            $exchangeService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
            $exchangeService.TraceEnabled = $True
        }
        else
        {
            Log "Failed to create EWS trace listener.  Throttling back-off time won't be detected." Yellow
        }
    }

    SetClientRequestId $exchangeService
    $exchangeService.ReturnClientRequestId = $true

    $script:services.Add($smtpAddress, $exchangeService)
    LogVerbose "Currently caching $($script:services.Count) ExchangeService objects" $true
    return $exchangeService
}

Function CreateAutoDiscoverService($exchangeService)
{
    $autodiscover = new-object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService(new-object Uri("https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"), ExchangeVersion.Exchange2016)
    $autodiscover.Credentials = $exchangeService.Credentials
    $autodiscover.TraceListener = $exchangeService.TraceListener
    $autodiscover.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
    $autodiscover.TraceEnabled = $true
    $autodiscover.UserAgent = $UserAgent
    return $autodiscover;
}

function SetPublicFolderHeirarchyHeaders($ExchangeService, $AutodiscoverAddress)
{
    # Sets the X-PublicFolderMailbox and X-AnchorMailbox properties for a request to public folders

    if (!$OAuth -and !$Office365)
    {
        LogDebug "Public folder headers not implemented for on-premises Exchange"
        return $null        
    }

    if ($null -eq $script:publicFolderHeirarchyHeaders)
    {
        # We keep a cache of known folders to avoid unnecessary AutoDiscover
        $script:publicFolderHeirarchyHeaders = New-Object 'System.Collections.Generic.Dictionary[String,String[]]'
    }

    $xAnchorMailbox = ""
    $xPublicFolderMailbox = ""
    if ($script:publicFolderHeirarchyHeaders.ContainsKey($AutodiscoverAddress))
    {
        $xAnchorMailbox = $script:publicFolderHeirarchyHeaders[$AutodiscoverAddress][0]
        $xPublicFolderMailbox = $script:publicFolderHeirarchyHeaders[$AutodiscoverAddress][1]
    }    
    else {
        $autoDiscoverService = CreateAutoDiscoverService $ExchangeService
        [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[]]$userSettingsRequired = @([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer)
        $userSettings = $autoDiscoverService.GetUserSettings($AutodiscoverAddress, $userSettingsRequired)
    
        if ($null -eq $userSettings)
        {
            LogVerbose "Failed to obtain Autodiscover public folder settings"
            return
        }
    
        $xAnchorMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation]
        if ([String]::IsNullOrEmpty($xAnchorMailbox))
        {
            LogVerbose "PublicFolderInformation not present in Autodiscover response"
            return
        }
        LogVerbose "Public folder heirarchy X-AnchorMailbox set to $xAnchorMailbox"
    
        # Now we need to retrieve the X-PublicFolderMailbox value
        $userSettings = $autoDiscoverService.GetUserSettings($xAnchorMailbox, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer)
        $xPublicFolderMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer]
        if ([String]::IsNullOrEmpty($xPublicFolderMailbox))
        {
            LogVerbose "PublicFolderInformation not present in Autodiscover response fpr $xAnchorMailbox"
            return
        }
        LogVerbose "Public folder heirarchy X-PublicFolderMailbox set to $xPublicFolderMailbox"

        $script:publicFolderHeirarchyHeaders.Add($AutodiscoverAddress, @($xAnchorMailbox, $xPublicFolderMailbox))
    }

    if ($ExchangeService.HttpHeaders.ContainsKey("X-PublicFolderMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-PublicFolderMailbox") | out-null
    }
    if ($ExchangeService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-AnchorMailbox") | out-null
    }
    $ExchangeService.HttpHeaders.Add("X-PublicFolderMailbox", $xPublicFolderMailbox) | out-null
    $ExchangeService.HttpHeaders.Add("X-AnchorMailbox", $xAnchorMailbox) | out-null
}

Function SetPublicFolderContentHeaders($ExchangeService, $PublicFolder)
{
    # Sets the X-PublicFolderMailbox and X-AnchorMailbox properties for a content request to public folders

    if (!$OAuth -and !$Office365)
    {
        LogDebug "Public folder headers not implemented for on-premises Exchange"
        return $null        
    }

    if ($null -eq $script:publicFolderContentHeaders)
    {
        # We keep a cache of known folders to avoid unnecessary AutoDiscover
        $script:publicFolderContentHeaders = New-Object 'System.Collections.Generic.Dictionary[String,String]'
    }

    $xPublicFolderMailbox = ""
    if ($script:publicFolderContentHeaders.ContainsKey($PublicFolder.FolderId.UniqueId))
    {
        $xPublicFolderMailbox = $script:publicFolderContentHeaders[$PublicFolder.FolderId.UniqueId]
    }
    else
    {
        # We need to perform an AutoDiscover request to obtain the correct header value        
        $replicaGuid = ""
        foreach ($extendedProperty in $PublicFolder.ExtendedProperties)
        {
            if ($extendedProperty.PropertyDefinition -eq $script:PR_REPLICA_LIST)
            {
                #[byte[]]$ByteArr = ([byte[]])$extendedProperty.Value;
                $replicaGuid =[System.Text.Encoding]::ASCII.GetString($extendedProperty.Value, 0, 36)
                break
            }
        }
        if ([String]::IsNullOrEmpty($replicaGuid))
        {
            LogVerbose "Public folder PR_REPLICA_LIST not present"
            return
        }

        # Work out the AutoDiscover address from the replica GUID and domain
        if ($null -eq $Mailbox -and $null -ne $SourceMailbox)
        {
            $Mailbox = $SourceMailbox
        }
        $domainStart = $Mailbox.IndexOf("@")
        if ($domainStart -lt 0)
        {
            Log "Invalid mailbox: $Mailbox" Red
            return
        }
        $autoDiscoverAddress = "$replicaGuid$($Mailbox.Substring($domainStart))"
        LogVerbose "AutoDiscover address for $autoDiscoverAddress to access public folder $($PublicFolder.DisplayName)"
        
        $autoDiscoverService = CreateAutoDiscoverService $ExchangeService
        if ($null -eq $autoDiscoverService)
        {
            LogVerbose "Failed to create AutoDiscover service"
            return
        }

        [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[]]$userSettingsRequired = @([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress)
        $userSettings = $autoDiscoverService.GetUserSettings($autodiscoverAddress, $userSettingsRequired)

        if ($null -eq $userSettings)
        {
            LogVerbose "Failed to obtain user settings for $autoDiscoverAddress"
            return
        }

        $xPublicFolderMailbox = $userSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress]
        if ([String]::IsNullOrEmpty($xPublicFolderMailbox))
        {
            LogVerbose "Failed to obtain AutoDiscoverSMTPAddress for $autoDiscoverAddress"
            return
        }

        $script:publicFolderContentHeaders.Add($PublicFolder.FolderId.UniqueId, $xPublicFolderMailbox)
        LogVerbose "Caching X-PublicFolderMailbox for folder $($PublicFolder.DisplayName): $xPublicFolderMailbox"
    }

    # Both X-PublicFolderMailbox and X-AnchorMailbox are required for public folder content requests, but they have the same value
    if ($ExchangeService.HttpHeaders.ContainsKey("X-PublicFolderMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-PublicFolderMailbox") | out-null
    }
    if ($ExchangeService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
    {
        $ExchangeService.HttpHeaders.Remove("X-AnchorMailbox") | out-null
    }
    $ExchangeService.HttpHeaders.Add("X-PublicFolderMailbox", $xPublicFolderMailbox) | out-null
    $ExchangeService.HttpHeaders.Add("X-AnchorMailbox", $xPublicFolderMailbox) | out-null
    LogVerbose "Set X-PublicFolderMailbox and X-AnchorMailbox to $xPublicFolderMailbox"
}

Function Throttled()
{
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning
    if ( $null -eq $script:Tracer -or [String]::IsNullOrEmpty($script:Tracer.LastResponse))
    {
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds")
    {
        # We are throttled, and the server has told us how long to back off for
        $backOffMilliseconds = [int]::Parse($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text")
        $resumeTime = [DateTime]::Now.AddMilliseconds($backOffMilliseconds)
        Log "Back off for $backOffMilliseconds milliseconds (will resume at $($resumeTime.ToLongTimeString()))" Yellow
        Start-Sleep -Milliseconds $backOffMilliseconds
        Log "Resuming operations" Gray
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

    if ($null -eq $folderId)
    {
        Log "FolderId missing" Red
        return $null
    }

    LogVerbose "Attempting to bind to folder $folderId"
    $folder = $null
    if ($null -eq $exchangeService)
    {
        $exchangeService = $script:service
        if ($null -eq $exchangeService)
        {
            Log "No ExchangeService object set" Red
            return $null
        }
    }

    if ($null -eq $script:requiredFolderProperties)
    {
        # If scripts require a custom property set, this variable should be set before calling this function.  If it's missing at this point, we just retrieve the most useful folder properties
        $script:requiredFolderProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount, $script:PR_FOLDER_TYPE)
        $script:requiredFolderProperties.Add($script:PR_REPLICA_LIST) # Required for public folder Autodiscover
    }    
    if ($null -eq $propset)
    {
        $propset = $script:requiredFolderProperties
    }

    try
    {
        ApplyEWSOauthCredentials
        SetClientRequestId $exchangeService
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
        if (!($null -eq $folder))
        {
            LogVerbose "Successfully bound to folder $folderId"
        }
        return $folder
    }
    catch {}

    if (Throttled)
    {
        try
        {
            ApplyEWSOauthCredentials
            SetClientRequestId $exchangeService
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propset)
            if (!($null -eq $folder))
            {
                LogVerbose "Successfully bound to folder $folderId"
            }
            return $folder
        }
        catch {}
    }

    # If we get to this point, we have been unable to bind to the folder
    ReportError "ThrottledFolderBind"
    LogVerbose "FAILED to bind to folder $folderId"
    return $null
}

#>** EWS/OAUTH FUNCTIONS END **#