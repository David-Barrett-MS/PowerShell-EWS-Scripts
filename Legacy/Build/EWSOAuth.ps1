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

#>** AUTHENTICATION PARAMETERS START **#
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

    [Parameter(Mandatory=$False,HelpMessage="For debugging purposes.")]
    [switch]$OAuthDebug,

    [Parameter(Mandatory=$False,HelpMessage="Enables token debugging (for development purposes)")]	
    [switch]$DebugTokenRenewal,
#>** AUTHENTICATION PARAMETERS END **#

    [Parameter(Mandatory=$False,HelpMessage="This is a dummy parameter to avoid syntax error.")]
    $Dummy = $null
)




#>** OAUTH FUNCTIONS START **#
function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!script:msalApiLoaded)
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
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
    $script:Impersonate = $true
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
        $script:AuthCode = $authcode.Substring(0, $codeEnd)
    }
    Write-Verbose "Using auth code: $authcode"

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$OAuthClientId;code=$script:AuthCode;redirect_uri=$OAuthRedirectUri}
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

    if ($script:oAuthToken -ne $null)
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
    $tokenHeaderObject = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    while ($tokenPayload.Length % 4) { $tokenPayload = "$tokenPayload=" }
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)
    $tokenObject = $tokenArray | ConvertFrom-Json
    return $tokenObject
}

function LogOAuthTokenInfo
{
    if ($global:OAuthResponse -eq $null)
    {
        Log "No OAuth token obtained." Red
        return
    }

    if ([String]::IsNullOrEmpty($global:OAuthResponse.id_token))
    {
        Log "OAuth ID Token not present" Yellow
    }
    else
    {
        $global:idTokenDecoded = JWTToPSObject($global:OAuthResponse.id_token)
        Log "OAuth ID Token (`$idTokenDecoded):" Yellow
        Log $global:idTokenDecoded Yellow
    }

    $global:accessTokenDecoded = JWTToPSObject($global:OAuthResponse.access_token)
    Log "OAuth Access Token (`$accessTokenDecoded):" Yellow
    Log $global:accessTokenDecoded Yellow
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
        if ($RenewToken)
        {
            RenewOAuthToken
        }
        else
        {
            GetTokenViaCode
        }
    }
    if ($OAuthDebug)
    {
        $global:OAuthResponse = $script:oAuthToken
        LogVerbose "`$oAuthResponse contains token response"
        $global:OAuthAccessToken = $script:oAuthAccessToken
        LogVerbose "`$OAuthAccessToken: `r`n$($global:OAuthAccessToken)"
        LogOAuthTokenInfo
    }
    

    # If we get here we have a valid token
    $exchangeCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oAuthAccessToken)
    return $exchangeCredentials
}

$script:oAuthDebugStop = $false
function ApplyEWSOAuthCredentials
{
    # Apply EWS OAuth credentials to all our service objects

    if ( -not $OAuth ) { return }
    if ( $script:services -eq $null ) { return }

    
    if ($DebugTokenRenewal -and $script:oauthToken)
    {
        # When debugging tokens, we stop after on every other EWS call and wait for the token to expire
        if ($script:oAuthDebugStop)
        {
            # Wait until token expires (we do this after every call when debugging OAuth)
            # Access tokens can't be revoked, but a policy can be assigned to reduce lifetime to 10 minutes: https://learn.microsoft.com/en-us/graph/api/resources/tokenlifetimepolicy?view=graph-rest-1.0
            if ($OAuthCertificate -ne $null)
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
        }
        else
        {
            $script:oAuthDebugStop = $true
        }
    }
    
    if ($OAuthCertificate -ne $null)
    {
        if ( [DateTime]::UtcNow -ge $script:oauthToken.ExpiresOn.UtcDateTime) { return }
    }
    elseif ($script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -gt [DateTime]::UtcNow.AddMinutes(1)) { return }

    # The token has expired and needs refreshing
    LogVerbose("[ApplyEWSOAuthCredentials] OAuth access token invalid, attempting to renew")
    $exchangeCredentials = GetOAuthCredentials -RenewToken
    if ($exchangeCredentials -eq $null) { return }

    if ($OAuthCertificate -ne $null)
    {
        $tokenExpire = $script:oauthToken.ExpiresOn.UtcDateTime
        if ( [DateTime]::UtcNow -ge $tokenExpire)
        {
            Log "[ApplyEWSOAuthCredentials] OAuth Token renewal failed (certificate auth)"
            exit # We no longer have access to the mailbox, so we stop here
        }
    }
    else
    {
        if ( $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in) -lt [DateTime]::UtcNow )
        { 
            Log "[ApplyEWSOAuthCredentials] OAuth Token renewal failed"
            exit # We no longer have access to the mailbox, so we stop here
        }
        $tokenExpire = $script:oauthTokenAcquireTime.AddSeconds($script:oauthToken.expires_in)
    }

    Log "[ApplyEWSOAuthCredentials] OAuth token successfully renewed; new expiry: $tokenExpire"
    if ($script:services.Count -gt 0)
    {
        foreach ($service in $script:services.Values)
        {
            $service.Credentials = $exchangeCredentials
        }
        LogVerbose "[ApplyEWSOAuthCredentials] Updated OAuth token for $($script.services.Count) ExchangeService object(s)"
    }

    if ($DebugTokenRenewal)
    {
        $global:oAuthTokenDebug = $script:oauthToken
    }
}
#>** OAUTH FUNCTIONS END **#