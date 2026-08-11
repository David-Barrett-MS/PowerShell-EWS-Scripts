#
# Test-EWSAppAccess.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

<#
.SYNOPSIS
Tests accessing mailbox using EWS with application permissions.

.DESCRIPTION
Script to test accessing mailbox using EWS with application permissions.

.EXAMPLE

Delegated permissions:
.\Test-EWSAppAccess.ps1 -AppId $clientId -TenantId $tenantId

Application permissions:
.\Test-EWSAppAccess.ps1 -Mailbox $Mailbox -AppId $clientId -TenantId $tenantId -SecretKey $secretKey
.\Test-EWSAppAccess.ps1 -Mailbox $Mailbox -AppId $clientId -TenantId $tenantId -Certificate $certPath
#>

param (
    [Parameter(Mandatory=$False,HelpMessage="The client Id that this script will identify as.  Must be registered in EntraId.")]
    [string]$AppId = "",

    [Parameter(Mandatory=$False,HelpMessage="The tenant Id (application must be registered in the same tenant being accessed).")]
    [string]$TenantId = "",

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the EntraId registered application (delegated flow).")]
    [string]$RedirectUri = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.")]
    [string]$SecretKey = "",

    [Parameter(Mandatory=$False,HelpMessage="If using application permissions, specify the secret key OR certificate.  The certificate may be an X509Certificate2 object, certificate store path, or PFX path.")]
    $Certificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="Password for an encrypted PFX file specified with Certificate.")]
    [System.Security.SecureString]$CertificatePassword = $null,

    [Parameter(Mandatory=$False,HelpMessage="The mailbox to access.  Required when application permissions are used (the mailbox is then accessed using impersonation).")]
    [string]$Mailbox = ""
)


function GetTokenWithCertificate
{
    $certificate = $Certificate
    if ($certificate -is [string])
    {
        if (!(Test-Path -LiteralPath $certificate))
        {
            throw "Certificate path not found: $certificate"
        }

        $certificateItem = Get-Item -LiteralPath $certificate -ErrorAction Stop
        if ($certificateItem -is [System.Security.Cryptography.X509Certificates.X509Certificate2])
        {
            $certificate = $certificateItem
        }
        elseif ($null -ne $CertificatePassword)
        {
            $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certificate, $CertificatePassword)
        }
        else
        {
            $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certificate)
        }
    }   

    if ($certificate -isnot [System.Security.Cryptography.X509Certificates.X509Certificate2])
    {
        throw "Certificate must be an X509Certificate2 object, certificate store path, or PFX path"
    }
    if (!$certificate.HasPrivateKey)
    {
        throw "The certificate does not contain a private key"
    }

    function ConvertToBase64Url([byte[]]$bytes)
    {
        return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [DateTimeOffset]::UtcNow
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try
    {
        $thumbprint = ConvertToBase64Url ($sha256.ComputeHash($certificate.RawData))
    }
    finally
    {
        $sha256.Dispose()
    }

    $header = @{ alg = "RS256"; typ = "JWT"; 'x5t#S256' = $thumbprint } | ConvertTo-Json -Compress
    $payload = @{ aud = $tokenEndpoint; iss = $AppId; sub = $AppId; jti = [Guid]::NewGuid().ToString(); nbf = $now.AddMinutes(-5).ToUnixTimeSeconds(); exp = $now.AddMinutes(10).ToUnixTimeSeconds() } | ConvertTo-Json -Compress
    $unsignedToken = "$(ConvertToBase64Url ([Text.Encoding]::UTF8.GetBytes($header))).$(ConvertToBase64Url ([Text.Encoding]::UTF8.GetBytes($payload)))"

    $rsa = $null
    try
    {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
        if ($null -eq $rsa)
        {
            throw "Unable to access the certificate private key"
        }
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($unsignedToken), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $clientAssertion = "$unsignedToken.$(ConvertToBase64Url $signature)"

        $body = @{
            grant_type = "client_credentials"
            client_id = $AppId
            scope = "https://outlook.office365.com/.default"
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion = $clientAssertion
        }
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
    }
    catch
    {
        Write-Error "Failed to obtain OAuth token: $_"
        exit
    }

    if ($null -ne $rsa) { $rsa.Dispose() }
    $script:oAuthAccessToken = $script:oAuthToken.access_token
    if ($null -ne $script:oAuthAccessToken)
    {
        return
    }

    # If we get here, we don't have a token so can't continue
    if ($null -ne $execCall.Exception)
    {
        Write-Error "Failed to obtain OAuth token: $($execCall.Exception.Message)"
    }
    else {
        Write-Error "Failed to obtain OAuth token (no error thrown)."
    }
    exit
}

function GetTokenViaCode
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?client_id=$AppId&response_type=code&redirect_uri=$RedirectUri&response_mode=query&prompt=select_account&scope=openid%20profile%20email%20https://outlook.office365.com/.default"
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
        Write-Host "Auth code acquired, attempting to obtain access token" -ForegroundColor Green
    }
    else
    {
        throw "Failed to obtain Auth code from clipboard"
    }

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://outlook.office365.com/.default";client_id=$AppId;code=$authcode;redirect_uri=$RedirectUri}
    try
    {
        $script:oauthToken = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
        return
    }
    catch {}

    throw "Failed to obtain OAuth token"
}

function GetTokenWithKey
{
    $Body = @{
      "grant_type"    = "client_credentials";
      "client_id"     = "$AppId";
      "scope"         = "https://outlook.office365.com/.default";
      "client_secret" = "$SecretKey"
    }

    try
    {
        $script:oAuthToken = Invoke-RestMethod -Method POST -uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
        $script:oAuthAccessToken = $script:oAuthToken.access_token
    }
    catch
    {
        Write-Error "Failed to obtain OAuth token: $_"
        exit # Failed to obtain a token
    }
}

function GetOAuthAccessToken
{
    # Obtain an OAuth access token for use in HTTP Authorization headers.

    if (![String]::IsNullOrEmpty($SecretKey))
    {
        GetTokenWithKey
    }
    elseif ($null -ne $Certificate)
    {
        GetTokenWithCertificate
    }
    else
    {
        GetTokenViaCode
    }

    return $script:oAuthAccessToken
}

function SendEWSRequest
{
    # Send a SOAP request to the EWS endpoint and return the response as XML
    param (
        [Parameter(Mandatory=$true)][string]$RequestBody
    )

    $headers = @{
        "Authorization" = "Bearer $($script:oAuthAccessToken)"
        "Content-Type" = "text/xml; charset=utf-8"
    }
    if (![String]::IsNullOrEmpty($Mailbox))
    {
        $headers["X-AnchorMailbox"] = $Mailbox
    }

    $soapEnvelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
$($script:impersonationHeader)    <t:RequestServerVersion Version="Exchange2016" />
  </soap:Header>
  <soap:Body>
$RequestBody
  </soap:Body>
</soap:Envelope>
"@

    $response = Invoke-WebRequest -Uri $script:ewsUrl -Method Post -Headers $headers -Body $soapEnvelope -UseBasicParsing
    return [xml]$response.Content
}

function GetInboxFolderId
{
    # Retrieve the Id of the Inbox folder
    $request = @"
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox" />
      </m:FolderIds>
    </m:GetFolder>
"@

    $responseXml = SendEWSRequest -RequestBody $request

    $namespaces = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
    $namespaces.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages")
    $namespaces.AddNamespace("t", "http://schemas.microsoft.com/exchange/services/2006/types")

    $responseMessage = $responseXml.SelectSingleNode("//m:GetFolderResponseMessage", $namespaces)
    if ($null -eq $responseMessage -or $responseMessage.ResponseClass -ne "Success")
    {
        throw "GetFolder failed: $($responseMessage.MessageText)"
    }

    $folderId = $responseXml.SelectSingleNode("//t:Folder/t:FolderId", $namespaces)
    if ($null -eq $folderId)
    {
        throw "GetFolder succeeded but no folder Id was returned"
    }
    return $folderId
}

# Application permissions require a mailbox to impersonate; delegated permissions access the signed-in user's mailbox
$script:appFlow = (![String]::IsNullOrEmpty($SecretKey)) -or ($null -ne $Certificate)
if ($script:appFlow -and [String]::IsNullOrEmpty($Mailbox))
{
    Write-Error "-Mailbox must be specified when using application permissions (secret key or certificate)."
    exit 1
}

$script:ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$script:impersonationHeader = ""
if ($script:appFlow)
{
    $script:impersonationHeader = @"
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$([System.Security.SecurityElement]::Escape($Mailbox))</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>

"@
}

$script:oAuthAccessToken = GetOAuthAccessToken
if ([String]::IsNullOrEmpty($script:oAuthAccessToken))
{
    Write-Error "Failed to obtain an OAuth access token."
    exit 1
}

$mailboxAccessed = $false
$inboxFolderId = $null

try
{
    $inboxFolderId = GetInboxFolderId
    $mailboxAccessed = $true
}
catch
{
    Write-Host "Failed to retrieve Inbox folder: $_" -ForegroundColor Red
}

if ($mailboxAccessed)
{
    Write-Host "Application $AppId successfully accessed mailbox $Mailbox" -ForegroundColor Green
    exit 0
}

Write-Host "Application $AppId failed to access mailbox $Mailbox" -ForegroundColor Red
exit 1

