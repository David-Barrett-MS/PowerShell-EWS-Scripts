#
# Get-Office365ServiceStatus.ps1
#
# By David Barrett, Microsoft Ltd. 2021. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

<#
.SYNOPSIS
Retrieve data from the Service Communications API.

.DESCRIPTION
This script demonstrates how to retrieve data from the Service Communications API (which provides current and historical status about Office 365 services).

.EXAMPLE
.\Get-Office365ServiceStatus.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -CurrentStatus

This will display the current reported status for each workload

.EXAMPLE
.\Get-Office365ServiceStatus.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -Message

This will display the current list of messages for each service and workload.

#>


param (
	[Parameter(Mandatory=$True,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$True,HelpMessage="Application secret key (obtained when registering the application in Azure AD")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

	[Parameter(Mandatory=$True,HelpMessage="Tenant Id")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId,

	[Parameter(Mandatory=$True,HelpMessage="Tenant domain")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantDomain,

	[Parameter(Mandatory=$False,HelpMessage="Retrieve list of subscribed services")]
	[ValidateNotNullOrEmpty()]
	[switch]$Services,

	[Parameter(Mandatory=$False,HelpMessage="Retrieve the status of the service from the previous 24 hours")]
	[ValidateNotNullOrEmpty()]
	[switch]$CurrentStatus,

	[Parameter(Mandatory=$False,HelpMessage="Retrieve the status of the service from the previous 24 hours")]
	[ValidateNotNullOrEmpty()]
	[switch]$HistoricalStatus,

	[Parameter(Mandatory=$False,HelpMessage="Retrieve the status of the service from the previous 24 hours")]
	[ValidateNotNullOrEmpty()]
	[switch]$Messages,

	[Parameter(Mandatory=$False,HelpMessage="Report save path (reported are prepended by the current date)")]
	[ValidateNotNullOrEmpty()]
	[string]$ReportSavePath
)


# Acquire token
$body = @{grant_type="client_credentials";resource="https://manage.office.com";client_id=$AppId;client_secret=$AppSecretKey}
#$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
try
{
    #$oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.windows.net/$TenantDomain/oauth2/token?api-version=1.0 -Body $body

}
catch
{
    Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
    exit # Failed to obtain a token
}
$token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
Write-Verbose "$($oauth.token_type) $($oauth.access_token)"


if ($Services)
{
    # Get Services
    $getServices = Invoke-WebRequest -Method 'GET' -Uri "https://manage.office.com/api/v1.0/$TenantId/ServiceComms/Services" -Headers $token
    $global:services = $getServices.Content

    if (![String]::IsNullOrEmpty($ReportSavePath))
    {
        $getServices.Content | Out-File "$ReportSavePath$([DateTime]::Today.ToString("yyyyMMdd")) Services.CSV"
    }
}

if ($CurrentStatus)
{
    # Get current Status - we just report the feature status, not the individual workloads
    $getCurrentStatus = Invoke-WebRequest -Method Get -Uri "https://manage.office.com/api/v1.0/$TenantId/ServiceComms/CurrentStatus" -Headers $token
    $statusJson = ConvertFrom-Json $getCurrentStatus.Content
    $global:currentStatus = $statusJson

    Write-Host "Current Status"
    Write-Host "--------------"
    Write-Host ""

    foreach ($FeatureStatus in $statusJson.value)
    {
        Write-Host "$($FeatureStatus.Id): $($FeatureStatus.Status)"
    }

    if (![String]::IsNullOrEmpty($ReportSavePath))
    {
        $getCurrentStatus.Content | Out-File "$ReportSavePath$([DateTime]::Now.ToString("yyyyMMddhhmmss")) CurrentStatus.CSV"
    }
    Write-Host ""
}

if ($HistoricalStatus)
{
    # Get historical Status
    $getHistoricalStatus = Invoke-WebRequest -Method Get -Uri "https://manage.office.com/api/v1.0/$TenantId/ServiceComms/HistoricalStatus" -Headers $token
    $statusJson = ConvertFrom-Json $getHistoricalStatus.Content
    $global:historicalStatus = $statusJson

    Write-Host "Historical Status"
    Write-Host "-----------------"
    Write-Host ""

    foreach ($FeatureStatus in $statusJson.value)
    {
        Write-Host "$($FeatureStatus.Id): $($FeatureStatus.Status)"
    }

    if (![String]::IsNullOrEmpty($ReportSavePath))
    {
        $getHistoricalStatus.Content | Out-File "$ReportSavePath$([DateTime]::Now.ToString("yyyyMMddhhmmss")) HistoricalStatus.CSV"
    }
    Write-Host ""
}

if ($Messages)
{
    # Get messages
    $getMessages = Invoke-WebRequest -Method Get -Uri "https://manage.office.com/api/v1.0/$TenantId/ServiceComms/Messages" -Headers $token
    $statusJson = ConvertFrom-Json $getMessages.Content

    Write-Host "Messages"
    Write-Host "--------"
    Write-Host ""

    foreach ($messageGroup in $statusJson.value)
    {
        Write-Host "$($messageGroup.Workload) $($messageGroup.Id): $($messageGroup.Status) - $($messageGroup.Messages.Count) message(s)"
    }

    if (![String]::IsNullOrEmpty($ReportSavePath))
    {
        $getMessages.Content | Out-File "$ReportSavePath$([DateTime]::Now.ToString("yyyyMMddhhmmss")) Messages.CSV"
    }
}