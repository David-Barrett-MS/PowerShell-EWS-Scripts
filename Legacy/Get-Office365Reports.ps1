#
# Get-Office365Reports.ps1
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
Retrieve Graph reports.

.DESCRIPTION
This script demonstrates how to retrieve multiple reports from Graph and deal with any throttling.

.EXAMPLE
.\Get-Office365Reports.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -ReportSavePath "c:\Reports"

This will download getOffice365ActiveUserDetail, getOffice365ActiveUserCounts, getEmailActivityUserDetail, getOneDriveActivityUserDetail for the default periods (7, 30, 90, 180 days). Reports are saved to the specified folder.

.EXAMPLE
.\Get-Office365Reports.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -ReportSavePath "c:\Reports" -RequestedReports "getOffice365ActiveUserDetail"

This will download getOffice365ActiveUserDetail report for the default periods (7, 30, 90, 180 days). Reports are saved to the specified folder.

.EXAMPLE
.\Get-Office365Reports.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -ReportSavePath "c:\Reports" -RequestedReports "getOffice365ActiveUserDetail" -RequestedPeriods "D7"

This will download getOffice365ActiveUserDetail report for the last 7 days. Report is saved to the specified folder.

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

	[Parameter(Mandatory=$True,HelpMessage="Report save path (reported are prepended by the current date)")]
	[ValidateNotNullOrEmpty()]
	[string]$ReportSavePath,

	[Parameter(Mandatory=$False,HelpMessage="The report (or list of reports) to retrieve.  Defaults to getOffice365ActiveUserDetail, getOffice365ActiveUserCounts, getEmailActivityUserDetail, getOneDriveActivityUserDetail.")]
    [ValidateNotNullOrEmpty()]
	$RequestedReports = @( "getOffice365ActiveUserDetail", "getOffice365ActiveUserCounts", "getEmailActivityUserDetail", "getOneDriveActivityUserDetail"),

	[Parameter(Mandatory=$False,HelpMessage="The report period(s) to retrieve.  Defaults to all (D7, D30, D90, D180).")]
    [ValidateNotNullOrEmpty()]
	$RequestedPeriods = @( "D7", "D30", "D90", "D180" )
)


# Acquire token
$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
try
{
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
}
catch
{
    exit # Failed to obtain a token
}
$token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

# Retrieve each of the requested reports
foreach ($report in $RequestedReports)
{
    foreach ($period in $RequestedPeriods)
    {
        $reportUri = "https://graph.microsoft.com/v1.0/reports/$report(period='$period')"
        $results = $null
        while ($results -eq $null)
        {
            try
            {
                Write-Host "GET $reportUri" -ForegroundColor Gray
                $results = Invoke-RestMethod -Method Get -Uri $reportUri -Headers $token
                if ($results -eq $null)
                {
                    # If there is no response, but no error, this is unexpected and we don't want to retry
                    $results = "Response was empty"
                }
            }
            catch
            {
                # We check for throttling - if we are throttled, we simply sleep for a while and try again
                if ($Error[0].ErrorDetails.Message.ToString().Contains("Please retry later"))
                {
                    Write-Host "Throttled. Waiting for thirty seconds before continuing." -ForegroundColor Yellow
                    Start-Sleep -Seconds 30
                    $results = $null
                }
            }
        }
        $results | ConvertFrom-Csv | Export-Csv "$ReportSavePath$([DateTime]::Today.ToString("yyyyMMdd"))$report$period.CSV" -NoTypeInformation
    }
}
