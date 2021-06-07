#
# Import-CalendarCSV.ps1
#
# By David Barrett, Microsoft Ltd. 2015-2021. Use at your own risk.  No warranties are given.
#
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

param(
    [string]$CSVFileName,
    [string]$EmailAddress,
    [string]$Username,
    [string]$Password,
    [string]$Domain,
    [bool]$Impersonate,
    [string]$EwsUrl,
    [string]$EWSManagedApiPath
)

Function ShowParams()
{
	Write-Host "Import-CalendarCSV -CSVFileName <string> -EmailAddress <string>"
	Write-Host "                   [-Username <string> -Password <string> [-Domain <string>]]"
	Write-Host "                   [-Impersonate <bool>]"
	Write-Host "                   [-EwsUrl <string>]"
	Write-Host "                   [-EWSManagedApiPath <string>]"
	Write-Host "";
	Write-Host "Required:"
	Write-Host " -CSVFileName : Filename of the CSV file to import appointments for this user from."
	Write-Host " -EmailAddress : Mailbox SMTP email address"
	Write-Host ""
	Write-Host "Optional:"
	Write-Host " -Username : Username for the account being used to connect to EWS (if not specified, current user is assumed)"
	Write-Host " -Password : Password for the specified user (required if username specified)"
	Write-Host " -Domain : If specified, used for authentication (not required even if username specified)"
	Write-Host " -Impersonate : Set to $true to use impersonation."
	Write-Host " -EwsUrl : Forces a particular EWS URl (otherwise autodiscover is used, which is recommended)"
	Write-Host " -EWSManagedApiDLLFilePath : Full and path to the DLL for EWS Managed API (if not specified, default path for v1.1 is used)"
	Write-Host ""
}

$RequiredFields=@{
	"Subject" = "Subject";
	"StartDate" = "Start Date";
	"StartTime" = "Start Time";
	"EndDate" = "End Date";
	"EndTime" = "End Time"
}
 
# Check email address
 if (!$EmailAddress)
 {
	ShowParams;
    throw "Required parameter EmailAddress missing"
 }
 
# CSV File Checks
if (!$CSVFileName)
{
	ShowParams
	throw "Required parameter CSVFileName missing"
}
if (!(Get-Item -Path $CSVFileName -ErrorAction SilentlyContinue))
{
	throw "Unable to open file: $CSVFileName"
}
 
# Import CSV File
try
{
	$CSVFile = Import-Csv -Path $CSVFileName
}
catch { }
if (!$CSVFile)
{
	Write-Host "CSV header line not found, using predefined header: Subject;StartDate;StartTime;EndDate;EndTime"
	$CSVFile = Import-Csv -Path $CSVFileName -header Subject,StartDate,StartTime,EndDate,EndTime
}

# Check file has required fields
foreach ($Key in $RequiredFields.Keys)
{
	if (!$CSVFile[0].$Key)
	{
		# Missing required field
		throw "Import file is missing required field: $Key"
	}
}
 
# Check EWS Managed API available
 if (!$EWSManagedApiPath)
 {
     $EWSManagedApiPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
 }
 if (!(Get-Item -Path $EWSManagedApiPath -ErrorAction SilentlyContinue))
 {
     throw "EWS Managed API could not be found at $($EWSManagedApiPath).";
 }
 
# Load EWS Managed API
 [void][Reflection.Assembly]::LoadFile($EWSManagedApiPath);
 
# Create Service Object.  We only need Exchange 2007 schema for creating calendar items (this will work with Exchange>=12)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

# Set credentials if specified, or use logged on user.
 if ($Username -and $Password)
 {
     if ($Domain)
     {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
     } else {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
     }
     
} else {
     $service.UseDefaultCredentials = $true
 }
 

# Set EWS URL if specified, or use autodiscover if no URL specified.
if ($EwsUrl)
{
	$service.URL = New-Object Uri($EwsUrl)
}
else
{
	Write-Host "Performing autodiscover for $EmailAddress"
	$service.AutodiscoverUrl($EmailAddress)
}
 
# Bind to the calendar folder
 
if ($Impersonate)
{
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
}

$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
if (!$CalendarFolder)
{
    Write-Host "Failed to locate calendar folder" -ForegroundColor Red
    exit
}

# Parse the CSV file and add the appointments
foreach ($CalendarItem in $CSVFile)
{ 
	# Create the appointment and set the fields
	$NoError=$true
	try
	{
		$Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
		$Appointment.Subject=$CalendarItem."Subject"
		$StartDate=[DateTime]($CalendarItem."StartDate" + " " + $CalendarItem."StartTime")
		$Appointment.Start=$StartDate
		$EndDate=[DateTime]($CalendarItem."EndDate" + " " + $CalendarItem."EndTime")
		$Appointment.End=$EndDate
	}
	catch
	{
		# If we fail to set any of the required fields, we will not write the appointment
		$NoError=$false
	}
	
	# Check for any other fields
	foreach ($Field in ($CalendarItem | Get-Member -MemberType Properties))
	{
		if (!($RequiredFields.Keys -contains $Field.Name))
		{
			# This is a custom (optional) field, so try to map it
			try
			{
				$Appointment.$($Field.Name)=$CalendarItem.$($Field.Name)
			}
			catch
			{
				# Failed to write this field
				Write-Host "Failed to set custom field $($Field.Name)" -ForegroundColor yellow
			}
		}
	}

	if ($NoError)
	{
		# Save the appointment
		$Appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
		Write-Host "Created $($CalendarItem."Subject")" -ForegroundColor green
	}
	else
	{
		# Failed to set a required field
		Write-Host "Failed to create appointment: $($CalendarItem."Subject")" -ForegroundColor red
	}
}