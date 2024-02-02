#
# Logging.ps1
#
# By David Barrett, Microsoft Ltd. 2016-2023. Use at your own risk.  No warranties are given.
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

#>** LOGGING PARAMETERS START **#
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified.")]	
    [string]$LogFile = "",

    [Parameter(Mandatory=$False,HelpMessage="Enable verbose log file.  Verbose logging is written to the log whether -Verbose is enabled or not.")]	
    [switch]$VerboseLogFile,

    [Parameter(Mandatory=$False,HelpMessage="Enable debug log file.  Debug logging is written to the log whether -Debug is enabled or not.")]	
    [switch]$DebugLogFile,

    [Parameter(Mandatory=$False,HelpMessage="If selected, an optimised log file creator is used that should be signficantly faster (but may leave file lock applied if script is cancelled).")]
    [switch]$FastFileLogging,
#>** LOGGING PARAMETERS END **#

    [Parameter(Mandatory=$False,HelpMessage="This is a dummy parameter to avoid syntax error.")]
    $Dummy = $null
)



#>** LOGGING FUNCTIONS START **#
$scriptStartTime = [DateTime]::Now

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}

Function UpdateDetailsWithCallingMethod([string]$Details)
{
    # Update the log message with details of the function that logged it
    $timeInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())"
    $callingFunction = (Get-PSCallStack)[2].Command # The function we are interested in will always be frame 2 on the stack
    if (![String]::IsNullOrEmpty($callingFunction))
    {
        return "$timeInfo [$callingFunction] $Details"
    }
    return "$timeInfo $Details"
}

Function LogToFile([string]$logInfo)
{
    if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    
    if ($FastFileLogging)
    {
        # Writing the log file using a FileStream (that we keep open) is significantly faster than using out-file (which opens, writes, then closes the file each time it is called)
        $fastFileLogError = $Error[0]
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            Write-Verbose "Opening/creating log file: $LogFile"
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
                $script:logFileStream = $null
            }
        }
        if ($script:logFileStream)
        {
            if (!$script:logFileStreamWriter)
            {
                $script:logFileStreamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            }
            $script:logFileStreamWriter.WriteLine($logInfo)
            $script:logFileStreamWriter.Flush()
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
            }
            else
            {
                return
            }
        }
    }

	$logInfo | Out-File $LogFile -Append
}

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Host $Details -ForegroundColor $Colour
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" Green

Function LogVerbose([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Verbose $Details
    if ( !$VerboseLogFile -and !$DebugLogFile -and ($VerbosePreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Debug $Details
    if (!$DebugLogFile -and ($DebugPreference -eq "SilentlyContinue") ) { return }
    LogToFile $Details
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
        Log "ERROR ($Context): $($Error[0])" Red
    }
    else
    {
        $log = UpdateDetailsWithCallingMethod("ERROR: $($Error[0])")
        Log $log Red
    }
    return $true
}

Function ReportError($Context)
{
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}
#>** LOGGING FUNCTIONS END **#