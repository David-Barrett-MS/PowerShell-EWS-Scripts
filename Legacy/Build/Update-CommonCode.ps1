#
# Update-CommonCode.ps1
#
# By David Barrett, Microsoft Ltd. 2023. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.


# $ScriptFolder = "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy"
# $SharedFolder = "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Build"
# $BackupFolder = "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Backup"
# .\Update-CommonCode.ps1 -ScriptFolder $ScriptFolder -SharedCode @("EWSOAuth.ps1", "Logging.ps1") -BackupFolder $BackupFolder -SharedCodeFolder $SharedFolder
# .\Update-CommonCode.ps1 -ScriptFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy" -SharedCode @("EWSOAuth.ps1", "Logging.ps1") -BackupFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Backup" -SharedCodeFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Build"
# .\Update-CommonCode.ps1 -ScriptFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy" -SharedCode @("EWSOAuth.ps1", "Logging.ps1") -BackupFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Backup" -SharedCodeFolder "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy\Build" -StripSharedCode


param (

    [Parameter(Mandatory=$True,HelpMessage="Source folder containing scripts to be processed.")]
    $ScriptFolder,
	
    [Parameter(Mandatory=$True,HelpMessage="Shared code file(s) to be injected.")]
    $SharedCode,
	
    [Parameter(Mandatory=$False,HelpMessage="Folder where shared code files are located.")]
    $SharedCodeFolder,

    [Parameter(Mandatory=$False,HelpMessage="Folder where files will be backed up prior to update (if not specified, .ccbak file will be created)")]
    $BackupFolder,

    [Parameter(Mandatory=$False,HelpMessage="If specified, removes the shared code modules from the target PowerShell scripts.")]
    [switch]$StripSharedContent
)


$psSourceFiles = Get-ChildItem -Path $ScriptFolder -Include *.ps1 -Name

function ReplaceSharedCode()
{
    param (
        $sourceCode,
        [String]$SharedCodeTemplate
    )

    $sharedCode = Get-Content $SharedCodeTemplate
    $sharedCodeIndex = 0
    $sourceCodeIndex = 0
    $updatedCode = @()
    $moduleApplied = $false

    Write-Verbose "Checking $($sourceCode.Length) lines of code"
    Write-Verbose "Using template: $SharedCodeTemplate"

    while ($sharedCodeIndex -lt $sharedCode.Length -and $sourceCodeIndex -lt $sourceCode.Length)
    {
        while ($sharedCodeIndex -lt $sharedCode.Length -and -not $sharedCode[$sharedCodeIndex].StartsWith("#>**"))
            { $sharedCodeIndex++ }

        if ($sharedCode[$sharedCodeIndex].StartsWith("#>**"))
        {
            if ($sharedCode[$sharedCodeIndex].EndsWith("START **#"))
            {
                while ($sourceCodeIndex -lt $sourceCode.Length)
                {
                    if ($sourceCode[$sourceCodeIndex].Equals($sharedCode[$sharedCodeIndex]))
                    {
                        # Start of share code injection
                        Write-Verbose $sharedCode[$sharedCodeIndex]
                        $script:codeUpdated = $true
                        $moduleApplied = $true
                        do 
                        {
                            if (!$StripSharedContent) { $updatedCode += $sharedCode[$sharedCodeIndex++] }
                        } while ($sharedCodeIndex -lt $sharedCode.Length -and -not $sharedCode[$sharedCodeIndex].StartsWith("#>**"))
                        
                        do
                        {
                            $sourceCodeIndex++
                        } while ($sourceCodeIndex -lt $sourceCode.Length -and -not $sourceCode[$sourceCodeIndex].Equals($sharedCode[$sharedCodeIndex]))
                        Write-Verbose $sharedCode[$sharedCodeIndex]
                        $sharedCodeIndex++
                        break
                    }
                    else
                    {
                        $updatedCode += $sourceCode[$SourceCodeIndex]
                    }
                    $sourceCodeIndex++
                }
            }
        }
    }

    if (!$moduleApplied)
    {
        return $sourceCode
    }

    Write-Host "Applied code from template: $SharedCodeTemplate"
    while ($sourceCodeIndex -lt $sourceCode.Length)
    {
        $updatedCode += $sourceCode[$SourceCodeIndex++]
    }
    return $updatedCode
}


foreach ($psSourceFile in $psSourceFiles)
{
    Write-Host "Processing $psSourceFile"
    $psCode = Get-Content "$ScriptFolder\$psSourceFile"
    $updatedCode = $psCode

    $script:codeUpdated = $false
    $SharedCode | foreach {
        $sharedCodeModule = $_
        if (-not (Test-Path $sharedCodeModule))
        {
            $sharedCodeModule = "$SharedCodeFolder\$sharedCodeModule"
        }
        if (Test-Path $sharedCodeModule)
        {
            $updatedCode = ReplaceSharedCode $updatedCode $sharedCodeModule
        }
    }

    $backupFileName = "$ScriptFolder\$psSourceFile.ccbak"
    if (![String]::IsNullOrEmpty($BackupFolder))
    {
        $backupFileName = "$BackupFolder\$psSourceFile"
    }

    if ($script:codeUpdated)
    {
        if (Test-Path $backupFileName)
        {
            Remove-Item $backupFileName
        }
        Copy-Item -Path "$ScriptFolder\$psSourceFile" -Destination $backupFileName
        if (Test-Path $backupFileName)
        {
            Write-Host "Original backed up to: $backupFileName"
            $updatedCode | Out-File "$ScriptFolder\$psSourceFile" -Encoding utf8
            Write-Host "Updated $psSourceFile" -ForegroundColor Green
        }
        else
        {
            Write-Host "Failed to backup $psSourceFile - changes not written" -ForegroundColor Red
        }
    }
    else
    {
        Write-Host "No change to $psSourceFile" -ForegroundColor Yellow
    }
}