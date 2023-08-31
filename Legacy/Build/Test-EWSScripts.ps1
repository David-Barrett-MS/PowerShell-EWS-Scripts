#
# Test-EWSScripts.ps1
#
# By David Barrett, Microsoft Ltd. 2023. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.


# Folder locations
$scriptFolder = "C:\Tools\PowerShell-EWS-Scripts\PowerShell-EWS-Scripts\Legacy"
$traceFolder = "c:\Temp\Traces"
$logFolder = "c:\Temp\Logs"

# Tenant
$tenantId = "fc69f6a8-90cd-4047-977d-0c768925b8ec"

# App permissions client info
$clientIdAppPermissions = "f61d7821-7aaf-4d24-b34f-ca50528bcc7b" # App Id for app granted full_access_as_app
# Secret key and/or certificate needed (depending which tests enabled). Set prior to calling script.  $secretKey for secret key, and $certificate for certificate
# e.g. $certificate = Get-Item Cert:\CurrentUser\My\50B510B4AE120D9B0EE3F059B6DD494469CD6D3B

# Delegated permissions client info
$clientIdDelegatePermissions = ""#"42eb458d-96d4-4a5b-9d0c-2467e1cf2e59" # App Id for app granted EWS.AccessAsUser.All

# Mailboxes
$Mailbox = "dave@demonmaths.co.uk" # Primary mailbox
$DelegatedMailbox = "1@demonmaths.co.uk" # Mailbox to which Primary mailbox has FullAccess permission
$InaccessibleMailbox = "100@demonmaths.co.uk" # Mailbox to which Primary mailbox has no permissions

# Store current path and then change path to script folder
$currentPath = (Get-Location).Path
cd $scriptFolder

$runAppPermissionTests = $true
$runDelegatePermissionTests = ![String]::IsNullOrEmpty($clientIdDelegatePermissions)
$skipOAuthDebug = $true

# Use Azure Key Vault to store sensitive info (secret key) - not yet implemented
function CreateAzureKeyVault()
{
    Connect-AzAccount -TenantId $tenantId
    New-AzResourceGroup -Name "ScriptData" -Location
    New-AzKeyVault -Name "Test-EWSScripts" -ResourceGroupName "ScriptData"
}

# Check that we have valid parameters for app permission
function AppPermissionsCheck()
{
    if ([String]::IsNullOrEmpty($secretKey) -and $certificate -eq $null)
    {
        Write-Host "Application permission tests will not be run as neither `$secretKey nor `$certificate are set" -ForegroundColor Yellow
        $runAppPermissionTests = $false
    }
}



# Test RecoverDeletedItems with delegated permissions to delegated mailbox
function TestRecoverDeletedItems1()
{
    $global:testDescriptions.Add("TestRecoverDeletedItems1", "RecoverDeletedItems.ps1: attempt to access delegated archive mailbox using delegated permissions and show which items would be restored from ArchiveRecoverableItemsDeletions.")
    if (!$runDelegatePermissionTests) { return "Skipped as delegate configuration incomplete/disabled" }

    $Error.Clear()
    trap {}
    .\RecoverDeletedItems.ps1 -Mailbox $DelegatedMailbox -Archive -RestoreFromFolder "WellKnownFolderName.ArchiveRecoverableItemsDeletions" -Office365 -OAuthTenantId $tenantId -OAuthClientId $clientIdDelegatePermissions -WhatIf
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $DelegatedMailbox"
    }
    return "Succeeded, $DelegatedMailbox accessible"
}


# Test RecoverDeletedItems with delegated permissions to inaccessible mailbox
function TestRecoverDeletedItems2()
{
    $global:testDescriptions.Add("TestRecoverDeletedItems2", "RecoverDeletedItems.ps1: attempt to access other (inaccessible) archive mailbox using delegated permissions and show which items would be restored from ArchiveRecoverableItemsDeletions.")
    if (!$runDelegatePermissionTests) { return "Skipped as delegate configuration incomplete/disabled" }

    $Error.Clear()
    trap {}
    .\RecoverDeletedItems.ps1 -Mailbox $InaccessibleMailbox -Archive -RestoreFromFolder "WellKnownFolderName.ArchiveRecoverableItemsDeletions" -Office365 -OAuthTenantId $tenantId -OAuthClientId $clientIdDelegatePermissions -WhatIf
    if ($Error.Count -eq 1 -and $Error[0].Exception.Message.Contains("The specified folder could not be found in the store."))
    {
        return "Succeeded, $InaccessibleMailbox not accessible"
    }
    else
    {
        if ($Error.Count -eq 0)
        {
            return "Failed, $InaccessibleMailbox was accessible (expected to be inaccessible)"
        }
        else
        {
            return "Failed, unexpected error when accessing $InaccessibleMailbox"
        }
    }
}


# Test OAuth token renewal with delegated permissions to delegated mailbox
function TestRecoverDeletedItems3()
{
    $global:testDescriptions.Add("TestRecoverDeletedItems3", "RecoverDeletedItems.ps1: access delegated archive mailbox using delegated permissions and show which items would be restored from ArchiveRecoverableItemsDeletions.")
    if ($skipOAuthDebug) { return "Skipped as OAuth debugging disabled" }
    if (!$runDelegatePermissionTests) { return "Skipped as delegate configuration incomplete/disabled" }

    $Error.Clear()
    trap {}
    .\RecoverDeletedItems.ps1 -Mailbox $DelegatedMailbox -Archive -RestoreFromFolder "WellKnownFolderName.ArchiveRecoverableItemsDeletions" -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdDelegatePermissions -WhatIf -OAuthDebug -DebugTokenRenewal 1
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $DelegatedMailbox"
    }
    return "Succeeded, $DelegatedMailbox accessible (token renewal succeeded)"
}


# Test UpdateFolderItems with delegated permissions to primary mailbox
function TestUpdateFolderItems1()
{
    $global:testDescriptions.Add("TestUpdateFolderItems1", "Update-FolderItems.ps1: access primary mailbox using delegated permissions and set isRead for first 5 items in Inbox to true.")
    if (!$runDelegatePermissionTests) { return "Skipped as delegate configuration incomplete/disabled" }

    $Error.Clear()
    trap {}
    .\Update-FolderItems.ps1 -Mailbox $Mailbox -FolderPath "WellKnownFolderName.Inbox" -MarkAsRead -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdDelegatePermissions -MaximumNumberOfItemsToProcess 5
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $Mailbox"
    }
    return "Succeeded, $Mailbox accessible"
}


# Test UpdateFolderItems with application permissions to primary mailbox
function TestUpdateFolderItems2()
{
    $global:testDescriptions.Add("TestUpdateFolderItems2", "Update-FolderItems.ps1: access primary mailbox using application permissions and set isRead for first 5 items in Inbox to true.")
    if (!$runAppPermissionTests) { return "Skipped as app configuration incomplete" }

    $Error.Clear()
    trap {}
    .\Update-FolderItems.ps1 -Mailbox $Mailbox -FolderPath "WellKnownFolderName.Inbox" -MarkAsRead -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdAppPermissions -OAuthSecretKey $secretKey -MaximumNumberOfItemsToProcess 5 -StopAfterMaximumNumberOfItemsProcessed
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $Mailbox"
    }
    return "Succeeded, $Mailbox accessible"
}


# Test UpdateFolderItems with application permissions to primary mailbox forcing testing of token renewal
function TestUpdateFolderItems3()
{
    $global:testDescriptions.Add("TestUpdateFolderItems3", "Update-FolderItems.ps1: debug OAuth token renewal while accessing primary mailbox using application permissions and set isRead for first 5 items in Inbox to true.")
    if ($skipOAuthDebug) { return "Skipped as OAuth debugging disabled" }
    if (!$runAppPermissionTests) { return "Skipped as app configuration incomplete" }

    $Error.Clear()
    trap {}
    .\Update-FolderItems.ps1 -Mailbox $Mailbox -FolderPath "WellKnownFolderName.Inbox" -MarkAsRead -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdAppPermissions -OAuthSecretKey $secretKey -MaximumNumberOfItemsToProcess 5 -StopAfterMaximumNumberOfItemsProcessed -OAuthDebug -DebugTokenRenewal 1
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $Mailbox"
    }
    return "Succeeded, $Mailbox accessible"
}


# Test UpdateFolderItems with application permissions to primary mailbox
function TestSearchMailboxForMessageClass1()
{
    $global:testDescriptions.Add("TestSearchMailboxForMessageClass1", "Search-MailboxForMessageClass.ps1: access primary mailbox using application permissions and search for all IPM.Note items.")
    if (!$runAppPermissionTests) { return "Skipped as app configuration incomplete" }

    $Error.Clear()
    trap {}
    $matches = .\Search-MailboxForMessageClass.ps1 -Mailbox $Mailbox -MessageClass "IPM.Note" -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdAppPermissions -OAuthSecretKey $secretKey
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $Mailbox"
    }
    if ($matches.Count -eq 0)
    {
        return "Failed, access to $Mailbox succeeded, but no IPM.Note items found"
    }
    return "Succeeded, $Mailbox accessible and $($matches.Count) IPM.Note item(s) found"
}


# Test MergeMailboxFolder with application permissions to primary mailbox
function TestMergeMailboxFolder1()
{
    $global:testDescriptions.Add("TestMergeMailboxFolder1", "Merge-MailboxItems.ps1: access primary mailbox using application permissions and show what would be copied from Inbox to InboxCopy folder.")
    if (!$runAppPermissionTests) { return "Skipped as app configuration incomplete" }

    $Error.Clear()
    trap {}
    $mmresult = .\Merge-MailboxFolder.ps1 -SourceMailbox $Mailbox -MergeFolderList @{"InboxCopy" = "WellKnownFolderName.Inbox"} -WhatIf -ReturnTotalItemsAffected -Office365 -OAuth -OAuthTenantId $tenantId -OAuthClientId $clientIdAppPermissions -OAuthSecretKey $secretKey
    if ($Error.Count -gt 0)
    {
        return "Failed, error when accessing $Mailbox"
    }
    if ($mmresult -gt 0)
    {
        return "Succeeded, $Mailbox accessible and $mmresult items found to copy"
    }
    return "Check mailbox contents - no error reported, but no items found to copy (is Inbox empty?)"
}

# Run tests and collate results
AppPermissionsCheck
$global:testDescriptions = @{}
$results = @{}
$results.Add("TestRecoverDeletedItems1", "$(TestRecoverDeletedItems1)")
$results.Add("TestRecoverDeletedItems2", "$(TestRecoverDeletedItems2)")
$results.Add("TestRecoverDeletedItems3", "$(TestRecoverDeletedItems3)")
$results.Add("TestUpdateFolderItems1", "$(TestUpdateFolderItems1)")
$results.Add("TestUpdateFolderItems2", "$(TestUpdateFolderItems2)")
$results.Add("TestUpdateFolderItems3", "$(TestUpdateFolderItems3)")
$results.Add("TestSearchMailboxForMessageClass1", "$(TestSearchMailboxForMessageClass1)")
$results.Add("TestMergeMailboxFolder1", "$(TestMergeMailboxFolder1)")
$global:testResults = $results


# Output results
Write-Host
Write-Host "Test Results (available in `$testResults)"
Write-Host

foreach ($testName in $results.Keys)
{
    Write-Host "$($testName): " -NoNewline
    if ($results[$testName].StartsWith("Succeeded"))
    {
        Write-Host "$($results[$testName])" -ForegroundColor Green
    }
    elseif ($results[$testName].StartsWith("Failed"))
    {
        Write-Host "$($results[$testName])" -ForegroundColor Red        
    }
    else
    {
        Write-Host "$($results[$testName])" -ForegroundColor Yellow
    }
}

Write-Host
Write-Host "Test descriptions available in $testDescriptions"


# Restore original path
cd $currentPath