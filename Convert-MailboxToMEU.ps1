<#
.SYNOPSIS
Converts Office 365 mailboxes to Mail-Enabled Users (MEU) with external forwarding while preserving essential attributes.

.DESCRIPTION
This script converts Office 365 mailboxes to Mail-Enabled Users (MEU) and configures external forwarding.
The conversion preserves essential mailbox attributes including email addresses, display names, and organizational data.
The process maintains organizational directory presence while redirecting email to external addresses.

.PARAMETER InputCsvPath
Optional path to CSV file containing mailboxes to convert. CSV must have 'upn' and 'externalEmail' columns.
If not provided, script will prompt for individual mailbox conversion.

.PARAMETER OutputPath
Directory path where conversion reports will be saved. Defaults to current directory.

.EXAMPLE
.\Convert-MailboxToMEU.ps1 -WhatIf
Simulates MEU conversion without making changes, shows what would be processed.

.EXAMPLE
.\Convert-MailboxToMEU.ps1 -InputCsvPath "C:\temp\conversion-list.csv" -OutputPath "C:\Reports"
Converts mailboxes listed in CSV file to MEUs with external forwarding.

.EXAMPLE
.\Convert-MailboxToMEU.ps1 -InputCsvPath "C:\temp\conversion-list.csv" -Confirm:$false
Converts mailboxes without confirmation prompts for each conversion.

.NOTES
Author: Exchange Online Administration
Requires: Exchange Online Management module, on-premises Exchange PowerShell session
Prerequisites:
  - Active Exchange Online session (Connect-ExchangeOnline)
  - Active on-premises Exchange PowerShell session for hybrid operations
  - Appropriate permissions for mailbox operations in both environments
Output: CSV reports of conversion results, deleted mailbox attributes, and any errors encountered
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$InputCsvPath,
    [string]$OutputPath = (Get-Location).Path
)

# Generate output file paths
$ConversionReportFile = Join-Path $OutputPath "MEUConversionReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$DeletedMailboxAttributesFile = Join-Path $OutputPath "DeletedMailboxAttributes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$ErrorReportFile = Join-Path $OutputPath "MEUConversionErrors_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

Write-Information "Starting Mailbox to MEU Conversion Process..." -InformationAction Continue

if ($WhatIfPreference)
{
    Write-Warning "Running in SIMULATION mode - no changes will be made"
}

try
{
    # Phase 1: Validate prerequisites and connection
    Write-Information "Phase 1: Validating prerequisites and connections..." -InformationAction Continue

    # Check for existing Exchange Online session
    Write-Verbose "Checking Exchange Online connection..."
    try
    {
        $sessionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $sessionInfo -or $sessionInfo.State -ne "Connected")
        {
            Write-Error "No active Exchange Online session found. Please run Connect-ExchangeOnline first."
            Write-Information "Example: Connect-ExchangeOnline -UserPrincipalName youradmin@domain.com" -InformationAction Continue
            exit 1
        }
        else
        {
            Write-Information "Found active Exchange Online session: $($sessionInfo.Name)" -InformationAction Continue
        }
    }
    catch
    {
        Write-Error "Exchange Online module not available or no active session. Please run Connect-ExchangeOnline first."
        exit 1
    }

    # Check for on-premises Exchange session (required for hybrid operations)
    Write-Verbose "Checking on-premises Exchange connection..."
    try
    {
        $onPremSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.ComputerName -notlike "*.outlook.com" }
        if (-not $onPremSession)
        {
            Write-Error "No active on-premises Exchange session found. Please establish session to on-premises Exchange server first."
            Write-Information "Example: `$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver/PowerShell/ -Authentication Kerberos" -InformationAction Continue
            Write-Information "         Import-PSSession `$Session" -InformationAction Continue
            exit 1
        }
        else
        {
            Write-Information "Found active on-premises Exchange session: $($onPremSession.ComputerName)" -InformationAction Continue
        }
    }
    catch
    {
        Write-Error "Unable to verify on-premises Exchange connection. This is required for hybrid operations."
        exit 1
    }

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath))
    {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Information "Created output directory: $OutputPath" -InformationAction Continue
    }

    # Phase 1: Determine target mailboxes for conversion
    $conversionList = @()

    if ($InputCsvPath -and (Test-Path $InputCsvPath))
    {
        Write-Information "Processing CSV input file: $InputCsvPath" -InformationAction Continue
        $csvData = Import-Csv $InputCsvPath

        # Validate CSV structure
        $requiredColumns = @("upn", "externalEmail")
        foreach ($column in $requiredColumns)
        {
            if (-not ($csvData | Get-Member -Name $column -MemberType NoteProperty))
            {
                Write-Error "CSV file must contain a '$column' column."
                exit 1
            }
        }

        Write-Information "Found $($csvData.Count) entries in CSV file" -InformationAction Continue

        # Validate each mailbox exists
        foreach ($entry in $csvData)
        {
            try
            {
                $mailbox = Get-Mailbox -Identity $entry.upn -ErrorAction SilentlyContinue
                if ($mailbox)
                {
                    $conversionList += [PSCustomObject]@{
                        Mailbox = $mailbox
                        ExternalEmail = $entry.externalEmail
                        SourceUPN = $entry.upn
                    }
                }
                else
                {
                    Write-Warning "Could not find mailbox for UPN: $($entry.upn)"
                }
            }
            catch
            {
                Write-Warning "Error validating mailbox $($entry.upn): $($_.Exception.Message)"
            }
        }
    }
    else
    {
        Write-Warning "No CSV file specified. Interactive mode not implemented in this phase."
        Write-Information "Please provide a CSV file with 'upn' and 'externalEmail' columns." -InformationAction Continue
        exit 1
    }

    Write-Information "Validated $($conversionList.Count) mailboxes for conversion" -InformationAction Continue

    if ($conversionList.Count -eq 0)
    {
        Write-Warning "No valid mailboxes found for conversion. Exiting."
        exit 0
    }

    # Phase 2: Preserve mailbox attributes
    Write-Information "Phase 2: Analyzing and preserving mailbox attributes..." -InformationAction Continue

    $conversionResults = @()
    $deletedMailboxAttributes = @()
    $errorResults = @()

    foreach ($conversionItem in $conversionList)
    {
        $mailbox = $conversionItem.Mailbox
        $externalEmail = $conversionItem.ExternalEmail

        Write-Progress -Activity "Analyzing Mailboxes" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($conversionResults.Count / $conversionList.Count) * 100)

        try
        {
            # Gather essential attributes to preserve
            $attributesToPreserve = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                FirstName = $mailbox.FirstName
                LastName = $mailbox.LastName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                UserPrincipalName = $mailbox.UserPrincipalName
                EmailAddresses = $mailbox.EmailAddresses
                Office = $mailbox.Office
                Department = $mailbox.Department
                Title = $mailbox.Title
                Company = $mailbox.Company
                Manager = $mailbox.Manager
                Phone = $mailbox.Phone
                MobilePhone = $mailbox.MobilePhone
                Fax = $mailbox.Fax
                CustomAttribute1 = $mailbox.CustomAttribute1
                CustomAttribute2 = $mailbox.CustomAttribute2
                CustomAttribute3 = $mailbox.CustomAttribute3
                CustomAttribute4 = $mailbox.CustomAttribute4
                CustomAttribute5 = $mailbox.CustomAttribute5
                ExtensionCustomAttribute1 = $mailbox.ExtensionCustomAttribute1
                ExtensionCustomAttribute2 = $mailbox.ExtensionCustomAttribute2
                ExtensionCustomAttribute3 = $mailbox.ExtensionCustomAttribute3
                ExtensionCustomAttribute4 = $mailbox.ExtensionCustomAttribute4
                ExtensionCustomAttribute5 = $mailbox.ExtensionCustomAttribute5
                MailboxType = $mailbox.RecipientTypeDetails
                ExternalEmailAddress = $externalEmail
                ConversionDate = Get-Date
            }

            # Save complete mailbox attributes for deletion record
            $deletedMailboxRecord = [PSCustomObject]@{
                OriginalDisplayName = $mailbox.DisplayName
                OriginalUserPrincipalName = $mailbox.UserPrincipalName
                OriginalPrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                OriginalEmailAddresses = ($mailbox.EmailAddresses -join "; ")
                OriginalRecipientTypeDetails = $mailbox.RecipientTypeDetails
                OriginalFirstName = $mailbox.FirstName
                OriginalLastName = $mailbox.LastName
                OriginalOffice = $mailbox.Office
                OriginalDepartment = $mailbox.Department
                OriginalTitle = $mailbox.Title
                OriginalCompany = $mailbox.Company
                OriginalManager = $mailbox.Manager
                OriginalPhone = $mailbox.Phone
                OriginalMobilePhone = $mailbox.MobilePhone
                OriginalFax = $mailbox.Fax
                OriginalCustomAttribute1 = $mailbox.CustomAttribute1
                OriginalCustomAttribute2 = $mailbox.CustomAttribute2
                OriginalCustomAttribute3 = $mailbox.CustomAttribute3
                OriginalCustomAttribute4 = $mailbox.CustomAttribute4
                OriginalCustomAttribute5 = $mailbox.CustomAttribute5
                OriginalExtensionCustomAttribute1 = $mailbox.ExtensionCustomAttribute1
                OriginalExtensionCustomAttribute2 = $mailbox.ExtensionCustomAttribute2
                OriginalExtensionCustomAttribute3 = $mailbox.ExtensionCustomAttribute3
                OriginalExtensionCustomAttribute4 = $mailbox.ExtensionCustomAttribute4
                OriginalExtensionCustomAttribute5 = $mailbox.ExtensionCustomAttribute5
                OriginalArchiveStatus = $mailbox.ArchiveStatus
                OriginalArchiveDatabase = $mailbox.ArchiveDatabase
                OriginalDatabase = $mailbox.Database
                OriginalServerName = $mailbox.ServerName
                OriginalWhenCreated = $mailbox.WhenCreated
                OriginalWhenChanged = $mailbox.WhenChanged
                OriginalMailboxSizeGB = if ($mailbox.UseDatabaseQuotaDefaults) { "Using Database Defaults" } else { "$($mailbox.ProhibitSendQuota)" }
                ConvertedToExternalEmail = $externalEmail
                ConversionDate = Get-Date
                ConversionStatus = "Pending"
            }

            # Get current delegation permissions (for reporting)
            $delegationInfo = @{
                FullAccessDelegates = @()
                SendAsDelegates = @()
                SendOnBehalfDelegates = @()
            }

            try
            {
                $delegationInfo.FullAccessDelegates = Get-MailboxPermission $mailbox.PrimarySmtpAddress |
                    Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false } |
                    Select-Object -ExpandProperty User
            }
            catch
            {
                Write-Warning "Could not retrieve Full Access permissions for $($mailbox.DisplayName)"
            }

            try
            {
                $delegationInfo.SendAsDelegates = Get-RecipientPermission $mailbox.PrimarySmtpAddress |
                    Where-Object { $_.AccessRights -eq 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false } |
                    Select-Object -ExpandProperty Trustee
            }
            catch
            {
                Write-Warning "Could not retrieve Send As permissions for $($mailbox.DisplayName)"
            }

            if ($mailbox.GrantSendOnBehalfTo)
            {
                $delegationInfo.SendOnBehalfDelegates = $mailbox.GrantSendOnBehalfTo
            }

            # Add delegation info to deletion record
            $deletedMailboxRecord | Add-Member -NotePropertyName "OriginalFullAccessDelegates" -NotePropertyValue ($delegationInfo.FullAccessDelegates -join "; ")
            $deletedMailboxRecord | Add-Member -NotePropertyName "OriginalSendAsDelegates" -NotePropertyValue ($delegationInfo.SendAsDelegates -join "; ")
            $deletedMailboxRecord | Add-Member -NotePropertyName "OriginalSendOnBehalfDelegates" -NotePropertyValue ($delegationInfo.SendOnBehalfDelegates -join "; ")

            # Add to deletion tracking
            $deletedMailboxAttributes += $deletedMailboxRecord

            $conversionResults += [PSCustomObject]@{
                OriginalMailboxName = $attributesToPreserve.DisplayName
                OriginalUPN = $attributesToPreserve.UserPrincipalName
                OriginalMailboxType = $attributesToPreserve.MailboxType
                PrimarySmtpAddress = $attributesToPreserve.PrimarySmtpAddress
                ExternalForwardingAddress = $attributesToPreserve.ExternalEmailAddress
                PreservedAttributes = ($attributesToPreserve | Get-Member -MemberType NoteProperty | Measure-Object).Count
                FullAccessDelegatesCount = $delegationInfo.FullAccessDelegates.Count
                SendAsDelegatesCount = $delegationInfo.SendAsDelegates.Count
                SendOnBehalfDelegatesCount = $delegationInfo.SendOnBehalfDelegates.Count
                FullAccessDelegates = ($delegationInfo.FullAccessDelegates -join "; ")
                SendAsDelegates = ($delegationInfo.SendAsDelegates -join "; ")
                SendOnBehalfDelegates = ($delegationInfo.SendOnBehalfDelegates -join "; ")
                ConversionStatus = if ($WhatIfPreference) { "Simulated" } else { "Ready for Conversion" }
                ConversionDate = $attributesToPreserve.ConversionDate
                Notes = "Attributes analyzed and preserved"
            }
        }
        catch
        {
            $errorResults += [PSCustomObject]@{
                MailboxName = $mailbox.DisplayName
                MailboxUPN = $mailbox.UserPrincipalName
                ErrorType = "AttributePreservation"
                ErrorMessage = $_.Exception.Message
                ErrorDate = Get-Date
            }
            Write-Warning "Error preserving attributes for $($mailbox.DisplayName): $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "Analyzing Mailboxes" -Completed

    # Phase 3: Configure external forwarding (simulation or actual)
    Write-Information "Phase 3: Configuring external forwarding..." -InformationAction Continue

    if (-not $WhatIfPreference)
    {
        if ($PSCmdlet.ShouldProcess("$($conversionResults.Count) mailboxes", "Convert to MEU with external forwarding"))
        {
            Write-Information "Proceeding with MEU conversion..." -InformationAction Continue
        }
        else
        {
            Write-Information "Conversion cancelled by user." -InformationAction Continue
            exit 0
        }

        Write-Information "Proceeding with actual MEU conversion..." -InformationAction Continue

        foreach ($result in $conversionResults)
        {
            Write-Progress -Activity "Converting to MEU" -Status "Converting $($result.OriginalMailboxName)" -PercentComplete (($conversionResults.IndexOf($result) / $conversionResults.Count) * 100)

            try
            {
                Write-Information "Converting $($result.OriginalMailboxName) to MEU..." -InformationAction Continue

                # Step 1: Disable the remote mailbox (on-premises Exchange)
                Write-Verbose "Step 1: Disabling remote mailbox $($result.OriginalUPN) on-premises..."
                Disable-RemoteMailbox -Identity $result.OriginalUPN -Confirm:$false

                # Step 2: Enable as Mail User in Exchange Online
                Write-Verbose "Step 2: Enabling mail user $($result.OriginalUPN) in Exchange Online..."
                Enable-MailUser -Identity $result.OriginalUPN -ExternalEmailAddress $result.ExternalForwardingAddress

                # Step 3: Configure mail user with preserved attributes
                Write-Verbose "Step 3: Configuring mail user attributes for $($result.OriginalUPN)..."

                # Find the corresponding preserved attributes
                $preservedAttributes = $deletedMailboxAttributes | Where-Object { $_.OriginalUserPrincipalName -eq $result.OriginalUPN }
                if ($preservedAttributes)
                {
                    $setMailUserParams = @{
                        Identity = $result.OriginalUPN
                    }

                    # Add non-null attributes to the parameter set
                    if ($preservedAttributes.OriginalDisplayName) { $setMailUserParams.DisplayName = $preservedAttributes.OriginalDisplayName }
                    if ($preservedAttributes.OriginalFirstName) { $setMailUserParams.FirstName = $preservedAttributes.OriginalFirstName }
                    if ($preservedAttributes.OriginalLastName) { $setMailUserParams.LastName = $preservedAttributes.OriginalLastName }
                    if ($preservedAttributes.OriginalOffice) { $setMailUserParams.Office = $preservedAttributes.OriginalOffice }
                    if ($preservedAttributes.OriginalDepartment) { $setMailUserParams.Department = $preservedAttributes.OriginalDepartment }
                    if ($preservedAttributes.OriginalTitle) { $setMailUserParams.Title = $preservedAttributes.OriginalTitle }
                    if ($preservedAttributes.OriginalCompany) { $setMailUserParams.Company = $preservedAttributes.OriginalCompany }
                    if ($preservedAttributes.OriginalPhone) { $setMailUserParams.Phone = $preservedAttributes.OriginalPhone }
                    if ($preservedAttributes.OriginalMobilePhone) { $setMailUserParams.MobilePhone = $preservedAttributes.OriginalMobilePhone }
                    if ($preservedAttributes.OriginalFax) { $setMailUserParams.Fax = $preservedAttributes.OriginalFax }
                    if ($preservedAttributes.OriginalCustomAttribute1) { $setMailUserParams.CustomAttribute1 = $preservedAttributes.OriginalCustomAttribute1 }
                    if ($preservedAttributes.OriginalCustomAttribute2) { $setMailUserParams.CustomAttribute2 = $preservedAttributes.OriginalCustomAttribute2 }
                    if ($preservedAttributes.OriginalCustomAttribute3) { $setMailUserParams.CustomAttribute3 = $preservedAttributes.OriginalCustomAttribute3 }
                    if ($preservedAttributes.OriginalCustomAttribute4) { $setMailUserParams.CustomAttribute4 = $preservedAttributes.OriginalCustomAttribute4 }
                    if ($preservedAttributes.OriginalCustomAttribute5) { $setMailUserParams.CustomAttribute5 = $preservedAttributes.OriginalCustomAttribute5 }
                    if ($preservedAttributes.OriginalExtensionCustomAttribute1) { $setMailUserParams.ExtensionCustomAttribute1 = $preservedAttributes.OriginalExtensionCustomAttribute1 }
                    if ($preservedAttributes.OriginalExtensionCustomAttribute2) { $setMailUserParams.ExtensionCustomAttribute2 = $preservedAttributes.OriginalExtensionCustomAttribute2 }
                    if ($preservedAttributes.OriginalExtensionCustomAttribute3) { $setMailUserParams.ExtensionCustomAttribute3 = $preservedAttributes.OriginalExtensionCustomAttribute3 }
                    if ($preservedAttributes.OriginalExtensionCustomAttribute4) { $setMailUserParams.ExtensionCustomAttribute4 = $preservedAttributes.OriginalExtensionCustomAttribute4 }
                    if ($preservedAttributes.OriginalExtensionCustomAttribute5) { $setMailUserParams.ExtensionCustomAttribute5 = $preservedAttributes.OriginalExtensionCustomAttribute5 }

                    Set-MailUser @setMailUserParams

                    # Update status in deletion record
                    $preservedAttributes.ConversionStatus = "Successfully Converted to MEU"
                }

                $result.ConversionStatus = "Successfully Converted to MEU"
                $result.Notes = "Mailbox disabled on-premises, MEU created in Exchange Online with external forwarding to $($result.ExternalForwardingAddress)"

                Write-Information "Successfully converted $($result.OriginalMailboxName) to MEU with forwarding to $($result.ExternalForwardingAddress)" -InformationAction Continue
            }
            catch
            {
                $errorResults += [PSCustomObject]@{
                    MailboxName = $result.OriginalMailboxName
                    MailboxUPN = $result.OriginalUPN
                    ErrorType = "MEUConversion"
                    ErrorMessage = $_.Exception.Message
                    ErrorDate = Get-Date
                }
                $result.ConversionStatus = "Failed"
                $result.Notes = "Conversion failed: $($_.Exception.Message)"

                # Update status in deletion record
                $preservedAttributes = $deletedMailboxAttributes | Where-Object { $_.OriginalUserPrincipalName -eq $result.OriginalUPN }
                if ($preservedAttributes)
                {
                    $preservedAttributes.ConversionStatus = "Conversion Failed: $($_.Exception.Message)"
                }

                Write-Error "Failed to convert $($result.OriginalMailboxName): $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Converting to MEU" -Completed
    }
    else
    {
        Write-Information "Simulation complete - no actual conversions performed" -InformationAction Continue
        foreach ($result in $conversionResults)
        {
            $result.ConversionStatus = "Simulated"
            $result.Notes = "Simulation: Would convert to MEU with external forwarding to $($result.ExternalForwardingAddress)"
        }
    }

    # Export results
    if ($conversionResults.Count -gt 0)
    {
        Write-Verbose "Exporting conversion results..."
        $conversionResults | Sort-Object OriginalMailboxName | Export-Csv $ConversionReportFile -NoTypeInformation
        Write-Information "Conversion report saved to: $ConversionReportFile" -InformationAction Continue
    }

    if ($deletedMailboxAttributes.Count -gt 0)
    {
        Write-Verbose "Exporting deleted mailbox attributes..."
        $deletedMailboxAttributes | Sort-Object OriginalDisplayName | Export-Csv $DeletedMailboxAttributesFile -NoTypeInformation
        Write-Information "Deleted mailbox attributes saved to: $DeletedMailboxAttributesFile" -InformationAction Continue
    }

    if ($errorResults.Count -gt 0)
    {
        Write-Verbose "Exporting error results..."
        $errorResults | Sort-Object MailboxName | Export-Csv $ErrorReportFile -NoTypeInformation
        Write-Information "Error report saved to: $ErrorReportFile" -InformationAction Continue
    }

    # Display summary
    Write-Information "MEU Conversion Process Summary:" -InformationAction Continue
    Write-Information "  Total Mailboxes Processed: $($conversionResults.Count)" -InformationAction Continue
    Write-Information "  Successful Preparations: $(($conversionResults | Where-Object { $_.ConversionStatus -notlike '*Failed*' }).Count)" -InformationAction Continue
    Write-Information "  Errors Encountered: $($errorResults.Count)" -InformationAction Continue
    Write-Information "  Mode: $(if ($WhatIfPreference) { 'Simulation' } else { 'Live Conversion' })" -InformationAction Continue
    Write-Information "  Conversion Report: $ConversionReportFile" -InformationAction Continue
    if ($deletedMailboxAttributes.Count -gt 0) { Write-Information "  Deleted Mailbox Attributes: $DeletedMailboxAttributesFile" -InformationAction Continue }
    if ($errorResults.Count -gt 0) { Write-Information "  Error Report: $ErrorReportFile" -InformationAction Continue }

}
catch
{
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
}
finally
{
    Write-Information "MEU Conversion Process completed." -InformationAction Continue
}
