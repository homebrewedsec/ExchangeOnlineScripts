<#
.SYNOPSIS
Sets email forwarding and out-of-office messages for users from a CSV file

.DESCRIPTION
This script processes a CSV file containing user UPNs and new email addresses to:
1. Set email forwarding to the specified new email address
2. Configure out-of-office automatic replies with HTML template

The script sends OOO messages to all external senders and includes safety features
like WhatIf mode and detailed logging.

.PARAMETER InputCsvPath
Path to the CSV file containing user data. CSV must have 'upn' and 'newemail' columns.

.PARAMETER OutputPath
Directory path for output logs. Defaults to current directory.

.PARAMETER KeepCopy
When specified, keeps a copy of forwarded emails in the original mailbox.
By default, emails are forwarded without keeping a copy.

.PARAMETER DisableOOO
When specified, disables the out-of-office auto-reply instead of enabling it.
Use this to turn off OOO for users who previously had it enabled.

.PARAMETER OldCompanyName
The former company name to display in the OOO message. Required.

.PARAMETER NewCompanyName
The acquiring company name. Required.

.PARAMETER NewBrandName
The new brand name after acquisition. Required.

.PARAMETER AcquisitionDate
The date of acquisition to display in OOO message (e.g., "January 2, 2026"). Required.

.PARAMETER ForwardingEndDate
The date when automatic forwarding will stop (e.g., "April 1st 2026"). Required.

.PARAMETER NewWebsiteUrl
The new website URL. Required.

.PARAMETER NAStoreUrl
The North America e-store URL. Required.

.PARAMETER ContactUrl
The contact page URL. Required.

.PARAMETER WhatIf
Shows what would be changed without making actual modifications

.PARAMETER Confirm
Prompts for confirmation before making changes. Use -Confirm:$false to bypass prompts.

.PARAMETER Verbose
Provides detailed output during execution

.EXAMPLE
Set-MailboxForwarderAndOOO.ps1 -InputCsvPath "users.csv" -OldCompanyName "Acme Corp" -NewCompanyName "Contoso" -NewBrandName "Contoso Acme" -AcquisitionDate "January 2, 2026" -ForwardingEndDate "April 1st 2026" -NewWebsiteUrl "www.contoso-acme.com" -NAStoreUrl "https://store.contoso.com/parts" -ContactUrl "www.contoso-acme.com/contact" -WhatIf
Tests the script without making changes

.EXAMPLE
Set-MailboxForwarderAndOOO.ps1 -InputCsvPath "users.csv" -OldCompanyName "Acme Corp" -NewCompanyName "Contoso" -NewBrandName "Contoso Acme" -AcquisitionDate "January 2, 2026" -ForwardingEndDate "April 1st 2026" -NewWebsiteUrl "www.contoso-acme.com" -NAStoreUrl "https://store.contoso.com/parts" -ContactUrl "www.contoso-acme.com/contact" -KeepCopy
Processes all users, keeping a copy of forwarded emails in the original mailbox

.EXAMPLE
Set-MailboxForwarderAndOOO.ps1 -InputCsvPath "users.csv" -OldCompanyName "Acme Corp" -NewCompanyName "Contoso" -NewBrandName "Contoso Acme" -AcquisitionDate "January 2, 2026" -ForwardingEndDate "April 1st 2026" -NewWebsiteUrl "www.contoso-acme.com" -NAStoreUrl "https://store.contoso.com/parts" -ContactUrl "www.contoso-acme.com/contact" -Confirm:$false
Processes all users without confirmation prompts

.NOTES
Author: Hudson Bush, Seguri - hudson@seguri.io
Requires active Exchange Online PowerShell session.
Use Connect-ExchangeOnline before running this script.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'EnableOOO')]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [switch]$KeepCopy,

    [Parameter(ParameterSetName = 'DisableOOO')]
    [switch]$DisableOOO,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$OldCompanyName,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$NewCompanyName,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$NewBrandName,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$AcquisitionDate,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$ForwardingEndDate,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$NewWebsiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$NAStoreUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'EnableOOO')]
    [string]$ContactUrl
)

# Initialize variables
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "ForwarderAndOOO_Results_$timestamp.csv"
$results = @()
$processedCount = 0
$errorCount = 0

Write-Information "Starting email forwarding and OOO configuration..." -InformationAction Continue -Tags @('Status')
Write-Information "Input file: $InputCsvPath" -InformationAction Continue -Tags @('Config')
Write-Information "Output file: $outputFile" -InformationAction Continue -Tags @('Config')

# Validate input file exists
if (-not (Test-Path $InputCsvPath))
{
    Write-Error "Input CSV file not found: $InputCsvPath"
    exit 1
}

# Import and validate CSV
try
{
    $csvData = Import-Csv $InputCsvPath
    Write-Information "Imported $($csvData.Count) records from CSV" -InformationAction Continue -Tags @('Status')
}
catch
{
    Write-Error "Failed to import CSV file: $($_.Exception.Message)"
    exit 1
}

# Validate required columns
$requiredColumns = @('upn', 'newemail')
$csvColumns = $csvData[0].PSObject.Properties.Name

foreach ($column in $requiredColumns)
{
    if ($column -notin $csvColumns)
    {
        Write-Error "Required column '$column' not found in CSV. Available columns: $($csvColumns -join ', ')"
        exit 1
    }
}

Write-Information "CSV validation completed successfully" -InformationAction Continue -Tags @('Status')

# Validate Exchange Online connection
try
{
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Information "Exchange Online connection verified" -InformationAction Continue -Tags @('Status')
}
catch
{
    Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
    exit 1
}

# Process each user
foreach ($user in $csvData)
{
    $upn = $user.upn.Trim()
    $newEmail = $user.newemail.Trim()

    Write-Information "`nProcessing user: $upn -> $newEmail" -InformationAction Continue -Tags @('Progress')

    $result = [PSCustomObject]@{
        UPN = $upn
        NewEmail = $newEmail
        ForwardingStatus = "Not Processed"
        OOOStatus = "Not Processed"
        ErrorMessage = ""
        ProcessedTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    try
    {
        # Check if mailbox exists
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        Write-Information "Found mailbox: $($mailbox.DisplayName)" -InformationAction Continue -Tags @('Status')

        # Set email forwarding
        if ($PSCmdlet.ShouldProcess($upn, "Set forwarding to $newEmail"))
        {
            try
            {
                Set-Mailbox -Identity $upn -ForwardingAddress $null -ForwardingSmtpAddress $newEmail -DeliverToMailboxAndForward $KeepCopy -ErrorAction Stop
                $result.ForwardingStatus = "Success"
                Write-Information "  Forwarding configured successfully" -InformationAction Continue -Tags @('Success')
            }
            catch
            {
                $result.ForwardingStatus = "Failed"
                $result.ErrorMessage += "Forwarding error: $($_.Exception.Message); "
                Write-Warning "  Failed to set forwarding: $($_.Exception.Message)"
                $errorCount++
            }
        }
        else
        {
            $result.ForwardingStatus = "WhatIf - Would set forwarding"
            Write-Information "  [WhatIf] Would set forwarding to: $newEmail (KeepCopy: $KeepCopy)" -InformationAction Continue -Tags @('WhatIf')
        }

        # Set or disable out-of-office message
        if ($DisableOOO)
        {
            if ($PSCmdlet.ShouldProcess($upn, "Disable out-of-office message"))
            {
                try
                {
                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Disabled -ErrorAction Stop
                    $result.OOOStatus = "Disabled"
                    Write-Information "  Out-of-office disabled successfully" -InformationAction Continue -Tags @('Success')
                }
                catch
                {
                    $result.OOOStatus = "Failed"
                    $result.ErrorMessage += "OOO error: $($_.Exception.Message); "
                    Write-Warning "  Failed to disable out-of-office: $($_.Exception.Message)"
                    $errorCount++
                }
            }
            else
            {
                $result.OOOStatus = "WhatIf - Would disable OOO"
                Write-Information "  [WhatIf] Would disable OOO message" -InformationAction Continue -Tags @('WhatIf')
            }
        }
        else
        {
            if ($PSCmdlet.ShouldProcess($upn, "Set out-of-office message"))
            {
                try
                {
                    # HTML OOO message template
                    $oooMessage = @"
<p><strong>Subject:</strong> Announcement: Our Email and Website Addresses Have Changed (Formerly $OldCompanyName)</p>
<p>Dear Sender,</p>
<p>As of $AcquisitionDate, $OldCompanyName has been acquired by $NewCompanyName. We are now doing business as $NewBrandName.</p>
<p>Your email temporarily will be forwarded to our new email addresses. However, to ensure your ability to contact us, we recommend that you update your records:</p>
<p>My new email address is: <strong>$newEmail</strong></p>
<ul>
<li>Our new website address is: $NewWebsiteUrl</li>
<li>Our new e-store address for spares, consumables, and tools for customers in North America is: $NAStoreUrl</li>
<li>Our e-store for Europe and Asia is under construction (please contact your local office with queries).</li>
</ul>
<p>Contacts for sales &amp; service support can be found at: $ContactUrl</p>
<p>Your email has been forwarded to the appropriate person/department/party. This email is just a reminder to update your records. The automatic forwarding is only temporary and will stop on $ForwardingEndDate.</p>
<p>We appreciate your understanding during this transition. We look forward to continuing to serve you under the new $NewBrandName brand.</p>
<p>Sincerely,<br>$NewBrandName</p>
"@

                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -ExternalAudience All -InternalMessage $oooMessage -ExternalMessage $oooMessage -ErrorAction Stop
                    $result.OOOStatus = "Success"
                    Write-Information "  Out-of-office message configured successfully" -InformationAction Continue -Tags @('Success')
                }
                catch
                {
                    $result.OOOStatus = "Failed"
                    $result.ErrorMessage += "OOO error: $($_.Exception.Message); "
                    Write-Warning "  Failed to set out-of-office: $($_.Exception.Message)"
                    $errorCount++
                }
            }
            else
            {
                $result.OOOStatus = "WhatIf - Would set OOO message"
                Write-Information "  [WhatIf] Would set OOO message" -InformationAction Continue -Tags @('WhatIf')
            }
        }

        $processedCount++
    }
    catch
    {
        $result.ForwardingStatus = "Failed"
        $result.OOOStatus = "Failed"
        $result.ErrorMessage = "Mailbox not found or access error: $($_.Exception.Message)"
        Write-Error "  Failed to process $upn : $($_.Exception.Message)"
        $errorCount++
    }

    $results += $result

    # Progress indicator
    if ($csvData.Count -gt 0)
    {
        $percentComplete = [math]::Round(($processedCount / $csvData.Count) * 100, 1)
        Write-Progress -Activity "Processing Users" -Status "$processedCount of $($csvData.Count) processed" -PercentComplete $percentComplete
    }
}

# Export results
try
{
    $results | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
    Write-Information "`nResults exported to: $outputFile" -InformationAction Continue -Tags @('Status')
}
catch
{
    Write-Error "Failed to export results: $($_.Exception.Message)"
}

# Summary
Write-Information "`n=== PROCESSING SUMMARY ===" -InformationAction Continue -Tags @('Summary')
Write-Information "Total users processed: $($csvData.Count)" -InformationAction Continue -Tags @('Summary')
Write-Information "Successful operations: $($processedCount - $errorCount)" -InformationAction Continue -Tags @('Summary')
Write-Information "Failed operations: $errorCount" -InformationAction Continue -Tags @('Summary')

if ($errorCount -gt 0)
{
    Write-Information "`nCheck the output CSV file for detailed error information." -InformationAction Continue -Tags @('Warning')
}

Write-Information "`nScript completed at $(Get-Date)" -InformationAction Continue -Tags @('Status')
