<#
.SYNOPSIS
Sets email forwarding and out-of-office messages for users from a CSV file

.DESCRIPTION
This script processes a CSV file containing user UPNs and new email addresses to:
1. Set email forwarding to the specified new email address
2. Configure out-of-office automatic replies (placeholder implementation)

The script includes safety features like WhatIf mode and detailed logging.

.PARAMETER InputCsvPath
Path to the CSV file containing user data. CSV must have 'upn' and 'newemail' columns.

.PARAMETER WhatIf
Shows what would be changed without making actual modifications

.PARAMETER Force
Skips confirmation prompts for batch processing

.PARAMETER OutputPath
Directory path for output logs. Defaults to current directory.

.EXAMPLE
Set-MailboxForwarderAndOOO.ps1 -InputCsvPath "users.csv" -WhatIf
Tests the script without making changes

.EXAMPLE
Set-MailboxForwarderAndOOO.ps1 -InputCsvPath "users.csv" -Force
Processes all users without confirmation prompts

.NOTES
Requires active Exchange Online PowerShell session.
Use Connect-ExchangeOnline before running this script.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter()]
    [string]$OutputPath = ".",

    [Parameter()]
    [switch]$Force  # Used in ShouldProcess calls for confirmation prompts
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
        if ($Force -or $PSCmdlet.ShouldProcess($upn, "Set forwarding to $newEmail"))
        {
            try
            {
                Set-Mailbox -Identity $upn -ForwardingAddress $null -ForwardingSmtpAddress $newEmail -DeliverToMailboxAndForward $false -ErrorAction Stop
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
            Write-Information "  [WhatIf] Would set forwarding to: $newEmail" -InformationAction Continue -Tags @('WhatIf')
        }

        # Set out-of-office message (placeholder implementation)
        if ($Force -or $PSCmdlet.ShouldProcess($upn, "Set out-of-office message"))
        {
            try
            {
                # Placeholder OOO message
                $oooMessage = "I am no longer with the organization. Please contact me at my new email address: $newEmail"

                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $oooMessage -ExternalMessage $oooMessage -ErrorAction Stop
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