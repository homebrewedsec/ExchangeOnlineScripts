<#
.SYNOPSIS
Updates primary SMTP addresses in on-premises Active Directory based on CSV input with new domain

.DESCRIPTION
This script takes a CSV file with email addresses and updates the primary SMTP address in on-premises
Active Directory by swapping the domain. It checks if the new email address already exists in
proxyAddresses before making changes and provides comprehensive logging.

.PARAMETER InputCsvPath
Path to CSV file containing email addresses with "email" column header

.PARAMETER NewDomain
The new domain to replace the existing domain in email addresses

.PARAMETER UpdateUPN
Switch to also update the UserPrincipalName to match the new primary SMTP address

.PARAMETER OutputPath
Directory path for output files. Defaults to current directory.

.PARAMETER WhatIf
Shows what would be changed without making actual modifications

.PARAMETER Verbose
Provides detailed output during execution

.EXAMPLE
.\Update-ADEmailAddresses.ps1 -InputCsvPath "emails.csv" -NewDomain "newdomain.com"

.EXAMPLE
.\Update-ADEmailAddresses.ps1 -InputCsvPath "emails.csv" -NewDomain "newdomain.com" -UpdateUPN -WhatIf

.NOTES
Author: Hudson Bush, Seguri - hudson@seguri.io
Requires Active Directory PowerShell module and appropriate permissions to modify user objects
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$NewDomain,

    [switch]$UpdateUPN,

    [string]$OutputPath = (Get-Location).Path
)

# Generate log file path
$LogPath = Join-Path $OutputPath "ADEmailUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Import Active Directory module
try
{
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch
{
    Write-Error "Failed to import Active Directory module: $($_.Exception.Message)"
    exit 1
}

# Function to write to log file with timestamp
function Write-LogEntry
{
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"

    # Write to console
    switch ($Level)
    {
        "ERROR" { Write-Error $Message }
        "WARNING" { Write-Warning $Message }
        default { Write-Information $logEntry -InformationAction Continue }
    }

    # Write to log file
    try
    {
        Add-Content -Path $LogPath -Value $logEntry -ErrorAction Stop
    }
    catch
    {
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

# Validate input CSV file
if (-not (Test-Path $InputCsvPath))
{
    Write-LogEntry "Input CSV file not found: $InputCsvPath" "ERROR"
    exit 1
}

Write-LogEntry "Starting AD email address update process"
Write-LogEntry "Input CSV: $InputCsvPath"
Write-LogEntry "New Domain: $NewDomain"
Write-LogEntry "Update UPN: $UpdateUPN"
Write-LogEntry "WhatIf Mode: $WhatIfPreference"

# Import CSV data
try
{
    $csvData = Import-Csv -Path $InputCsvPath -ErrorAction Stop
    Write-LogEntry "Successfully imported $($csvData.Count) records from CSV"
}
catch
{
    Write-LogEntry "Failed to import CSV file: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Validate CSV has email column
if (-not ($csvData | Get-Member -Name "email" -MemberType NoteProperty))
{
    Write-LogEntry "CSV file must contain 'email' column header" "ERROR"
    exit 1
}

# Initialize counters
$processedCount = 0
$successCount = 0
$errorCount = 0
$skippedCount = 0

# Process each email address
foreach ($record in $csvData)
{
    $processedCount++
    $currentEmail = $record.email

    if ([string]::IsNullOrWhiteSpace($currentEmail))
    {
        Write-LogEntry "Skipping empty email address at record $processedCount" "WARNING"
        $skippedCount++
        continue
    }

    Write-LogEntry "Processing email: $currentEmail"

    # Calculate new email address
    try
    {
        $emailParts = $currentEmail -split "@"
        if ($emailParts.Count -ne 2)
        {
            Write-LogEntry "Invalid email format: $currentEmail" "WARNING"
            $skippedCount++
            continue
        }

        $newEmail = "$($emailParts[0])@$NewDomain"
    }
    catch
    {
        Write-LogEntry "Failed to calculate new email for $currentEmail`: $($_.Exception.Message)" "ERROR"
        $errorCount++
        continue
    }

    # Find user in AD by current email
    try
    {
        $adUser = Get-ADUser -Filter "mail -eq '$currentEmail' -or proxyAddresses -like 'SMTP:$currentEmail' -or UserPrincipalName -eq '$currentEmail'" -Properties mail, proxyAddresses, UserPrincipalName -ErrorAction Stop

        if (-not $adUser)
        {
            Write-LogEntry "User not found in AD with email: $currentEmail" "WARNING"
            $skippedCount++
            continue
        }

        if ($adUser.Count -gt 1)
        {
            Write-LogEntry "Multiple users found with email $currentEmail. Skipping." "WARNING"
            $skippedCount++
            continue
        }

    }
    catch
    {
        Write-LogEntry "Error searching for user with email $currentEmail`: $($_.Exception.Message)" "ERROR"
        $errorCount++
        continue
    }

    # Check if new email already exists in proxyAddresses
    $existingPrimary = $adUser.proxyAddresses | Where-Object { $_ -ceq "SMTP:$newEmail" }
    $existingSecondary = $adUser.proxyAddresses | Where-Object { $_ -ceq "smtp:$newEmail" }

    if ($existingPrimary)
    {
        Write-LogEntry "New email $newEmail is already the primary SMTP address for user $($adUser.SamAccountName). Skipping." "WARNING"
        $skippedCount++
        continue
    }

    # Flag to indicate we need to promote existing secondary to primary
    $promoteExisting = $false
    if ($existingSecondary)
    {
        Write-LogEntry "New email $newEmail exists as secondary address for user $($adUser.SamAccountName). Will promote to primary."
        $promoteExisting = $true
    }

    # Prepare changes
    try
    {
        # Get current proxyAddresses and identify old primary SMTP
        $oldPrimarySMTP = $adUser.proxyAddresses | Where-Object { $_ -like "SMTP:*" }
        $currentProxyAddresses = @($adUser.proxyAddresses | Where-Object { $_ -notlike "SMTP:*" })

        # If promoting existing secondary to primary, remove it from the list to avoid duplicates
        if ($promoteExisting)
        {
            $currentProxyAddresses = @($currentProxyAddresses | Where-Object { $_ -cne "smtp:$newEmail" })
        }

        # Log step 1: Remove old primary SMTP
        if ($oldPrimarySMTP)
        {
            if ($WhatIfPreference)
            {
                Write-LogEntry "WHATIF: Would remove old primary SMTP: $oldPrimarySMTP for user $($adUser.SamAccountName)"
            }
            else
            {
                Write-LogEntry "Step 1: Removing old primary SMTP: $oldPrimarySMTP for user $($adUser.SamAccountName)"
            }
        }

        # Prepare new proxyAddresses array
        $newProxyAddresses = @("SMTP:$newEmail") + $currentProxyAddresses

        # Log step 2: Add new primary SMTP
        if ($WhatIfPreference)
        {
            Write-LogEntry "WHATIF: Would add new primary SMTP: SMTP:$newEmail for user $($adUser.SamAccountName)"
        }
        else
        {
            Write-LogEntry "Step 2: Adding new primary SMTP: SMTP:$newEmail for user $($adUser.SamAccountName)"
        }

        # Add old primary as secondary if it exists and is different from new email
        if ($adUser.mail -and $adUser.mail -ne $newEmail)
        {
            $newProxyAddresses += "smtp:$($adUser.mail)"

            # Log step 3: Add old as secondary
            if ($WhatIfPreference)
            {
                Write-LogEntry "WHATIF: Would add old email as secondary: smtp:$($adUser.mail) for user $($adUser.SamAccountName)"
            }
            else
            {
                Write-LogEntry "Step 3: Adding old email as secondary: smtp:$($adUser.mail) for user $($adUser.SamAccountName)"
            }
        }

        $updateParams = @{
            Identity = $adUser.SamAccountName
            Replace = @{
                mail = $newEmail
                proxyAddresses = $newProxyAddresses
            }
        }

        # Add UPN update if requested
        if ($UpdateUPN)
        {
            $updateParams.Replace.UserPrincipalName = $newEmail

            if ($WhatIfPreference)
            {
                Write-LogEntry "WHATIF: Would update UPN to $newEmail for user $($adUser.SamAccountName)"
            }
            else
            {
                Write-LogEntry "Step 4: Updating UPN to $newEmail for user $($adUser.SamAccountName)"
            }
        }

        # Apply changes
        if ($WhatIfPreference)
        {
            Write-LogEntry "WHATIF: All changes would be applied for user $($adUser.SamAccountName)"
            $successCount++
        }
        else
        {
            Set-ADUser @updateParams -ErrorAction Stop
            Write-LogEntry "SUCCESS: All changes applied successfully for user $($adUser.SamAccountName) - Email changed from $($adUser.mail) to $newEmail"
            $successCount++
        }
    }
    catch
    {
        Write-LogEntry "ERROR: Failed to update user $($adUser.SamAccountName): $($_.Exception.Message)" "ERROR"
        $errorCount++
    }
}

# Final summary
Write-LogEntry "=== SUMMARY ==="
Write-LogEntry "Total records processed: $processedCount"
Write-LogEntry "Successful updates: $successCount"
Write-LogEntry "Errors: $errorCount"
Write-LogEntry "Skipped: $skippedCount"
Write-LogEntry "Log file saved to: $LogPath"

if ($WhatIfPreference)
{
    Write-LogEntry "WhatIf mode was enabled - no actual changes were made"
}

Write-Information "`nScript completed. Check log file for details: $LogPath" -InformationAction Continue
