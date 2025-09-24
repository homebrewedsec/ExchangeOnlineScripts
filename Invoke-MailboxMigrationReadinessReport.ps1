<#
.SYNOPSIS
Assesses mailboxes for potential migration or eDiscovery export issues

.DESCRIPTION
This script evaluates Exchange Online mailboxes to identify potential issues that could
affect migration or eDiscovery export processes. It checks for:
- Large mailboxes (>100GB)
- Auto-expanding archives
- Litigation hold status
- In-Place hold policies
- Retention policies
- Mailbox health and consistency issues
- Folder statistics and item counts

The script takes a CSV input with mailbox UPNs and generates a detailed assessment report.

.PARAMETER InputCsvPath
Path to CSV file containing mailbox UPNs to assess. Must have a 'upn' column.

.PARAMETER OutputPath
Directory path for output files. Defaults to current directory.

.PARAMETER SizeThresholdGB
Size threshold in GB for flagging large mailboxes. Defaults to 100GB.

.PARAMETER IncludeDetailedStats
Include detailed folder statistics and item counts (slower but more comprehensive).

.PARAMETER IncludeMigrationReadinessSuccess
Include mailboxes that are migration ready in the output. By default, only mailboxes with issues are exported.

.EXAMPLE
Invoke-MailboxMigrationReadinessReport.ps1 -InputCsvPath "mailboxes.csv"

.EXAMPLE
Invoke-MailboxMigrationReadinessReport.ps1 -InputCsvPath "mailboxes.csv" -SizeThresholdGB 50 -IncludeDetailedStats

.EXAMPLE
Invoke-MailboxMigrationReadinessReport.ps1 -InputCsvPath "mailboxes.csv" -IncludeMigrationReadinessSuccess -OutputPath "C:\Reports"

.NOTES
Requires active Exchange Online PowerShell session.
Use Connect-ExchangeOnline before running this script.

Some features require specific Exchange Online permissions:
- Litigation Hold: Compliance Management or eDiscovery Manager
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [int]$SizeThresholdGB = 100,

    [Parameter()]
    [switch]$IncludeDetailedStats,

    [Parameter()]
    [switch]$IncludeMigrationReadinessSuccess
)

# Initialize variables
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "MailboxMigrationReadiness_$timestamp.csv"
$results = @()
$processedCount = 0
$errorCount = 0

Write-Information "Starting mailbox migration readiness assessment..." -InformationAction Continue -Tags @('Status')
Write-Information "Input file: $InputCsvPath" -InformationAction Continue -Tags @('Config')
Write-Information "Output file: $outputFile" -InformationAction Continue -Tags @('Config')
Write-Information "Size threshold: $SizeThresholdGB GB" -InformationAction Continue -Tags @('Config')

# Test Exchange Online connection
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
$requiredColumns = @('upn')
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

# Helper function to convert bytes to GB
function ConvertTo-GB
{
    param([long]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

# Helper function to parse mailbox statistics
function Get-MailboxSizeInfo
{
    param($Stats)

    $sizeInfo = @{
        TotalItemSizeGB = 0
        ItemCount = 0
        DeletedItemSizeGB = 0
        DeletedItemCount = 0
    }

    if ($Stats.TotalItemSize)
    {
        $sizeString = $Stats.TotalItemSize.ToString()
        if ($sizeString -match '([\d,]+)\s*bytes')
        {
            $bytes = [long]($matches[1] -replace ',', '')
            $sizeInfo.TotalItemSizeGB = ConvertTo-GB $bytes
        }
    }

    if ($Stats.ItemCount)
    {
        $sizeInfo.ItemCount = $Stats.ItemCount
    }

    if ($Stats.TotalDeletedItemSize)
    {
        $sizeString = $Stats.TotalDeletedItemSize.ToString()
        if ($sizeString -match '([\d,]+)\s*bytes')
        {
            $bytes = [long]($matches[1] -replace ',', '')
            $sizeInfo.DeletedItemSizeGB = ConvertTo-GB $bytes
        }
    }

    if ($Stats.DeletedItemCount)
    {
        $sizeInfo.DeletedItemCount = $Stats.DeletedItemCount
    }

    return $sizeInfo
}

# Process each mailbox
foreach ($user in $csvData)
{
    $processedCount++
    $upn = $user.upn.Trim()

    Write-Information "Processing mailbox $processedCount of $($csvData.Count): $upn" -InformationAction Continue -Tags @('Progress')

    # Initialize result object
    $result = [PSCustomObject]@{
        UPN = $upn
        DisplayName = ""
        MailboxType = ""
        TotalSizeGB = 0
        ItemCount = 0
        ArchiveEnabled = $false
        ArchiveSizeGB = 0
        ArchiveItemCount = 0
        AutoExpandingArchive = $false
        LitigationHoldEnabled = $false
        LitigationHoldDate = ""
        InPlaceHolds = ""
        RetentionPolicy = ""
        ManagedFolderPolicy = ""
        SingleItemRecoveryEnabled = $false
        RetainDeletedItemsFor = ""
        RecoverableItemsSize = 0
        RecoverableItemsCount = 0
        SizeIssue = $false
        ArchiveIssue = $false
        ComplianceIssue = $false
        HealthIssue = $false
        IssuesSummary = ""
        LastLogonTime = ""
        FolderCount = 0
        LargestFolderSizeGB = 0
        LargestFolderName = ""
        IsExchangeCloudManaged = $false
        IsInactiveMailbox = $false
        IsSoftDeletedByDisable = $false
        IsSoftDeletedByRemove = $false
        UserSMimeCertificate = ""
        ProcessingStatus = "Success"
        ErrorDetails = ""
        MigrationReadinessStatus = $true
        FailureReason = ""
    }

    try
    {
        # Get mailbox information
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        $result.DisplayName = $mailbox.DisplayName
        $result.MailboxType = $mailbox.RecipientTypeDetails
        $result.ArchiveEnabled = $mailbox.ArchiveStatus -eq "Active"
        $result.AutoExpandingArchive = $mailbox.AutoExpandingArchiveEnabled
        $result.LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        $result.LitigationHoldDate = $mailbox.LitigationHoldDate
        $result.InPlaceHolds = ($mailbox.InPlaceHolds -join "; ")
        $result.RetentionPolicy = $mailbox.RetentionPolicy
        $result.ManagedFolderPolicy = $mailbox.ManagedFolderMailboxPolicy
        $result.SingleItemRecoveryEnabled = $mailbox.SingleItemRecoveryEnabled
        $result.RetainDeletedItemsFor = $mailbox.RetainDeletedItemsFor
        $result.IsExchangeCloudManaged = $mailbox.IsExchangeCloudManaged
        $result.IsInactiveMailbox = $mailbox.IsInactiveMailbox
        $result.IsSoftDeletedByDisable = $mailbox.IsSoftDeletedByDisable
        $result.IsSoftDeletedByRemove = $mailbox.IsSoftDeletedByRemove
        $result.UserSMimeCertificate = if ($mailbox.UserSMimeCertificate) { ($mailbox.UserSMimeCertificate -join "; ") } else { "" }

        Write-Information "  Found mailbox: $($mailbox.DisplayName)" -InformationAction Continue -Tags @('Progress')

        # Get mailbox statistics
        $mailboxStats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
        $sizeInfo = Get-MailboxSizeInfo $mailboxStats
        $result.TotalSizeGB = $sizeInfo.TotalItemSizeGB
        $result.ItemCount = $sizeInfo.ItemCount
        $result.LastLogonTime = $mailboxStats.LastLogonTime

        # Get recoverable items statistics
        if ($mailboxStats.TotalItemSize)
        {
            $recoverableStats = Get-MailboxFolderStatistics -Identity $upn -FolderScope RecoverableItems -ErrorAction SilentlyContinue
            if ($recoverableStats)
            {
                $totalRecoverableSize = 0
                foreach ($folder in $recoverableStats)
                {
                    if ($folder.FolderSize)
                    {
                        $sizeString = $folder.FolderSize.ToString()
                        if ($sizeString -match '([\d,]+)\s*bytes')
                        {
                            $bytes = [long]($matches[1] -replace ',', '')
                            $totalRecoverableSize += $bytes
                        }
                    }
                }
                if ($totalRecoverableSize -gt 0)
                {
                    $result.RecoverableItemsSize = ConvertTo-GB $totalRecoverableSize
                }
                $result.RecoverableItemsCount = ($recoverableStats | Measure-Object ItemsInFolder -Sum).Sum
            }
        }

        # Get archive statistics if enabled
        if ($result.ArchiveEnabled)
        {
            try
            {
                $archiveStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction Stop
                $archiveSizeInfo = Get-MailboxSizeInfo $archiveStats
                $result.ArchiveSizeGB = $archiveSizeInfo.TotalItemSizeGB
                $result.ArchiveItemCount = $archiveSizeInfo.ItemCount

                Write-Information "  Archive size: $($result.ArchiveSizeGB) GB" -InformationAction Continue -Tags @('Progress')
            }
            catch
            {
                $result.ArchiveIssue = $true
                Write-Information "  Warning: Could not retrieve archive statistics" -InformationAction Continue -Tags @('Warning')
            }
        }

        # Get detailed folder statistics if requested
        if ($IncludeDetailedStats)
        {
            try
            {
                $folderStats = Get-MailboxFolderStatistics -Identity $upn -ErrorAction Stop
                $result.FolderCount = $folderStats.Count

                # Find largest folder by size
                $largestFolder = $null
                $largestSizeBytes = 0
                foreach ($folder in $folderStats)
                {
                    if ($folder.FolderSize)
                    {
                        $sizeString = $folder.FolderSize.ToString()
                        if ($sizeString -match '([\d,]+)\s*bytes')
                        {
                            $bytes = [long]($matches[1] -replace ',', '')
                            if ($bytes -gt $largestSizeBytes)
                            {
                                $largestSizeBytes = $bytes
                                $largestFolder = $folder
                            }
                        }
                    }
                }
                if ($largestFolder -and $largestSizeBytes -gt 0)
                {
                    $result.LargestFolderSizeGB = ConvertTo-GB $largestSizeBytes
                    $result.LargestFolderName = $largestFolder.Name
                }
            }
            catch
            {
                Write-Information "  Warning: Could not retrieve detailed folder statistics" -InformationAction Continue -Tags @('Warning')
            }
        }

        # Check for potential issues
        $issues = @()

        # Size issues
        if ($result.TotalSizeGB -gt $SizeThresholdGB)
        {
            $result.SizeIssue = $true
            $issues += "Large mailbox (>$SizeThresholdGB GB)"
        }

        # Archive issues
        if ($result.AutoExpandingArchive)
        {
            $result.ArchiveIssue = $true
            $issues += "Auto-expanding archive enabled"
        }

        if ($result.ArchiveSizeGB -gt $SizeThresholdGB)
        {
            $result.ArchiveIssue = $true
            $issues += "Large archive (>$SizeThresholdGB GB)"
        }

        # Compliance issues
        if ($result.LitigationHoldEnabled)
        {
            $result.ComplianceIssue = $true
            $issues += "Litigation hold enabled"
        }

        if ($result.InPlaceHolds -ne "")
        {
            $result.ComplianceIssue = $true
            $issues += "In-Place holds applied"
        }

        if ($result.RecoverableItemsSize -gt 10)
        {
            $result.ComplianceIssue = $true
            $issues += "Large recoverable items (>10 GB)"
        }

        # Health issues (basic checks)
        if (-not $result.LastLogonTime -or $result.LastLogonTime -lt (Get-Date).AddDays(-90))
        {
            $result.HealthIssue = $true
            $issues += "Inactive mailbox (no logon >90 days)"
        }

        if ($result.ItemCount -gt 500000)
        {
            $result.HealthIssue = $true
            $issues += "High item count (>500,000)"
        }

        $result.IssuesSummary = $issues -join "; "

        # Set migration readiness status based on issues
        if ($issues.Count -gt 0)
        {
            $result.MigrationReadinessStatus = $false
            $result.FailureReason = $issues -join "; "
        }
        else
        {
            $result.MigrationReadinessStatus = $true
            $result.FailureReason = ""
        }

        Write-Information "  Assessment completed - Issues: $($issues.Count)" -InformationAction Continue -Tags @('Progress')
    }
    catch
    {
        $result.ProcessingStatus = "Failed"
        $result.ErrorDetails = $_.Exception.Message
        $result.MigrationReadinessStatus = $false
        $result.FailureReason = "Processing failed: $($_.Exception.Message)"
        Write-Information "  Failed to process $upn : $($_.Exception.Message)" -InformationAction Continue -Tags @('Warning')
        $errorCount++
    }

    $results += $result

    # Progress indicator
    if ($csvData.Count -gt 0)
    {
        $percentComplete = [math]::Round(($processedCount / $csvData.Count) * 100, 1)
        Write-Progress -Activity "Assessing Mailboxes" -Status "$processedCount of $($csvData.Count) processed" -PercentComplete $percentComplete
    }
}

# Filter results based on IncludeMigrationReadinessSuccess parameter
$resultsToExport = if ($IncludeMigrationReadinessSuccess)
{
    # Include all results
    $results
}
else
{
    # Only include results with migration readiness issues
    $results | Where-Object {$_.MigrationReadinessStatus -eq $false}
}

# Export results only if there are any to export
if ($resultsToExport.Count -gt 0)
{
    try
    {
        $resultsToExport | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
        Write-Information "Results exported to: $outputFile" -InformationAction Continue -Tags @('Status')
        Write-Information "Exported $($resultsToExport.Count) of $($results.Count) total mailboxes" -InformationAction Continue -Tags @('Status')
    }
    catch
    {
        Write-Error "Failed to export results: $($_.Exception.Message)"
        exit 1
    }
}
else
{
    Write-Information "No results to export - no CSV file created" -InformationAction Continue -Tags @('Status')
    if (-not $IncludeMigrationReadinessSuccess)
    {
        Write-Information "All mailboxes are migration ready! Use -IncludeMigrationReadinessSuccess to see all results" -InformationAction Continue -Tags @('Status')
    }
}

# Generate summary
Write-Information "" -InformationAction Continue -Tags @('Summary')
Write-Information "=== MIGRATION READINESS ASSESSMENT SUMMARY ===" -InformationAction Continue -Tags @('Summary')
Write-Information "Total mailboxes assessed: $($csvData.Count)" -InformationAction Continue -Tags @('Summary')
Write-Information "Successful assessments: $(($results | Where-Object {$_.ProcessingStatus -eq 'Success'}).Count)" -InformationAction Continue -Tags @('Summary')
Write-Information "Failed assessments: $errorCount" -InformationAction Continue -Tags @('Summary')
Write-Information "" -InformationAction Continue -Tags @('Summary')

$sizeIssues = ($results | Where-Object {$_.SizeIssue}).Count
$archiveIssues = ($results | Where-Object {$_.ArchiveIssue}).Count
$complianceIssues = ($results | Where-Object {$_.ComplianceIssue}).Count
$healthIssues = ($results | Where-Object {$_.HealthIssue}).Count

Write-Information "POTENTIAL MIGRATION ISSUES:" -InformationAction Continue -Tags @('Summary')
Write-Information "Size issues (>$SizeThresholdGB GB): $sizeIssues" -InformationAction Continue -Tags @('Summary')
Write-Information "Archive issues: $archiveIssues" -InformationAction Continue -Tags @('Summary')
Write-Information "Compliance issues: $complianceIssues" -InformationAction Continue -Tags @('Summary')
Write-Information "Health issues: $healthIssues" -InformationAction Continue -Tags @('Summary')

Write-Information "" -InformationAction Continue -Tags @('Summary')
Write-Information "Script completed at $(Get-Date)" -InformationAction Continue -Tags @('Summary')
