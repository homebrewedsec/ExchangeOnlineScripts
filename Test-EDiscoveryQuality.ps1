<#
.SYNOPSIS
    Validates downloaded eDiscovery files against expected sizes from summary CSVs.

.DESCRIPTION
    This script compares local files against expected sizes from either the download
    summary CSV or export summary CSV. It identifies missing, corrupt, or partial files
    and outputs a quality check report that can be used with Invoke-EDiscoveryDownload.ps1
    -ReDownloadFrom to re-download failed files.

    Status values in output:
    - Pass: File exists and size matches expected (within tolerance)
    - Fail_Missing: File does not exist in folder
    - Fail_SizeMismatch: File size differs from expected by more than tolerance
    - Fail_TooSmall: File is less than 1KB (likely error response)

.PARAMETER FolderPath
    Path to the folder containing downloaded files (ZIPs or PSTs).

.PARAMETER SummaryCsvPath
    Path to the summary CSV file (from download or export script).

.PARAMETER SummaryType
    Type of summary CSV: "Download" (default) or "Export".
    - Download: Uses FileName and SizeMB columns from Invoke-EDiscoveryDownload output
    - Export: Uses UPN and ArchiveSizeBytes columns from Export-ArchiveMailbox output

.PARAMETER TolerancePercent
    Acceptable size variance percentage before flagging as mismatch. Default is 1%.

.PARAMETER OutputPath
    Directory for the quality check report. Defaults to current directory.

.PARAMETER DeleteBadFiles
    If specified, automatically deletes files that fail validation.

.EXAMPLE
    .\Test-EDiscoveryQuality.ps1 -FolderPath "C:\Downloads" -SummaryCsvPath ".\EDiscoveryDownload_Summary.csv"
    Validates files against download summary with default 1% tolerance.

.EXAMPLE
    .\Test-EDiscoveryQuality.ps1 -FolderPath "C:\Downloads" -SummaryCsvPath ".\summary.csv" -DeleteBadFiles
    Validates files and deletes any that fail validation.

.EXAMPLE
    .\Test-EDiscoveryQuality.ps1 -FolderPath "C:\Downloads" -SummaryCsvPath ".\ArchiveExport_Summary.csv" -SummaryType Export
    Validates against export summary (uses ArchiveSizeBytes for comparison).

.NOTES
    Author: Hudson Bush / Claude AI
    Version: 1.0

    Output CSV can be used with:
    Invoke-EDiscoveryDownload.ps1 -ReDownloadFrom ".\QualityCheck_*.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(Mandatory = $true)]
    [string]$SummaryCsvPath,

    [ValidateSet("Download", "Export")]
    [string]$SummaryType = "Download",

    [int]$TolerancePercent = 1,

    [string]$OutputPath = (Get-Location).Path,

    [switch]$DeleteBadFiles
)

#region CONFIGURATION
$script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:QualityReportPath = Join-Path $OutputPath "QualityCheck_$script:Timestamp.csv"
#endregion

#region VALIDATION
if (-not (Test-Path $FolderPath))
{
    throw "Folder not found: $FolderPath"
}

if (-not (Test-Path $SummaryCsvPath))
{
    throw "Summary CSV not found: $SummaryCsvPath"
}

if (-not (Test-Path $OutputPath))
{
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
#endregion

#region LOAD SUMMARY
Write-Host "Loading summary CSV: $SummaryCsvPath"
$summaryData = Import-Csv -Path $SummaryCsvPath

# Detect and validate columns based on summary type
if ($SummaryType -eq "Download")
{
    # Download summary uses FileName and SizeMB
    if (-not ($summaryData | Get-Member -Name "FileName" -MemberType NoteProperty))
    {
        throw "Download summary CSV must have 'FileName' column"
    }
    if (-not ($summaryData | Get-Member -Name "SizeMB" -MemberType NoteProperty))
    {
        throw "Download summary CSV must have 'SizeMB' column"
    }
    Write-Host "  Summary type: Download (using FileName, SizeMB columns)"
    Write-Host "  Files in summary: $($summaryData.Count)"
}
else
{
    # Export summary uses UPN and ArchiveSizeBytes (or ArchiveSizeGB)
    if (-not ($summaryData | Get-Member -Name "UPN" -MemberType NoteProperty))
    {
        throw "Export summary CSV must have 'UPN' column"
    }
    $hasBytes = $summaryData | Get-Member -Name "ArchiveSizeBytes" -MemberType NoteProperty
    $hasGB = $summaryData | Get-Member -Name "ArchiveSizeGB" -MemberType NoteProperty
    if (-not $hasBytes -and -not $hasGB)
    {
        throw "Export summary CSV must have 'ArchiveSizeBytes' or 'ArchiveSizeGB' column"
    }
    Write-Host "  Summary type: Export (using UPN, ArchiveSize columns)"
    Write-Host "  Records in summary: $($summaryData.Count)"
}
#endregion

#region QUALITY CHECK
Write-Host ""
Write-Host "Validating files in: $FolderPath"
Write-Host "Tolerance: $TolerancePercent%"
Write-Host ""

$qualityResults = @()
$passCount = 0
$failCount = 0

foreach ($record in $summaryData)
{
    $result = [PSCustomObject]@{
        FileName              = $null
        ExpectedSizeBytes     = 0
        ActualSizeBytes       = 0
        SizeDifferencePercent = 0
        FileExists            = $false
        Status                = "Unknown"
        ExportId              = $null
    }

    if ($SummaryType -eq "Download")
    {
        $result.FileName = $record.FileName
        $result.ExpectedSizeBytes = [long]($record.SizeMB * 1MB)
        $result.ExportId = $record.ExportId
    }
    else
    {
        # For export summary, construct expected filename pattern
        # Files are named like: Export_Archive_user_at_domain_com_*.zip
        $upn = $record.UPN
        $encodedEmail = $upn -replace "@", "_at_" -replace "\.", "_"

        # Find matching file in folder
        $matchingFiles = Get-ChildItem -Path $FolderPath -Filter "*$encodedEmail*" -File -ErrorAction SilentlyContinue
        if ($matchingFiles.Count -gt 0)
        {
            $result.FileName = $matchingFiles[0].Name
        }
        else
        {
            $result.FileName = "Export_Archive_$encodedEmail*.zip"  # Expected pattern
        }

        if ($record.ArchiveSizeBytes)
        {
            $result.ExpectedSizeBytes = [long]$record.ArchiveSizeBytes
        }
        elseif ($record.ArchiveSizeGB)
        {
            $result.ExpectedSizeBytes = [long]($record.ArchiveSizeGB * 1GB)
        }
    }

    # Check if file exists
    $filePath = Join-Path $FolderPath $result.FileName

    # For export summary with pattern matching, use the found file
    if ($SummaryType -eq "Export" -and $result.FileName -like "*`**")
    {
        $result.FileExists = $false
        $result.Status = "Fail_Missing"
    }
    elseif (Test-Path $filePath)
    {
        $result.FileExists = $true
        $fileInfo = Get-Item $filePath
        $result.ActualSizeBytes = $fileInfo.Length

        # Calculate size difference
        if ($result.ExpectedSizeBytes -gt 0)
        {
            $difference = [math]::Abs($result.ActualSizeBytes - $result.ExpectedSizeBytes)
            $result.SizeDifferencePercent = [math]::Round(($difference / $result.ExpectedSizeBytes) * 100, 2)
        }

        # Determine status
        if ($result.ActualSizeBytes -lt 1000)
        {
            $result.Status = "Fail_TooSmall"
        }
        elseif ($result.SizeDifferencePercent -gt $TolerancePercent)
        {
            $result.Status = "Fail_SizeMismatch"
        }
        else
        {
            $result.Status = "Pass"
        }
    }
    else
    {
        $result.FileExists = $false
        $result.Status = "Fail_Missing"
    }

    # Log result
    $statusColor = switch ($result.Status)
    {
        "Pass" { "Green" }
        "Fail_Missing" { "Red" }
        "Fail_SizeMismatch" { "Yellow" }
        "Fail_TooSmall" { "Red" }
        default { "White" }
    }

    $sizeInfo = ""
    if ($result.FileExists)
    {
        $actualMB = [math]::Round($result.ActualSizeBytes / 1MB, 2)
        $expectedMB = [math]::Round($result.ExpectedSizeBytes / 1MB, 2)
        $sizeInfo = " ($actualMB MB / $expectedMB MB expected, $($result.SizeDifferencePercent)% diff)"
    }

    Write-Host "  $($result.Status): $($result.FileName)$sizeInfo" -ForegroundColor $statusColor

    if ($result.Status -eq "Pass")
    {
        $passCount++
    }
    else
    {
        $failCount++
    }

    $qualityResults += $result
}
#endregion

#region DELETE BAD FILES
if ($DeleteBadFiles)
{
    $filesToDelete = $qualityResults | Where-Object { $_.Status -like "Fail_*" -and $_.FileExists }

    if ($filesToDelete.Count -gt 0)
    {
        Write-Host ""
        Write-Host "Deleting $($filesToDelete.Count) failed file(s)..." -ForegroundColor Yellow

        foreach ($file in $filesToDelete)
        {
            $filePath = Join-Path $FolderPath $file.FileName
            try
            {
                Remove-Item -Path $filePath -Force
                Write-Host "  Deleted: $($file.FileName)" -ForegroundColor Yellow
            }
            catch
            {
                Write-Host "  Failed to delete: $($file.FileName) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
#endregion

#region EXPORT REPORT
$qualityResults | Export-Csv -Path $script:QualityReportPath -NoTypeInformation

Write-Host ""
Write-Host "Quality Check Complete" -ForegroundColor Cyan
Write-Host "  Passed: $passCount"
Write-Host "  Failed: $failCount"
Write-Host "  Report: $script:QualityReportPath"

if ($failCount -gt 0)
{
    Write-Host ""
    Write-Host "To re-download failed files:" -ForegroundColor Yellow
    Write-Host "  .\Invoke-EDiscoveryDownload.ps1 -CaseName `"...`" -ClientId `"...`" -ReDownloadFrom `"$script:QualityReportPath`""
}
#endregion
