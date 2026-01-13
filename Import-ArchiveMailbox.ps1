<#
.SYNOPSIS
    Extracts ZIP files containing PSTs and generates Purview import mapping CSV.

.DESCRIPTION
    This script prepares PST files for import to Exchange Online via Microsoft Purview:
    1. Extracts ZIP files containing exchange.pst to organized subfolders
    2. Matches extracted PSTs to source/target mapping from CSV
    3. Generates Purview-compatible import mapping CSV
    4. Displays AzCopy command for uploading to Azure

    IMPORT PROCESS (after running this script):
    1. Run the AzCopy command displayed to upload PSTs to Azure
    2. Navigate to https://compliance.microsoft.com
    3. Go to: Data lifecycle management > Import
    4. Create new import job and upload the generated mapping CSV
    5. Start the import job

.PARAMETER ZipFolderPath
    Folder containing ZIP files to extract.
    ZIP files should contain exchange.pst inside.

.PARAMETER MappingCsvPath
    CSV file with source-to-target mailbox mapping.
    Required columns: SourceEmail (or SourceUPN), TargetEmail (or TargetUPN)

.PARAMETER OutputPath
    Directory where PSTs will be extracted to subfolders.
    Default: Current directory

.PARAMETER AzureSasUrl
    Optional. Azure Blob Storage SAS URL from Purview portal.
    If provided, displays the complete AzCopy command to run.

.PARAMETER ZipNamePattern
    Pattern for ZIP filenames with {email} placeholder.
    The email in the filename has @ and . replaced with _
    Default: "{email}.zip"
    Example patterns: "{email}.zip", "Export_{email}.zip", "{email}_archive.zip"

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv"

    Extracts ZIPs named like user_domain_com.zip and generates Purview mapping CSV.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv" -ZipNamePattern "Export_{email}.zip"

    Handles ZIPs with "Export_" prefix.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv" -AzureSasUrl "https://..."

    Includes complete AzCopy command in output.

.NOTES
    Author: Hudson Bush
    Requires: PowerShell 5.1+, AzCopy for upload step

    Prerequisites for Purview Import:
    - Mailbox Import Export role in Exchange Online
    - Compliance Administrator or eDiscovery Manager role
    - AzCopy installed (https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10)

    Reference Documentation:
    - https://learn.microsoft.com/en-us/purview/use-network-upload-to-import-pst-files
    - https://learn.microsoft.com/en-us/purview/importing-pst-files-to-office-365
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ZipFolderPath,

    [Parameter(Mandatory = $true)]
    [string]$MappingCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [string]$AzureSasUrl,

    [Parameter(Mandatory = $false)]
    [string]$ZipNamePattern = "{email}.zip"
)

#region VALIDATION
Write-Output "Validating parameters..."

# Check if ZIP folder exists
if (-not (Test-Path $ZipFolderPath))
{
    Write-Error "ZIP folder not found: $ZipFolderPath"
    exit 1
}

# Check if mapping CSV exists
if (-not (Test-Path $MappingCsvPath))
{
    Write-Error "Mapping CSV not found: $MappingCsvPath"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $OutputPath))
{
    Write-Output "Creating output directory: $OutputPath"
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Load and validate mapping CSV
$mappingCsv = Import-Csv $MappingCsvPath
if (-not $mappingCsv -or $mappingCsv.Count -eq 0)
{
    Write-Error "Mapping CSV is empty: $MappingCsvPath"
    exit 1
}

# Determine source column name
$sourceColumn = $null
$targetColumn = $null
$csvColumns = $mappingCsv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

foreach ($col in @("SourceEmail", "SourceUPN", "sourceEmail", "sourceUPN", "Source"))
{
    if ($col -in $csvColumns)
    {
        $sourceColumn = $col
        break
    }
}

foreach ($col in @("TargetEmail", "TargetUPN", "targetEmail", "targetUPN", "Target"))
{
    if ($col -in $csvColumns)
    {
        $targetColumn = $col
        break
    }
}

if (-not $sourceColumn)
{
    Write-Error "Mapping CSV must contain a source column (SourceEmail, SourceUPN, or Source)"
    exit 1
}

if (-not $targetColumn)
{
    Write-Error "Mapping CSV must contain a target column (TargetEmail, TargetUPN, or Target)"
    exit 1
}

Write-Output "Using columns: Source='$sourceColumn', Target='$targetColumn'"
Write-Output "Found $($mappingCsv.Count) mappings in CSV."
Write-Output ""

# Build lookup hashtable for source -> target mapping
$sourceToTarget = @{}
foreach ($row in $mappingCsv)
{
    $source = $row.$sourceColumn.Trim().ToLower()
    $target = $row.$targetColumn.Trim()
    $sourceToTarget[$source] = $target
}

# Check for AzCopy
$azcopyPath = Get-Command azcopy -ErrorAction SilentlyContinue
if (-not $azcopyPath)
{
    Write-Warning "AzCopy not found in PATH. You will need it for the upload step."
    Write-Warning "Download from: https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10"
    Write-Output ""
}

Write-Output "Validation complete."
Write-Output ""
#endregion

#region STEP 1: ENUMERATE AND EXTRACT ZIP FILES
Write-Output "Step 1: Extracting ZIP files..."
Write-Output "================================"
Write-Output ""

# Get all ZIP files
$zipFiles = Get-ChildItem -Path $ZipFolderPath -Filter "*.zip"
if ($zipFiles.Count -eq 0)
{
    Write-Error "No ZIP files found in: $ZipFolderPath"
    exit 1
}

Write-Output "Found $($zipFiles.Count) ZIP files."
Write-Output ""

# Build regex pattern from ZipNamePattern
# Convert {email} placeholder to regex capture group
# Email in filename has @ and . replaced with _
$regexPattern = [regex]::Escape($ZipNamePattern) -replace '\\\{email\\\}', '(.+)'
$regexPattern = "^$regexPattern$"

Write-Output "ZIP name pattern: $ZipNamePattern"
Write-Output "Regex pattern: $regexPattern"
Write-Output ""

# Track results
$extractedPsts = @()
$unmatchedZips = @()
$failedExtractions = @()

foreach ($zip in $zipFiles)
{
    Write-Output "Processing: $($zip.Name)"

    # Try to extract email from filename using pattern
    $match = [regex]::Match($zip.BaseName + ".zip", $regexPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if (-not $match.Success)
    {
        Write-Warning "  - Does not match pattern, skipping"
        $unmatchedZips += $zip.Name
        continue
    }

    # Get the captured email (with _ substitutions)
    $emailEncoded = $match.Groups[1].Value

    # Convert back to email format: _ -> @ for first occurrence after username, remaining _ -> .
    # Strategy: The email format is user_domain_com, we need user@domain.com
    # Find the pattern: word_word_word where we replace appropriately

    # More robust approach: try to match against our source mapping
    $sourceEmail = $null
    foreach ($source in $sourceToTarget.Keys)
    {
        # Convert source email to encoded format for comparison
        $sourceEncoded = $source -replace '@', '_' -replace '\.', '_'
        if ($sourceEncoded -eq $emailEncoded.ToLower())
        {
            $sourceEmail = $source
            break
        }
    }

    if (-not $sourceEmail)
    {
        Write-Warning "  - Could not match to any source in mapping CSV"
        $unmatchedZips += $zip.Name
        continue
    }

    $targetEmail = $sourceToTarget[$sourceEmail]
    Write-Output "  - Source: $sourceEmail"
    Write-Output "  - Target: $targetEmail"

    # Create subfolder for this mailbox (use source email for folder name)
    $extractFolder = Join-Path $OutputPath $sourceEmail
    if (-not (Test-Path $extractFolder))
    {
        New-Item -ItemType Directory -Path $extractFolder -Force | Out-Null
    }

    # Extract ZIP
    try
    {
        Expand-Archive -Path $zip.FullName -DestinationPath $extractFolder -Force
        Write-Output "  - Extracted to: $extractFolder"

        # Verify exchange.pst exists
        $pstPath = Join-Path $extractFolder "exchange.pst"
        if (Test-Path $pstPath)
        {
            $pstFile = Get-Item $pstPath
            $sizeGB = [math]::Round($pstFile.Length / 1GB, 2)
            Write-Output "  - Found exchange.pst ($sizeGB GB)"

            $extractedPsts += [PSCustomObject]@{
                SourceEmail  = $sourceEmail
                TargetEmail  = $targetEmail
                PstPath      = $pstPath
                FolderName   = $sourceEmail
                SizeGB       = $sizeGB
            }
        }
        else
        {
            # Check if there's any PST file with different name
            $anyPst = Get-ChildItem -Path $extractFolder -Filter "*.pst" | Select-Object -First 1
            if ($anyPst)
            {
                # Rename to exchange.pst for consistency
                $newPath = Join-Path $extractFolder "exchange.pst"
                Rename-Item -Path $anyPst.FullName -NewName "exchange.pst"
                Write-Output "  - Renamed $($anyPst.Name) to exchange.pst"

                $pstFile = Get-Item $newPath
                $sizeGB = [math]::Round($pstFile.Length / 1GB, 2)

                $extractedPsts += [PSCustomObject]@{
                    SourceEmail  = $sourceEmail
                    TargetEmail  = $targetEmail
                    PstPath      = $newPath
                    FolderName   = $sourceEmail
                    SizeGB       = $sizeGB
                }
            }
            else
            {
                Write-Warning "  - No PST file found in ZIP"
                $failedExtractions += $zip.Name
            }
        }
    }
    catch
    {
        Write-Warning "  - Extraction failed: $_"
        $failedExtractions += $zip.Name
    }

    Write-Output ""
}

Write-Output "Extraction complete."
Write-Output "  - Successfully extracted: $($extractedPsts.Count)"
Write-Output "  - Unmatched ZIPs: $($unmatchedZips.Count)"
Write-Output "  - Failed extractions: $($failedExtractions.Count)"
Write-Output ""
#endregion

#region STEP 2: GENERATE PURVIEW MAPPING CSV
Write-Output "Step 2: Generating Purview Import Mapping CSV..."
Write-Output "================================================"
Write-Output ""

if ($extractedPsts.Count -eq 0)
{
    Write-Error "No PSTs were successfully extracted. Cannot generate mapping."
    exit 1
}

$purviewMapping = @()

foreach ($pst in $extractedPsts)
{
    $purviewMapping += [PSCustomObject]@{
        Workload            = "Exchange"
        FilePath            = "$($pst.FolderName)/exchange.pst"
        Name                = "exchange.pst"
        Mailbox             = $pst.TargetEmail
        IsArchive           = "TRUE"
        TargetRootFolder    = "/"
        SPFileContainer     = ""
        SPManifestContainer = ""
        SPSiteUrl           = ""
    }
}

# Generate output filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$purviewCsvPath = Join-Path $OutputPath "PurviewMapping_$timestamp.csv"

# Export mapping CSV
$purviewMapping | Export-Csv -Path $purviewCsvPath -NoTypeInformation

Write-Output "Generated: $purviewCsvPath"
Write-Output ""
Write-Output "Mapping preview:"
$purviewMapping | Select-Object FilePath, Mailbox, IsArchive | Format-Table -AutoSize
Write-Output ""
#endregion

#region STEP 3: OUTPUT SUMMARY AND AZCOPY COMMAND
Write-Output "============================================================"
Write-Output "PREPARATION COMPLETE"
Write-Output "============================================================"
Write-Output ""

# Calculate total size
$totalSizeGB = ($extractedPsts | Measure-Object -Property SizeGB -Sum).Sum
Write-Output "Summary:"
Write-Output "  - PSTs prepared: $($extractedPsts.Count)"
Write-Output "  - Total size: $([math]::Round($totalSizeGB, 2)) GB"
Write-Output "  - Output folder: $OutputPath"
Write-Output "  - Mapping CSV: $purviewCsvPath"
Write-Output ""

if ($unmatchedZips.Count -gt 0)
{
    Write-Warning "Unmatched ZIP files (not in mapping CSV):"
    foreach ($unmatched in $unmatchedZips)
    {
        Write-Warning "  - $unmatched"
    }
    Write-Output ""
}

if ($failedExtractions.Count -gt 0)
{
    Write-Warning "Failed extractions:"
    foreach ($failed in $failedExtractions)
    {
        Write-Warning "  - $failed"
    }
    Write-Output ""
}

Write-Output "NEXT STEPS:"
Write-Output "-----------"
Write-Output ""
Write-Output "1. Get SAS URL from Purview portal:"
Write-Output "   - Navigate to: https://compliance.microsoft.com"
Write-Output "   - Go to: Data lifecycle management > Import"
Write-Output "   - Click: 'New import job' to get the SAS URL"
Write-Output ""

if ($AzureSasUrl)
{
    Write-Output "2. Run AzCopy to upload PSTs:"
    Write-Output ""
    Write-Output "   azcopy copy `"$OutputPath\*`" `"$AzureSasUrl`" --recursive"
    Write-Output ""
}
else
{
    Write-Output "2. Run AzCopy to upload PSTs:"
    Write-Output ""
    Write-Output "   azcopy copy `"$OutputPath\*`" `"<SAS_URL_FROM_PURVIEW>`" --recursive"
    Write-Output ""
    Write-Output "   (Re-run this script with -AzureSasUrl parameter for complete command)"
    Write-Output ""
}

Write-Output "3. Create import job in Purview:"
Write-Output "   - Upload the mapping CSV: $purviewCsvPath"
Write-Output "   - Select 'Import to archive mailbox'"
Write-Output "   - Start the import job"
Write-Output ""
Write-Output "4. Monitor progress in Purview portal"
Write-Output ""
Write-Output "Reference: https://learn.microsoft.com/en-us/purview/use-network-upload-to-import-pst-files"
Write-Output ""
#endregion
