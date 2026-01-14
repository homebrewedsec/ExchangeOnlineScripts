<#
.SYNOPSIS
    Extracts ZIP files containing PSTs and generates Purview import mapping CSV.

.DESCRIPTION
    This script prepares PST files for import to Exchange Online via Microsoft Purview:
    1. Extracts ZIP files to a batch subfolder
    2. Matches extracted PSTs to target mailboxes from mapping CSV
    3. Generates Purview-compatible import mapping CSV
    4. Displays AzCopy command for uploading to Azure

    Expected ZIP naming format:
    PSTs.001.Export_Archive_{email}_{timestamp}.zip
    Where email has @ replaced with _at_ and . replaced with _
    Example: PSTs.001.Export_Archive_john_smith_at_domain_com_20260113_171458.zip

    IMPORT PROCESS (after running this script):
    1. Run the AzCopy command displayed to upload PSTs to Azure
    2. Navigate to https://compliance.microsoft.com
    3. Go to: Data lifecycle management > Import
    4. Create new import job and upload the generated mapping CSV
    5. Start the import job

.PARAMETER ZipFolderPath
    Folder containing ZIP files to extract.

.PARAMETER MappingCsvPath
    CSV file with source-to-target mailbox mapping.
    Required columns: SourceEmail (or SourceUPN), TargetEmail (or TargetUPN)

.PARAMETER OutputPath
    Directory where PSTs will be extracted to batch subfolder.
    Default: Current directory

.PARAMETER BatchName
    Name of the batch subfolder for organizing PSTs.
    Default: "Batch1"
    This becomes the FilePath value in the Purview mapping.

.PARAMETER AzureSasUrl
    Optional. Azure Blob Storage SAS URL from Purview portal.
    If provided, displays the complete AzCopy command to run.

.PARAMETER RunUpload
    Switch. If specified along with AzureSasUrl, executes AzCopy to upload PSTs.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv"

    Extracts ZIPs to Batch1 subfolder and generates Purview mapping CSV.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv" -BatchName "Batch2"

    Uses custom batch folder name.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv" -AzureSasUrl "https://..."

    Includes complete AzCopy command in output.

.EXAMPLE
    .\Import-ArchiveMailbox.ps1 -ZipFolderPath "C:\ZIPs" -MappingCsvPath "mapping.csv" -AzureSasUrl "https://..." -RunUpload

    Extracts, generates mapping, and uploads to Azure in one step.

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
    [string]$BatchName = "Batch1",

    [Parameter(Mandatory = $false)]
    [string]$AzureSasUrl,

    [Parameter(Mandatory = $false)]
    [switch]$RunUpload
)

#region VALIDATION
Write-Output "Validating parameters..."

# Validate RunUpload requires AzureSasUrl
if ($RunUpload -and -not $AzureSasUrl)
{
    Write-Error "-RunUpload requires -AzureSasUrl to be specified"
    exit 1
}

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

# Create batch subfolder
$batchFolder = Join-Path $OutputPath $BatchName
if (-not (Test-Path $batchFolder))
{
    Write-Output "Creating batch folder: $batchFolder"
    New-Item -ItemType Directory -Path $batchFolder -Force | Out-Null
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
# Key is encoded email format (lowercase): john_smith_at_domain_com
$sourceToTarget = @{}
$sourceToOriginal = @{}
foreach ($row in $mappingCsv)
{
    $source = $row.$sourceColumn.Trim()
    $target = $row.$targetColumn.Trim()

    # Create encoded version for matching: @ -> _at_, . -> _
    $sourceEncoded = $source.ToLower() -replace '@', '_at_' -replace '\.', '_'

    $sourceToTarget[$sourceEncoded] = $target
    $sourceToOriginal[$sourceEncoded] = $source
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
Write-Output "Batch folder: $batchFolder"
Write-Output ""

# Regex to extract email from filename
# Pattern: PSTs.001.Export_Archive_{email}_{timestamp}.zip
# or similar patterns - we look for _at_ to find the email portion
$emailPattern = '(?i)Export_Archive_(.+_at_.+?)_\d{8}_\d{6}'

# Track results
$extractedPsts = @()
$unmatchedZips = @()
$failedExtractions = @()

foreach ($zip in $zipFiles)
{
    Write-Output "Processing: $($zip.Name)"

    # Try to extract encoded email from filename
    $match = [regex]::Match($zip.BaseName, $emailPattern)

    $emailEncoded = $null
    if ($match.Success)
    {
        $emailEncoded = $match.Groups[1].Value.ToLower()
    }
    else
    {
        # Fallback: try to find _at_ pattern anywhere in filename
        $atMatch = [regex]::Match($zip.BaseName, '([a-z0-9_]+_at_[a-z0-9_]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($atMatch.Success)
        {
            $emailEncoded = $atMatch.Groups[1].Value.ToLower()
        }
    }

    if (-not $emailEncoded)
    {
        Write-Warning "  - Could not extract email from filename"
        $unmatchedZips += $zip.Name
        continue
    }

    Write-Output "  - Encoded email: $emailEncoded"

    # Look up in our mapping
    if (-not $sourceToTarget.ContainsKey($emailEncoded))
    {
        Write-Warning "  - Not found in mapping CSV"
        $unmatchedZips += $zip.Name
        continue
    }

    $sourceEmail = $sourceToOriginal[$emailEncoded]
    $targetEmail = $sourceToTarget[$emailEncoded]
    Write-Output "  - Source: $sourceEmail"
    Write-Output "  - Target: $targetEmail"

    # Extract ZIP to a temp location first, then rename and move PST
    try
    {
        # Create temp extraction folder
        $tempExtractPath = Join-Path $batchFolder "_temp_extract"
        if (Test-Path $tempExtractPath)
        {
            Remove-Item $tempExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null

        # Extract ZIP
        Expand-Archive -Path $zip.FullName -DestinationPath $tempExtractPath -Force

        # Look for exchange.pst (or any .pst) in the extracted content
        $extractedPst = Get-ChildItem -Path $tempExtractPath -Filter "*.pst" -Recurse | Select-Object -First 1

        if ($extractedPst)
        {
            # Target PST name matches the ZIP filename
            $targetPstName = $zip.BaseName + ".pst"
            $targetPstPath = Join-Path $batchFolder $targetPstName

            # Move and rename the PST
            Move-Item -Path $extractedPst.FullName -Destination $targetPstPath -Force

            $pstFile = Get-Item $targetPstPath
            $sizeGB = [math]::Round($pstFile.Length / 1GB, 2)
            Write-Output "  - Extracted and renamed: $($extractedPst.Name) -> $targetPstName ($sizeGB GB)"

            $extractedPsts += [PSCustomObject]@{
                SourceEmail = $sourceEmail
                TargetEmail = $targetEmail
                PstName     = $targetPstName
                PstPath     = $targetPstPath
                SizeGB      = $sizeGB
            }
        }
        else
        {
            Write-Warning "  - No PST file found in ZIP"
            $failedExtractions += $zip.Name
        }

        # Clean up temp folder
        if (Test-Path $tempExtractPath)
        {
            Remove-Item $tempExtractPath -Recurse -Force
        }
    }
    catch
    {
        Write-Warning "  - Extraction failed: $_"
        $failedExtractions += $zip.Name

        # Clean up temp folder on error
        if (Test-Path $tempExtractPath)
        {
            Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
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
        FilePath            = $BatchName
        Name                = $pst.PstName
        Mailbox             = $pst.TargetEmail
        IsArchive           = "TRUE"
        TargetRootFolder    = "/"
        ContentCodePage     = ""
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
$purviewMapping | Select-Object FilePath, Name, Mailbox, IsArchive | Format-Table -AutoSize
Write-Output ""
#endregion

#region STEP 3: UPLOAD TO AZURE (OPTIONAL)
if ($RunUpload -and $AzureSasUrl)
{
    Write-Output "Step 3: Uploading PSTs to Azure..."
    Write-Output "==================================="
    Write-Output ""

    # Verify AzCopy is available
    if (-not $azcopyPath)
    {
        Write-Error "AzCopy is required for upload but was not found in PATH"
        Write-Error "Download from: https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10"
        exit 1
    }

    Write-Output "Running: azcopy.exe copy `"$batchFolder`" `"<SAS_URL>`" --recursive=true"
    Write-Output ""

    try
    {
        # Execute AzCopy
        $azcopyResult = & azcopy.exe copy "$batchFolder" "$AzureSasUrl" --recursive=true 2>&1

        # Output the result
        $azcopyResult | ForEach-Object { Write-Output $_ }

        if ($LASTEXITCODE -eq 0)
        {
            Write-Output ""
            Write-Output "Upload completed successfully!"
        }
        else
        {
            Write-Warning "AzCopy completed with exit code: $LASTEXITCODE"
        }
    }
    catch
    {
        Write-Error "AzCopy execution failed: $_"
        exit 1
    }

    Write-Output ""
}
#endregion

#region STEP 4: OUTPUT SUMMARY
Write-Output "============================================================"
Write-Output "PREPARATION COMPLETE"
Write-Output "============================================================"
Write-Output ""

# Calculate total size
$totalSizeGB = ($extractedPsts | Measure-Object -Property SizeGB -Sum).Sum
Write-Output "Summary:"
Write-Output "  - PSTs prepared: $($extractedPsts.Count)"
Write-Output "  - Total size: $([math]::Round($totalSizeGB, 2)) GB"
Write-Output "  - Batch folder: $batchFolder"
Write-Output "  - Mapping CSV: $purviewCsvPath"
if ($RunUpload -and $AzureSasUrl)
{
    Write-Output "  - Upload: Completed"
}
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

if ($RunUpload -and $AzureSasUrl)
{
    # Upload already done
    Write-Output "1. Create import job in Purview:"
    Write-Output "   - Navigate to: https://compliance.microsoft.com"
    Write-Output "   - Go to: Data lifecycle management > Import"
    Write-Output "   - Upload the mapping CSV: $purviewCsvPath"
    Write-Output "   - Select 'Import to archive mailbox'"
    Write-Output "   - Start the import job"
    Write-Output ""
    Write-Output "2. Monitor progress in Purview portal"
}
else
{
    Write-Output "1. Get SAS URL from Purview portal:"
    Write-Output "   - Navigate to: https://compliance.microsoft.com"
    Write-Output "   - Go to: Data lifecycle management > Import"
    Write-Output "   - Click: 'New import job' to get the SAS URL"
    Write-Output ""

    if ($AzureSasUrl)
    {
        Write-Output "2. Run AzCopy to upload PSTs:"
        Write-Output ""
        Write-Output "   azcopy.exe copy `"$batchFolder`" `"$AzureSasUrl`" --recursive=true"
        Write-Output ""
        Write-Output "   Or re-run this script with -RunUpload to execute automatically"
        Write-Output ""
    }
    else
    {
        Write-Output "2. Run AzCopy to upload PSTs:"
        Write-Output ""
        Write-Output "   azcopy.exe copy `"$batchFolder`" `"<SAS_URL_FROM_PURVIEW>`" --recursive=true"
        Write-Output ""
        Write-Output "   Or re-run with: -AzureSasUrl `"<SAS_URL>`" -RunUpload"
        Write-Output ""
    }

    Write-Output "3. Create import job in Purview:"
    Write-Output "   - Upload the mapping CSV: $purviewCsvPath"
    Write-Output "   - Select 'Import to archive mailbox'"
    Write-Output "   - Start the import job"
    Write-Output ""
    Write-Output "4. Monitor progress in Purview portal"
}
Write-Output ""
Write-Output "Reference: https://learn.microsoft.com/en-us/purview/use-network-upload-to-import-pst-files"
Write-Output ""
#endregion
