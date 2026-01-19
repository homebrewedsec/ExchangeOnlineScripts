<#
.SYNOPSIS
    Generates a Purview import mapping CSV by scanning Azure blob storage for uploaded PSTs.

.DESCRIPTION
    This script connects to Azure blob storage using a SAS URL, lists all PST files,
    matches them to target mailboxes using a source-to-target mapping CSV, and generates
    a Purview-compatible mapping CSV for the import job.

    Use this script when you need to regenerate the mapping file after losing the original,
    or when you've uploaded PSTs directly to blob storage without using the Import script.

.PARAMETER SasUrl
    Full Azure blob container URL including SAS token from Purview import job.
    Example: "https://storageaccount.blob.core.windows.net/container?sv=2022-11-02&ss=b..."

.PARAMETER MappingCsvPath
    Optional. Path to the source-to-target email mapping CSV.
    If not provided, emails are extracted from PST filenames and used as both source and target.
    Required columns if provided: SourceEmail (or SourceUPN) and TargetEmail (or TargetUPN).

.PARAMETER FilePath
    The FilePath value for the Purview mapping (folder within container).
    Defaults to the container name.

.PARAMETER OutputPath
    Directory for the generated mapping CSV. Defaults to current directory.

.PARAMETER AppendToExisting
    Path to an existing Purview mapping CSV to append to instead of creating new.

.PARAMETER BlobPrefix
    Optional prefix/folder path within the container to filter PSTs.
    Example: "Batch1/" to only scan PSTs in the Batch1 folder.

.EXAMPLE
    .\New-PurviewMappingFromBlob.ps1 -SasUrl "https://storage.blob.core.windows.net/psts?sv=..." -MappingCsvPath ".\mapping.csv"
    Scans container and generates Purview mapping.

.EXAMPLE
    .\New-PurviewMappingFromBlob.ps1 -SasUrl "https://..." -MappingCsvPath ".\mapping.csv" -BlobPrefix "Batch2/"
    Scans only the Batch2 folder within the container.

.NOTES
    Author: Hudson Bush / Claude AI
    Version: 1.1

    Get the SAS URL from the Purview import job page.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SasUrl,

    [Parameter(Mandatory = $false)]
    [string]$MappingCsvPath,

    [string]$FilePath,

    [string]$OutputPath = (Get-Location).Path,

    [string]$AppendToExisting,

    [string]$BlobPrefix
)

#region CONFIGURATION
$script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:OutputCsvPath = Join-Path $OutputPath "PurviewMapping_FromBlob_$script:Timestamp.csv"
$script:UnmatchedCsvPath = Join-Path $OutputPath "UnmatchedBlobs_$script:Timestamp.csv"
#endregion

#region VALIDATION
Write-Host "Using SAS URL to connect to blob storage"

# Validate mapping CSV if provided
$useMapping = $false
if ($MappingCsvPath)
{
    if (-not (Test-Path $MappingCsvPath))
    {
        throw "Mapping CSV not found: $MappingCsvPath"
    }
    $useMapping = $true
}
else
{
    Write-Host "No mapping CSV provided - will extract emails from filenames and use as target"
}

# Validate append target if specified
if ($AppendToExisting -and -not (Test-Path $AppendToExisting))
{
    throw "Append target not found: $AppendToExisting"
}

# Ensure output directory exists
if (-not (Test-Path $OutputPath))
{
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
#endregion

#region LOAD MAPPING CSV (OPTIONAL)
$emailLookup = @{}

if ($useMapping)
{
    Write-Host "Loading source-to-target mapping: $MappingCsvPath"
    $mappingData = Import-Csv -Path $MappingCsvPath

    # Detect column names (flexible naming)
    $sourceCol = $null
    $targetCol = $null

    $possibleSourceCols = @("SourceEmail", "SourceUPN", "sourceEmail", "sourceUPN", "Source")
    $possibleTargetCols = @("TargetEmail", "TargetUPN", "targetEmail", "targetUPN", "Target")

    foreach ($col in $possibleSourceCols)
    {
        if ($mappingData | Get-Member -Name $col -MemberType NoteProperty)
        {
            $sourceCol = $col
            break
        }
    }

    foreach ($col in $possibleTargetCols)
    {
        if ($mappingData | Get-Member -Name $col -MemberType NoteProperty)
        {
            $targetCol = $col
            break
        }
    }

    if (-not $sourceCol -or -not $targetCol)
    {
        throw "Mapping CSV must have source column ($($possibleSourceCols -join ', ')) and target column ($($possibleTargetCols -join ', '))"
    }

    Write-Host "  Source column: $sourceCol"
    Write-Host "  Target column: $targetCol"
    Write-Host "  Mappings loaded: $($mappingData.Count)"

    # Build lookup hashtable (encoded email -> target email)
    foreach ($row in $mappingData)
    {
        $sourceEmail = $row.$sourceCol
        $targetEmail = $row.$targetCol

        if ($sourceEmail -and $targetEmail)
        {
            # Store both original and encoded versions
            $emailLookup[$sourceEmail.ToLower()] = $targetEmail
            $encoded = ($sourceEmail -replace "@", "_at_" -replace "\.", "_").ToLower()
            $emailLookup[$encoded] = $targetEmail
        }
    }
}
#endregion

#region LIST BLOBS
Write-Host ""
Write-Host "Scanning blob storage for PST files..."

$pstBlobs = @()

# Parse SAS URL
$uri = [System.Uri]$SasUrl
$baseUrl = "$($uri.Scheme)://$($uri.Host)$($uri.AbsolutePath)"
$sasToken = $uri.Query

# List blobs using REST API
$marker = $null
do
{
    $listUrl = "$baseUrl$sasToken&restype=container&comp=list"
    if ($BlobPrefix)
    {
        $listUrl += "&prefix=$([System.Web.HttpUtility]::UrlEncode($BlobPrefix))"
    }
    if ($marker)
    {
        $listUrl += "&marker=$marker"
    }

    try
    {
        $response = Invoke-RestMethod -Uri $listUrl -Method Get
        $blobs = $response.EnumerationResults.Blobs.Blob

        foreach ($blob in $blobs)
        {
            if ($blob.Name -like "*.pst")
            {
                $pstBlobs += [PSCustomObject]@{
                    Name = $blob.Name
                    Size = [long]$blob.Properties.'Content-Length'
                }
            }
        }

        $marker = $response.EnumerationResults.NextMarker
    }
    catch
    {
        throw "Failed to list blobs: $($_.Exception.Message)"
    }
} while ($marker)

Write-Host "  PST files found: $($pstBlobs.Count)"
#endregion

#region MATCH AND BUILD MAPPING
Write-Host ""
Write-Host "Matching PSTs to target mailboxes..."

$purviewMapping = @()
$unmatchedBlobs = @()
$matchedCount = 0

# Determine FilePath value for Purview mapping
if (-not $FilePath)
{
    if ($BlobPrefix)
    {
        # Use BlobPrefix as FilePath (trim trailing slash)
        $FilePath = $BlobPrefix.TrimEnd('/')
    }
    else
    {
        # Default to container name from URL
        $FilePath = $uri.AbsolutePath.TrimStart('/')
    }
}
Write-Host "FilePath for Purview mapping: $FilePath"

foreach ($blob in $pstBlobs)
{
    $blobName = $blob.Name
    $pstFileName = [System.IO.Path]::GetFileName($blobName)

    # Try to extract email from filename
    $targetEmail = $null
    $extractedEmail = $null

    # Pattern 1: Export_Archive_user_at_domain_com_*
    if ($pstFileName -match 'Export_Archive_(.+?)_\d{8}_\d{6}')
    {
        $extractedEmail = $Matches[1]
    }
    # Pattern 2: PSTs.001.Export_Archive_user_at_domain_com.*
    elseif ($pstFileName -match 'PSTs\.\d+\.Export_Archive_(.+?)\.pst')
    {
        $extractedEmail = $Matches[1]
    }
    # Pattern 3: Anything with _at_ pattern
    elseif ($pstFileName -match '([^_]+_at_[^_]+(?:_[^_]+)*)')
    {
        $extractedEmail = $Matches[1]
    }

    if ($extractedEmail)
    {
        if ($useMapping)
        {
            # Look up target email from mapping CSV
            $lookupKey = $extractedEmail.ToLower()
            if ($emailLookup.ContainsKey($lookupKey))
            {
                $targetEmail = $emailLookup[$lookupKey]
            }
        }
        else
        {
            # No mapping CSV - convert encoded email back to proper format and use as target
            # Pattern: user_at_domain_com -> user@domain.com
            $targetEmail = $extractedEmail -replace '_at_', '@'
            # Replace remaining underscores with dots (for domain parts)
            # But need to be careful - only replace underscores that are part of the domain
            if ($targetEmail -match '^([^@]+)@(.+)$')
            {
                $localPart = $Matches[1]
                $domainPart = $Matches[2] -replace '_', '.'
                $targetEmail = "$localPart@$domainPart"
            }
        }
    }

    if ($targetEmail)
    {
        $purviewMapping += [PSCustomObject]@{
            Workload            = "Exchange"
            FilePath            = $FilePath
            Name                = $pstFileName
            Mailbox             = $targetEmail
            IsArchive           = "FALSE"
            TargetRootFolder    = "/"
            ContentCodePage     = ""
            SPFileContainer     = ""
            SPManifestContainer = ""
            SPSiteUrl           = ""
        }
        $matchedCount++
        Write-Host "  MATCH: $pstFileName -> $targetEmail" -ForegroundColor Green
    }
    else
    {
        $unmatchedBlobs += [PSCustomObject]@{
            BlobName       = $blobName
            PstFileName    = $pstFileName
            ExtractedEmail = $extractedEmail
            Reason         = if ($extractedEmail) { "Not in mapping CSV" } else { "Could not extract email from filename" }
        }
        Write-Host "  UNMATCHED: $pstFileName" -ForegroundColor Yellow
    }
}
#endregion

#region EXPORT RESULTS
Write-Host ""

# Handle append mode
if ($AppendToExisting)
{
    Write-Host "Appending to existing mapping: $AppendToExisting"
    $existingMapping = Import-Csv -Path $AppendToExisting
    $purviewMapping = @($existingMapping) + @($purviewMapping)
    $script:OutputCsvPath = $AppendToExisting
}

# Export Purview mapping
if ($purviewMapping.Count -gt 0)
{
    $purviewMapping | Export-Csv -Path $script:OutputCsvPath -NoTypeInformation
    Write-Host "Purview mapping exported: $script:OutputCsvPath" -ForegroundColor Cyan
    Write-Host "  Total entries: $($purviewMapping.Count)"
}
else
{
    Write-Host "No PSTs matched. Purview mapping not created." -ForegroundColor Yellow
}

# Export unmatched blobs
if ($unmatchedBlobs.Count -gt 0)
{
    $unmatchedBlobs | Export-Csv -Path $script:UnmatchedCsvPath -NoTypeInformation
    Write-Host ""
    Write-Host "Unmatched blobs exported: $script:UnmatchedCsvPath" -ForegroundColor Yellow
    Write-Host "  Unmatched count: $($unmatchedBlobs.Count)"
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  PSTs in blob: $($pstBlobs.Count)"
Write-Host "  Matched: $matchedCount"
Write-Host "  Unmatched: $($unmatchedBlobs.Count)"
#endregion
