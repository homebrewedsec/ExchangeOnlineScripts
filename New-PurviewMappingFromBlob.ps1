<#
.SYNOPSIS
    Generates a Purview import mapping CSV by scanning Azure blob storage for uploaded PSTs.

.DESCRIPTION
    This script connects to Azure blob storage, lists all PST files, matches them to target
    mailboxes using a source-to-target mapping CSV, and generates a Purview-compatible
    mapping CSV for the import job.

    Use this script when you need to regenerate the mapping file after losing the original,
    or when you've uploaded PSTs directly to blob storage without using the Import script.

.PARAMETER ContainerUrl
    Full Azure blob container URL including SAS token.
    Example: "https://storageaccount.blob.core.windows.net/container?sv=2022-11-02&ss=b..."

.PARAMETER StorageAccountName
    Storage account name (for Az module authentication instead of SAS token).

.PARAMETER ContainerName
    Container name (for Az module authentication instead of SAS token).

.PARAMETER MappingCsvPath
    Path to the source-to-target email mapping CSV.
    Required columns: SourceEmail (or SourceUPN) and TargetEmail (or TargetUPN).

.PARAMETER FilePath
    The FilePath value for the Purview mapping (folder within container).
    Defaults to the container name or can be specified manually.

.PARAMETER OutputPath
    Directory for the generated mapping CSV. Defaults to current directory.

.PARAMETER AppendToExisting
    Path to an existing Purview mapping CSV to append to instead of creating new.

.PARAMETER BlobPrefix
    Optional prefix/folder path within the container to filter PSTs.
    Example: "Batch1/" to only scan PSTs in the Batch1 folder.

.EXAMPLE
    .\New-PurviewMappingFromBlob.ps1 -ContainerUrl "https://storage.blob.core.windows.net/psts?sv=..." -MappingCsvPath ".\mapping.csv"
    Scans container using SAS token and generates mapping.

.EXAMPLE
    .\New-PurviewMappingFromBlob.ps1 -StorageAccountName "mystorageaccount" -ContainerName "pstcontainer" -MappingCsvPath ".\mapping.csv"
    Scans container using Az module authentication.

.EXAMPLE
    .\New-PurviewMappingFromBlob.ps1 -ContainerUrl "https://..." -MappingCsvPath ".\mapping.csv" -BlobPrefix "Batch2/"
    Scans only the Batch2 folder within the container.

.NOTES
    Author: Hudson Bush / Claude AI
    Version: 1.0

    Prerequisites:
    - For SAS token: No modules required (uses REST API)
    - For Az module: Install-Module Az.Storage -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ContainerUrl,

    [Parameter(Mandatory = $false)]
    [string]$StorageAccountName,

    [Parameter(Mandatory = $false)]
    [string]$ContainerName,

    [Parameter(Mandatory = $true)]
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
# Validate authentication method
$useSasToken = $false

if ($ContainerUrl)
{
    $useSasToken = $true
    Write-Host "Using SAS token authentication"
}
elseif ($StorageAccountName -and $ContainerName)
{
    Write-Host "Using Az module authentication"

    # Verify Az.Storage module
    if (-not (Get-Module -ListAvailable -Name Az.Storage))
    {
        throw "Az.Storage module not found. Install with: Install-Module Az.Storage -Scope CurrentUser"
    }
}
else
{
    throw "Must specify either -ContainerUrl (SAS token) OR -StorageAccountName and -ContainerName (Az module)"
}

# Validate mapping CSV
if (-not (Test-Path $MappingCsvPath))
{
    throw "Mapping CSV not found: $MappingCsvPath"
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

#region LOAD MAPPING CSV
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
$emailLookup = @{}
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
#endregion

#region LIST BLOBS
Write-Host ""
Write-Host "Scanning blob storage for PST files..."

$pstBlobs = @()

if ($useSasToken)
{
    # Parse container URL
    $uri = [System.Uri]$ContainerUrl
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
}
else
{
    # Use Az module
    Import-Module Az.Storage -ErrorAction Stop

    $context = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount

    $listParams = @{
        Container = $ContainerName
        Context   = $context
    }
    if ($BlobPrefix)
    {
        $listParams.Prefix = $BlobPrefix
    }

    $blobs = Get-AzStorageBlob @listParams | Where-Object { $_.Name -like "*.pst" }

    foreach ($blob in $blobs)
    {
        $pstBlobs += [PSCustomObject]@{
            Name = $blob.Name
            Size = $blob.Length
        }
    }
}

Write-Host "  PST files found: $($pstBlobs.Count)"
#endregion

#region MATCH AND BUILD MAPPING
Write-Host ""
Write-Host "Matching PSTs to target mailboxes..."

$purviewMapping = @()
$unmatchedBlobs = @()
$matchedCount = 0

# Determine FilePath value
if (-not $FilePath)
{
    if ($useSasToken)
    {
        $uri = [System.Uri]$ContainerUrl
        $FilePath = $uri.AbsolutePath.TrimStart('/')
    }
    else
    {
        $FilePath = $ContainerName
    }
}

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
        $lookupKey = $extractedEmail.ToLower()
        if ($emailLookup.ContainsKey($lookupKey))
        {
            $targetEmail = $emailLookup[$lookupKey]
        }
    }

    if ($targetEmail)
    {
        $purviewMapping += [PSCustomObject]@{
            Workload            = "Exchange"
            FilePath            = $FilePath
            Name                = $pstFileName
            Mailbox             = $targetEmail
            IsArchive           = "TRUE"
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
