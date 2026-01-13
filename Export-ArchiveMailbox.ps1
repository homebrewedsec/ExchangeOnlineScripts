<#
.SYNOPSIS
    Exports Exchange Online archive mailbox content to PST files using eDiscovery.

.DESCRIPTION
    This script automates the export of archive mailbox content from Exchange Online
    using the Microsoft Graph eDiscovery API. It creates targeted compliance searches
    that scope only to archive folders (excluding primary mailbox), exports to PST format,
    and downloads the files automatically.

    Key Features:
    - Targets ONLY archive mailbox content (not primary mailbox)
    - Processes largest archives first for optimal throughput
    - Supports both interactive and app-only authentication
    - Automated PST download (no manual portal interaction needed)
    - Individual PST per mailbox for easy management
    - Comprehensive progress reporting and error handling

.PARAMETER InputCsvPath
    Path to CSV file containing mailboxes to export.
    Required column: "upn" (user principal name)

.PARAMETER OutputPath
    Directory where PST files will be downloaded.
    Defaults to current directory.

.PARAMETER CaseName
    Prefix for the eDiscovery case name.
    Defaults to "ArchiveExport" with timestamp.

.PARAMETER AppId
    Azure AD application ID for app-only authentication.
    Optional - if not specified, uses interactive authentication.

.PARAMETER TenantId
    Azure AD tenant ID for app-only authentication.
    Required when using AppId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication.
    Required when using AppId with certificate auth.

.PARAMETER ClientSecret
    Client secret for app-only authentication.
    Alternative to CertificateThumbprint.
    App registration needs eDiscovery.ReadWrite.All (Graph).

.PARAMETER MaxBatchSizeGB
    Maximum total archive size per eDiscovery case in GB.
    When adding the next mailbox would exceed this limit, a new case is created.
    Default is 0 (unlimited - all mailboxes in one case).
    Recommended: 500-1000 GB per case for large migrations.

.PARAMETER SkipDownload
    Create searches and exports only, skip PST download.
    Useful for testing or when downloading via portal.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv"
    Exports archives using interactive authentication.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -OutputPath "C:\Exports"
    Exports archives to specific output directory.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -AppId "app-id" -TenantId "tenant-id" -CertificateThumbprint "thumbprint"
    Exports using certificate-based app authentication.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -AppId "app-id" -TenantId "tenant-id" -ClientSecret "secret"
    Exports using client secret authentication with automated PST downloads.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -SkipDownload
    Creates eDiscovery searches and exports without downloading (download via portal).

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -MaxBatchSizeGB 750
    Splits mailboxes into batches of ~750GB each, creating a separate eDiscovery case per batch.

.NOTES
    Author: Hudson Bush, Seguri - hudson@seguri.io
    Requires: ExchangeOnlineManagement, Microsoft.Graph modules
    Roles Required: eDiscovery Manager (Purview), Exchange Administrator

    Prerequisites:
    - For interactive auth: User with eDiscovery Manager role
    - For app-only auth: Azure AD app registration with:
      - Microsoft Graph: eDiscovery.ReadWrite.All (Application)

    Reference Documentation:
    - https://learn.microsoft.com/en-us/purview/edisc-ref-api-guide
    - https://practical365.com/purview-ediscovery-powershell/
    - https://www.thinformatics.com/blog/export-archiv-mailbox-content-using-ediscovery
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [string]$OutputPath = (Get-Location).Path,

    [string]$CaseName = "ArchiveExport",

    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ClientSecret,

    [int]$MaxBatchSizeGB = 0,

    [switch]$SkipDownload
)

#region CONFIGURATION
$script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:FullCaseName = "${CaseName}_$script:Timestamp"
$script:LogFile = Join-Path $OutputPath "ArchiveExport_Log_$script:Timestamp.log"
$script:SummaryFile = Join-Path $OutputPath "ArchiveExport_Summary_$script:Timestamp.csv"

# System folders to exclude from archive search
$script:SystemFoldersToExclude = @(
    "Audits"
    "Calendar Logging"
    "Deletions"
    "DiscoveryHolds"
    "ExternalContacts"
    "Purges"
    "Recoverable Items"
    "SubstrateHolds"
    "Versions"
)

# Polling intervals
$script:SearchPollIntervalSeconds = 30
$script:ExportPollIntervalSeconds = 60
$script:MaxWaitMinutes = 180
#endregion

#region LOGGING FUNCTION
function Write-ExportLog
{
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to console (use Write-Host to avoid polluting pipeline/return values)
    switch ($Level)
    {
        "WARNING" { Write-Warning $Message }
        "ERROR" { Write-Error $Message }
        default { Write-Host $Message }
    }

    # Write to log file
    Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
}
#endregion

#region ARCHIVE FOLDER QUERY FUNCTION
function Get-ArchiveFolderQuery
{
    <#
    .SYNOPSIS
        Builds a compliance search query targeting archive folder IDs.
    .DESCRIPTION
        Enumerates archive mailbox folders, converts folder IDs to compliance search format,
        and builds a ContentMatchQuery string that targets only archive content.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxUPN
    )

    Write-ExportLog "Getting archive folder IDs for: $MailboxUPN"

    try
    {
        # Get archive folder statistics - use ResultSize Unlimited to get ALL folders
        $archiveFolders = Get-MailboxFolderStatistics -Identity $MailboxUPN -Archive -ResultSize Unlimited -ErrorAction Stop

        if (-not $archiveFolders -or $archiveFolders.Count -eq 0)
        {
            Write-ExportLog "No archive folders found for: $MailboxUPN" -Level WARNING
            return $null
        }

        Write-ExportLog "Total folders in archive: $($archiveFolders.Count)"

        # Filter out system folders
        $userFolders = $archiveFolders | Where-Object {
            $folderPath = $_.FolderPath
            $isSystemFolder = $false

            foreach ($sysFolder in $script:SystemFoldersToExclude)
            {
                if ($folderPath -like "*$sysFolder*")
                {
                    $isSystemFolder = $true
                    break
                }
            }

            -not $isSystemFolder
        }

        if (-not $userFolders -or $userFolders.Count -eq 0)
        {
            Write-ExportLog "No user folders found in archive for: $MailboxUPN" -Level WARNING
            return $null
        }

        Write-ExportLog "Found $($userFolders.Count) user folders in archive"

        # Build folder ID query
        $folderQueries = @()

        foreach ($folder in $userFolders)
        {
            $folderId = $folder.FolderId

            if (-not $folderId)
            {
                continue
            }

            try
            {
                # Convert Exchange folder ID to compliance search format
                # The folder ID needs to be processed to extract the searchable portion
                $folderIdBytes = [Convert]::FromBase64String($folderId)

                if ($folderIdBytes.Length -ge 48)
                {
                    # Extract the relevant portion (bytes 24-47 typically contain the folder identifier)
                    $searchableBytes = $folderIdBytes[23..46]
                    $hexString = ($searchableBytes | ForEach-Object { $_.ToString("x2") }) -join ""
                    $folderQueries += "folderid:$hexString"
                }
            }
            catch
            {
                # If conversion fails, skip this folder
                Write-ExportLog "Could not convert folder ID for: $($folder.FolderPath)" -Level WARNING
            }
        }

        if ($folderQueries.Count -eq 0)
        {
            Write-ExportLog "No valid folder IDs extracted for: $MailboxUPN" -Level WARNING
            return $null
        }

        Write-ExportLog "Built $($folderQueries.Count) folder ID queries"

        # eDiscovery has a query length limit (~16KB is safe, truncation happens around 64KB)
        # Split into multiple queries if needed
        $maxQueryLength = 16000
        $queries = @()
        $currentBatch = @()
        $currentLength = 2  # Start with 2 for opening/closing parentheses

        foreach ($fq in $folderQueries)
        {
            $addLength = $fq.Length + 4  # " OR " separator
            if (($currentLength + $addLength) -gt $maxQueryLength -and $currentBatch.Count -gt 0)
            {
                # Save current batch and start new one
                $queries += "(" + ($currentBatch -join " OR ") + ")"
                $currentBatch = @()
                $currentLength = 2
            }
            $currentBatch += $fq
            $currentLength += $addLength
        }

        # Add final batch
        if ($currentBatch.Count -gt 0)
        {
            $queries += "(" + ($currentBatch -join " OR ") + ")"
        }

        if ($queries.Count -eq 1)
        {
            Write-ExportLog "Query length: $($queries[0].Length) characters"
            return $queries[0]
        }
        else
        {
            Write-ExportLog "Large mailbox - split into $($queries.Count) queries (max $maxQueryLength chars each)"
            foreach ($i in 0..($queries.Count - 1))
            {
                Write-ExportLog "  Query $($i + 1): $($queries[$i].Length) characters"
            }
            return $queries
        }
    }
    catch
    {
        Write-ExportLog "Error getting archive folders for $MailboxUPN : $($_.Exception.Message)" -Level ERROR
        return $null
    }
}
#endregion

#region WAIT FUNCTIONS
function Wait-SearchCompletion
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,

        [Parameter(Mandatory = $true)]
        [string]$SearchId,

        [int]$TimeoutMinutes = $script:MaxWaitMinutes
    )

    $startTime = Get-Date
    $timeout = New-TimeSpan -Minutes $TimeoutMinutes

    while ((Get-Date) - $startTime -lt $timeout)
    {
        try
        {
            # Must use -ExpandProperty to get the status from lastEstimateStatisticsOperation
            $search = Get-MgSecurityCaseEdiscoveryCaseSearch `
                -EdiscoveryCaseId $CaseId `
                -EdiscoverySearchId $SearchId `
                -ExpandProperty "lastEstimateStatisticsOperation"

            $searchStatus = $search.LastEstimateStatisticsOperation.Status

            if ($searchStatus -eq "succeeded" -or $searchStatus -eq "completed")
            {
                return $true
            }
            elseif ($searchStatus -eq "failed")
            {
                Write-ExportLog "Search failed: $SearchId" -Level ERROR
                return $false
            }

            Write-Output "  Search status: $searchStatus - waiting..."
            Start-Sleep -Seconds $script:SearchPollIntervalSeconds
        }
        catch
        {
            Write-ExportLog "Error checking search status: $($_.Exception.Message)" -Level WARNING
            Start-Sleep -Seconds $script:SearchPollIntervalSeconds
        }
    }

    Write-ExportLog "Search timed out after $TimeoutMinutes minutes" -Level ERROR
    return $false
}

function Wait-ExportCompletion
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaseId,

        [Parameter(Mandatory = $true)]
        [string]$OperationId,

        [int]$TimeoutMinutes = $script:MaxWaitMinutes
    )

    $startTime = Get-Date
    $timeout = New-TimeSpan -Minutes $TimeoutMinutes

    while ((Get-Date) - $startTime -lt $timeout)
    {
        try
        {
            $operation = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $CaseId -CaseOperationId $OperationId

            if ($operation.Status -eq "succeeded")
            {
                return $operation
            }
            elseif ($operation.Status -eq "failed")
            {
                Write-ExportLog "Export failed: $OperationId" -Level ERROR
                return $null
            }

            $percentComplete = $operation.PercentProgress
            Write-Output "  Export progress: $percentComplete% - waiting..."
            Start-Sleep -Seconds $script:ExportPollIntervalSeconds
        }
        catch
        {
            Write-ExportLog "Error checking export status: $($_.Exception.Message)" -Level WARNING
            Start-Sleep -Seconds $script:ExportPollIntervalSeconds
        }
    }

    Write-ExportLog "Export timed out after $TimeoutMinutes minutes" -Level ERROR
    return $null
}
#endregion

#region MAIN SCRIPT

try
{
    Write-ExportLog "============================================================"
    Write-ExportLog "Exchange Online Archive Mailbox Export"
    Write-ExportLog "============================================================"
    Write-ExportLog "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-ExportLog "Case Name: $script:FullCaseName"
    Write-ExportLog "Output Path: $OutputPath"
    Write-ExportLog ""

    #region VALIDATE PREREQUISITES
    Write-ExportLog "Validating prerequisites..."

    # Check output directory
    if (-not (Test-Path $OutputPath))
    {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-ExportLog "Created output directory: $OutputPath"
    }

    # Check CSV file
    if (-not (Test-Path $InputCsvPath))
    {
        throw "CSV file not found: $InputCsvPath"
    }

    # Load CSV
    $csvData = Import-Csv $InputCsvPath
    if (-not $csvData -or $csvData.Count -eq 0)
    {
        throw "CSV file is empty: $InputCsvPath"
    }

    # Check for upn column
    if (-not ($csvData | Get-Member -Name "upn" -MemberType NoteProperty))
    {
        throw "CSV file must contain a 'upn' column"
    }

    $upnList = $csvData | Select-Object -ExpandProperty upn | Where-Object { $_ }
    Write-ExportLog "Found $($upnList.Count) mailboxes in CSV"

    # Check required modules
    $requiredModules = @("ExchangeOnlineManagement", "Microsoft.Graph")
    foreach ($module in $requiredModules)
    {
        if (-not (Get-Module -Name $module -ListAvailable))
        {
            throw "Required module not found: $module. Install with: Install-Module $module"
        }
    }

    Write-ExportLog "Prerequisites validated"
    Write-ExportLog ""
    #endregion

    #region CONNECT TO SERVICES
    Write-ExportLog "Connecting to services..."

    # Determine authentication mode
    $useCertAuth = $AppId -and $TenantId -and $CertificateThumbprint
    $useSecretAuth = $AppId -and $TenantId -and $ClientSecret
    $useAppAuth = $useCertAuth -or $useSecretAuth

    # Connect to Exchange Online
    $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoSession -or $exoSession.State -ne "Connected")
    {
        Write-ExportLog "Connecting to Exchange Online..."
        if ($useCertAuth)
        {
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization "$TenantId.onmicrosoft.com" -ShowBanner:$false
        }
        else
        {
            # Interactive or existing session for Exchange (client secret not supported for EXO)
            Connect-ExchangeOnline -ShowBanner:$false
        }
    }
    else
    {
        Write-ExportLog "Using existing Exchange Online session"
    }

    # Connect to Microsoft Graph
    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $graphContext)
    {
        Write-ExportLog "Connecting to Microsoft Graph..."
        if ($useSecretAuth)
        {
            # Client secret authentication
            $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential
        }
        elseif ($useCertAuth)
        {
            # Certificate authentication
            $cert = Get-ChildItem "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            if (-not $cert)
            {
                $cert = Get-ChildItem "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            }
            if (-not $cert)
            {
                throw "Certificate not found: $CertificateThumbprint"
            }
            Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert
        }
        else
        {
            # Interactive authentication
            Connect-MgGraph -Scopes "eDiscovery.ReadWrite.All"
        }
    }
    else
    {
        Write-ExportLog "Using existing Microsoft Graph session"
    }

    # Store credentials for download phase if using app auth
    $script:UseAppAuth = $useAppAuth
    $script:AppId = $AppId
    $script:TenantId = $TenantId
    $script:ClientSecret = $ClientSecret
    $script:CertificateThumbprint = $CertificateThumbprint

    Write-ExportLog "Connected to all services"
    Write-ExportLog ""
    #endregion

    #region ENUMERATE ARCHIVES AND GET SIZES
    Write-ExportLog "Enumerating archive mailboxes and sizes..."

    $mailboxInfo = @()
    $counter = 0

    foreach ($upn in $upnList)
    {
        $counter++
        Write-Progress -Activity "Enumerating Archives" -Status "Processing $upn" -PercentComplete (($counter / $upnList.Count) * 100)

        try
        {
            $mailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue

            if (-not $mailbox)
            {
                Write-ExportLog "Mailbox not found: $upn" -Level WARNING
                continue
            }

            if (-not $mailbox.ArchiveStatus -or $mailbox.ArchiveStatus -eq "None")
            {
                Write-ExportLog "No archive enabled for: $upn" -Level WARNING
                continue
            }

            $archiveStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction SilentlyContinue

            if (-not $archiveStats)
            {
                Write-ExportLog "Could not get archive stats for: $upn" -Level WARNING
                continue
            }

            # Parse size
            $sizeString = $archiveStats.TotalItemSize.ToString()
            $sizeBytes = 0
            if ($sizeString -match "\(([0-9,]+) bytes\)")
            {
                $sizeBytes = [long]($matches[1] -replace ",", "")
            }

            $mailboxInfo += [PSCustomObject]@{
                UPN            = $upn
                DisplayName    = $mailbox.DisplayName
                ArchiveSizeBytes = $sizeBytes
                ArchiveSizeGB  = [math]::Round($sizeBytes / 1GB, 2)
                ItemCount      = $archiveStats.ItemCount
            }
        }
        catch
        {
            Write-ExportLog "Error processing $upn : $($_.Exception.Message)" -Level WARNING
        }
    }

    Write-Progress -Activity "Enumerating Archives" -Completed

    if ($mailboxInfo.Count -eq 0)
    {
        throw "No valid archive mailboxes found to process"
    }

    # Sort by size descending (largest first)
    $mailboxInfo = $mailboxInfo | Sort-Object -Property ArchiveSizeBytes -Descending

    Write-ExportLog "Found $($mailboxInfo.Count) archive mailboxes to process"
    Write-ExportLog "Total archive size: $([math]::Round(($mailboxInfo | Measure-Object -Property ArchiveSizeBytes -Sum).Sum / 1GB, 2)) GB"
    Write-ExportLog "Processing order (largest first):"
    foreach ($mb in $mailboxInfo | Select-Object -First 5)
    {
        Write-ExportLog "  - $($mb.UPN): $($mb.ArchiveSizeGB) GB ($($mb.ItemCount) items)"
    }
    if ($mailboxInfo.Count -gt 5)
    {
        Write-ExportLog "  ... and $($mailboxInfo.Count - 5) more"
    }
    Write-ExportLog ""

    #region SPLIT INTO BATCHES
    # Split mailboxes into batches based on MaxBatchSizeGB
    $batches = @()
    if ($MaxBatchSizeGB -gt 0)
    {
        $maxBatchBytes = [long]$MaxBatchSizeGB * 1GB
        $currentBatch = @()
        $currentBatchSize = 0

        foreach ($mb in $mailboxInfo)
        {
            # If adding this mailbox would exceed the limit and we have items, start new batch
            if ($currentBatchSize + $mb.ArchiveSizeBytes -gt $maxBatchBytes -and $currentBatch.Count -gt 0)
            {
                $batches += ,@($currentBatch)
                $currentBatch = @()
                $currentBatchSize = 0
            }

            $currentBatch += $mb
            $currentBatchSize += $mb.ArchiveSizeBytes
        }

        # Add final batch
        if ($currentBatch.Count -gt 0)
        {
            $batches += ,@($currentBatch)
        }

        Write-ExportLog "Split into $($batches.Count) batches (max $MaxBatchSizeGB GB each):"
        for ($i = 0; $i -lt $batches.Count; $i++)
        {
            $batchSize = [math]::Round(($batches[$i] | Measure-Object -Property ArchiveSizeBytes -Sum).Sum / 1GB, 2)
            Write-ExportLog "  Batch $($i + 1): $($batches[$i].Count) mailboxes, $batchSize GB"
        }
        Write-ExportLog ""
    }
    else
    {
        # Single batch with all mailboxes
        $batches = @(,@($mailboxInfo))
    }
    #endregion

    # Track all results and cases across batches
    $allResults = @()
    $allCases = @()

    #region PROCESS EACH BATCH
    for ($batchNum = 0; $batchNum -lt $batches.Count; $batchNum++)
    {
        $batchMailboxes = $batches[$batchNum]
        $batchLabel = if ($batches.Count -gt 1) { "_Batch$($batchNum + 1)" } else { "" }

        Write-ExportLog "============================================================"
        Write-ExportLog "PROCESSING BATCH $($batchNum + 1) OF $($batches.Count)"
        Write-ExportLog "  Mailboxes: $($batchMailboxes.Count)"
        Write-ExportLog "  Size: $([math]::Round(($batchMailboxes | Measure-Object -Property ArchiveSizeBytes -Sum).Sum / 1GB, 2)) GB"
        Write-ExportLog "============================================================"
        Write-ExportLog ""

        #region CREATE EDISCOVERY CASE
        $batchCaseName = "$script:FullCaseName$batchLabel"
        Write-ExportLog "Creating eDiscovery case: $batchCaseName"

        $caseParams = @{
            displayName  = $batchCaseName
            description  = "Archive mailbox export for migration - Batch $($batchNum + 1) of $($batches.Count) - Created $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            externalId   = "ArchiveExport-$script:Timestamp$batchLabel"
        }

        $case = New-MgSecurityCaseEdiscoveryCase -BodyParameter $caseParams
        $caseId = $case.Id

        $allCases += [PSCustomObject]@{
            BatchNum = $batchNum + 1
            CaseName = $batchCaseName
            CaseId   = $caseId
            MailboxCount = $batchMailboxes.Count
            Url = "https://compliance.microsoft.com/contentsearchv2?caseId=$caseId&caseType=eDiscovery"
        }

        Write-ExportLog "Created eDiscovery case: $caseId"
        Write-ExportLog ""
        #endregion

        #region PHASE 1: CREATE ALL SEARCHES
        Write-ExportLog "Phase 1: Creating compliance searches..."
        Write-ExportLog ""

        $results = @()
        $searchCounter = 0

        foreach ($mb in $batchMailboxes)
        {
            $searchCounter++
            $upn = $mb.UPN

            Write-ExportLog "[$searchCounter/$($batchMailboxes.Count)] Creating search for: $upn ($($mb.ArchiveSizeGB) GB)"

            $result = [PSCustomObject]@{
                UPN              = $upn
                DisplayName      = $mb.DisplayName
                ArchiveSizeGB    = $mb.ArchiveSizeGB
                ItemCount        = $mb.ItemCount
                BatchNum         = $batchNum + 1
                CaseId           = $caseId
                CaseName         = $batchCaseName
                SearchId         = $null
                SearchName       = $null
                ExportId         = $null
                ExportUrl        = $null
                PstPath          = $null
                PstSizeGB        = $null
                Status           = "Pending"
                ErrorMessage     = $null
            }

        try
        {
            # Get archive folder query (may return array for large mailboxes)
            $contentQueries = Get-ArchiveFolderQuery -MailboxUPN $upn

            if (-not $contentQueries)
            {
                $result.Status = "Failed"
                $result.ErrorMessage = "Could not build folder query"
                $results += $result
                continue
            }

            # Ensure it's an array for consistent handling
            if ($contentQueries -is [string])
            {
                $contentQueries = @($contentQueries)
            }

            # Create search(es) - multiple if mailbox has many folders
            $searchIds = @()
            $searchPartNum = 0

            foreach ($queryString in $contentQueries)
            {
                $searchPartNum++
                $searchSuffix = if ($contentQueries.Count -gt 1) { "_Part$searchPartNum" } else { "" }
                $searchName = "Archive_$($upn -replace '@', '_at_')_$script:Timestamp$searchSuffix"

                Write-ExportLog "  Creating search$searchSuffix (query length: $($queryString.Length) chars)"

                # Use direct parameters instead of BodyParameter to avoid serialization issues
                $search = New-MgSecurityCaseEdiscoveryCaseSearch `
                    -EdiscoveryCaseId $caseId `
                    -DisplayName $searchName `
                    -ContentQuery $queryString `
                    -DataSourceScopes "allTenantMailboxes"

                Write-ExportLog "  Created search: $($search.Id)"

                # Add mailbox as data source
                $userSource = @{
                    "@odata.type" = "microsoft.graph.security.userSource"
                    email         = $upn
                }

                New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id -BodyParameter $userSource | Out-Null

                # Start search estimate (non-blocking)
                Invoke-MgEstimateSecurityCaseEdiscoveryCaseSearchStatistics -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id | Out-Null

                $searchIds += $search.Id
            }

            $result.SearchName = "Archive_$($upn -replace '@', '_at_')_$script:Timestamp"
            $result.SearchId = $searchIds -join ";"  # Store multiple IDs separated by semicolon
            $result.Status = "SearchStarted"
            Write-ExportLog "  $($searchIds.Count) search(es) started"
        }
        catch
        {
            $result.Status = "Failed"
            $result.ErrorMessage = $_.Exception.Message
            Write-ExportLog "  Error: $($_.Exception.Message)" -Level ERROR
        }

        $results += $result
    }

    Write-ExportLog ""
    Write-ExportLog "All $($results.Count) searches created."
    Write-ExportLog ""
    #endregion

    #region PHASE 2: MONITOR SEARCHES AND CREATE EXPORTS
    Write-ExportLog "Phase 2: Monitoring searches and creating exports as they complete..."
    Write-ExportLog ""

    # Give searches a moment to start
    Write-ExportLog "Waiting for searches to initialize..."
    Start-Sleep -Seconds 10

    $startTime = Get-Date
    $timeout = New-TimeSpan -Minutes $script:MaxWaitMinutes

    # Keep looping until all searches are done (ExportStarted, SearchFailed, or timeout)
    while ($true)
    {
        $pendingSearches = @($results | Where-Object { $_.Status -eq "SearchStarted" })

        if ($pendingSearches.Count -eq 0)
        {
            Write-ExportLog "All searches have completed or failed."
            break
        }

        if (((Get-Date) - $startTime) -gt $timeout)
        {
            Write-ExportLog "Timeout reached after $($script:MaxWaitMinutes) minutes" -Level WARNING
            break
        }

        foreach ($result in $pendingSearches)
        {
            try
            {
                # Handle multiple search IDs (for large mailboxes split into parts)
                $searchIds = $result.SearchId -split ";"
                $allSearchesComplete = $true
                $anySearchFailed = $false
                $searchStatuses = @()

                foreach ($searchId in $searchIds)
                {
                    # Get search details with expanded lastEstimateStatisticsOperation (required to get status)
                    $search = Get-MgSecurityCaseEdiscoveryCaseSearch `
                        -EdiscoveryCaseId $caseId `
                        -EdiscoverySearchId $searchId `
                        -ExpandProperty "lastEstimateStatisticsOperation"

                    $status = $search.LastEstimateStatisticsOperation.Status
                    $searchStatuses += $status

                    if ($status -eq "failed")
                    {
                        $anySearchFailed = $true
                    }
                    elseif ($status -ne "succeeded" -and $status -ne "completed")
                    {
                        $allSearchesComplete = $false
                    }
                }

                $statusSummary = ($searchStatuses | Select-Object -Unique) -join "/"
                Write-ExportLog "  $($result.UPN): status = $statusSummary ($($searchIds.Count) search(es))"

                if ($anySearchFailed)
                {
                    Write-ExportLog "Search failed for: $($result.UPN)" -Level ERROR
                    $result.Status = "SearchFailed"
                    $result.ErrorMessage = "One or more searches failed"
                }
                elseif ($allSearchesComplete)
                {
                    Write-ExportLog "All searches completed for: $($result.UPN)"

                    # Create exports for each search
                    $exportIds = @()
                    $partNum = 0
                    foreach ($searchId in $searchIds)
                    {
                        $partNum++
                        $exportSuffix = if ($searchIds.Count -gt 1) { "_Part$partNum" } else { "" }

                        $exportParams = @{
                            displayName       = "Export_$($result.SearchName)$exportSuffix"
                            exportCriteria    = "searchHits"
                            exportFormat      = "pst"
                            additionalOptions = "subfolderContents"
                            exportLocation    = "responsiveLocations"
                        }

                        Export-MgSecurityCaseEdiscoveryCaseSearchResult -EdiscoveryCaseId $caseId -EdiscoverySearchId $searchId -BodyParameter $exportParams | Out-Null

                        # Get export operation ID
                        Start-Sleep -Seconds 2
                        $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $caseId |
                            Where-Object { $_.Action -eq "exportResult" } |
                            Sort-Object -Property CreatedDateTime -Descending |
                            Select-Object -First 1

                        if ($operations)
                        {
                            $exportIds += $operations.Id
                            Write-ExportLog "  Export started: $($operations.Id)"
                        }
                    }

                    $result.ExportId = $exportIds -join ";"
                    # Build direct Purview portal URL for easy access
                    $result.ExportUrl = "https://compliance.microsoft.com/contentsearchv2?caseId=$caseId&caseType=eDiscovery"
                    $result.Status = "ExportStarted"
                    Write-ExportLog "  Purview URL: $($result.ExportUrl)"
                }
                # If not all complete and none failed - keep waiting (status is "running", "notStarted", etc.)
            }
            catch
            {
                Write-ExportLog "Error checking search for $($result.UPN): $($_.Exception.Message)" -Level WARNING
            }
        }

        # Re-check pending count
        $stillPending = @($results | Where-Object { $_.Status -eq "SearchStarted" }).Count
        $completedCount = @($results | Where-Object { $_.Status -eq "ExportStarted" }).Count

        if ($stillPending -gt 0)
        {
            Write-Output "  Status: $completedCount exports started, $stillPending searches still running - waiting $($script:SearchPollIntervalSeconds)s..."
            Start-Sleep -Seconds $script:SearchPollIntervalSeconds
        }
    }

    # Check for timeouts
    foreach ($result in $results | Where-Object { $_.Status -eq "SearchStarted" })
    {
        $result.Status = "SearchTimeout"
        $result.ErrorMessage = "Search timed out after $($script:MaxWaitMinutes) minutes"
    }

    Write-ExportLog ""
    Write-ExportLog "All searches processed. $(($results | Where-Object { $_.Status -eq 'ExportStarted' }).Count) exports started."
    Write-ExportLog ""
    #endregion

    #region WAIT FOR EXPORTS AND DOWNLOAD
    # Note: Automated download via API is not supported by Microsoft for Purview eDiscovery.
    # The download URLs require interactive user authentication.
    # Downloads must be done manually from the Purview portal.

    if (-not $SkipDownload)
    {
        Write-ExportLog "Waiting for exports to complete..."
        Write-ExportLog ""
        Write-ExportLog "NOTE: Automated download is not supported by Microsoft."
        Write-ExportLog "      Downloads must be done manually from the Purview portal."
        Write-ExportLog ""

        foreach ($result in $results | Where-Object { $_.ExportId })
        {
            Write-ExportLog "Waiting for export: $($result.UPN)..."

            try
            {
                # Handle multiple export IDs (for split searches)
                $exportIds = $result.ExportId -split ";"
                $allExportsComplete = $true

                foreach ($exportId in $exportIds)
                {
                    $operation = Wait-ExportCompletion -CaseId $caseId -OperationId $exportId

                    if (-not $operation)
                    {
                        $allExportsComplete = $false
                        Write-ExportLog "  Export $exportId did not complete" -Level WARNING
                    }
                    else
                    {
                        Write-ExportLog "  Export $exportId completed"
                    }
                }

                if ($allExportsComplete)
                {
                    $result.Status = "ReadyForDownload"
                    Write-ExportLog "  All exports ready for manual download"
                }
                else
                {
                    $result.Status = "ExportFailed"
                    $result.ErrorMessage = "One or more exports did not complete"
                }
            }
            catch
            {
                $result.Status = "ExportFailed"
                $result.ErrorMessage = $_.Exception.Message
                Write-ExportLog "  Export error: $($_.Exception.Message)" -Level ERROR
            }
        }
    }
    else
    {
        Write-ExportLog "SkipDownload specified - PST files available in Purview portal"
        foreach ($result in $results | Where-Object { $_.ExportId })
        {
            $result.Status = "ExportCreated"
        }
    }
    #endregion

        # Add batch results to all results
        $allResults += $results

        Write-ExportLog ""
        Write-ExportLog "Batch $($batchNum + 1) complete."
        Write-ExportLog ""
    }
    #endregion PROCESS EACH BATCH

    #region GENERATE SUMMARY REPORT
    Write-ExportLog ""
    Write-ExportLog "Generating summary report..."

    $allResults | Export-Csv -Path $script:SummaryFile -NoTypeInformation

    Write-ExportLog ""
    Write-ExportLog "============================================================"
    Write-ExportLog "EXPORT COMPLETE"
    Write-ExportLog "============================================================"
    Write-ExportLog ""
    $totalCount = @($allResults).Count
    $readyCount = @($allResults | Where-Object { $_.Status -eq 'ReadyForDownload' -or $_.Status -eq 'ExportCreated' -or $_.Status -eq 'ExportStarted' }).Count
    $failedCount = @($allResults | Where-Object { $_.Status -like '*Failed*' -or $_.Status -like '*Timeout*' }).Count
    $pendingCount = @($allResults | Where-Object { $_.Status -eq 'Pending' -or $_.Status -eq 'SearchStarted' }).Count

    Write-ExportLog "Summary:"
    Write-ExportLog "  Total mailboxes: $totalCount"
    Write-ExportLog "  Total batches/cases: $($allCases.Count)"
    Write-ExportLog "  Ready for download: $readyCount"
    Write-ExportLog "  Failed: $failedCount"
    Write-ExportLog "  Still pending: $pendingCount"
    Write-ExportLog ""
    Write-ExportLog "Output files:"
    Write-ExportLog "  Summary CSV: $script:SummaryFile"
    Write-ExportLog "  Log file: $script:LogFile"
    Write-ExportLog ""
    Write-ExportLog "============================================================"
    Write-ExportLog "DOWNLOAD EXPORTS FROM PURVIEW"
    Write-ExportLog "============================================================"
    Write-ExportLog ""

    # List all cases with their download links
    foreach ($caseInfo in $allCases)
    {
        Write-ExportLog "Case $($caseInfo.BatchNum): $($caseInfo.CaseName)"
        Write-ExportLog "  Mailboxes: $($caseInfo.MailboxCount)"
        Write-ExportLog "  Direct link: $($caseInfo.Url)"
        Write-ExportLog ""
    }

    # List all exports ready for download grouped by case
    $readyExports = @($allResults | Where-Object { $_.ExportId })
    if ($readyExports.Count -gt 0)
    {
        Write-ExportLog "Exports ready for download ($($readyExports.Count) total):"
        Write-ExportLog ""

        foreach ($caseInfo in $allCases)
        {
            $caseExports = @($readyExports | Where-Object { $_.CaseId -eq $caseInfo.CaseId })
            if ($caseExports.Count -gt 0)
            {
                Write-ExportLog "  --- $($caseInfo.CaseName) ---"
                foreach ($export in $caseExports)
                {
                    $exportCount = ($export.ExportId -split ";").Count
                    $exportLabel = if ($exportCount -gt 1) { "($exportCount parts)" } else { "" }
                    Write-ExportLog "  $($export.UPN) $exportLabel - $($export.ArchiveSizeGB) GB"
                }
                Write-ExportLog ""
            }
        }

        Write-ExportLog "To download: Open each case link above, click each export, then 'Download results'"
    }
    Write-ExportLog ""
    #endregion
}
catch
{
    Write-ExportLog "Script execution failed: $($_.Exception.Message)" -Level ERROR
    Write-ExportLog "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    exit 1
}
finally
{
    Write-ExportLog "Script completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

#endregion
