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
    Required when using AppId.

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
    Exports using app-only authentication for unattended operation.

.EXAMPLE
    .\Export-ArchiveMailbox.ps1 -InputCsvPath "mailboxes.csv" -SkipDownload
    Creates eDiscovery searches and exports without downloading (download via portal).

.NOTES
    Author: Hudson Bush, Seguri - hudson@seguri.io
    Requires: ExchangeOnlineManagement, Microsoft.Graph, MSAL.PS modules
    Roles Required: eDiscovery Manager (Purview), Exchange Administrator

    Prerequisites:
    - For interactive auth: User with eDiscovery Manager role
    - For app-only auth: Azure AD app registration with:
      - Microsoft Graph: eDiscovery.ReadWrite.All (Application)
      - MicrosoftPurviewEDiscovery: eDiscovery.Download.Read (Application)

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
        # Get archive folder statistics
        $archiveFolders = Get-MailboxFolderStatistics -Identity $MailboxUPN -Archive -ErrorAction Stop

        if (-not $archiveFolders -or $archiveFolders.Count -eq 0)
        {
            Write-ExportLog "No archive folders found for: $MailboxUPN" -Level WARNING
            return $null
        }

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

        # Join folder queries with OR - ensure it's a string
        [string]$contentQuery = "(" + ($folderQueries -join " OR ") + ")"

        Write-ExportLog "Built query with $($folderQueries.Count) folder IDs"
        Write-ExportLog "Query length: $($contentQuery.Length) characters"

        return $contentQuery
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
            $search = Get-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $CaseId -EdiscoverySearchId $SearchId

            # Status is in lastEstimateStatisticsOperation, not directly on search object
            $searchStatus = $null
            if ($search.LastEstimateStatisticsOperation)
            {
                $searchStatus = $search.LastEstimateStatisticsOperation.Status
            }

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

    if (-not $SkipDownload)
    {
        if (-not (Get-Module -Name "MSAL.PS" -ListAvailable))
        {
            throw "MSAL.PS module required for PST download. Install with: Install-Module MSAL.PS"
        }
    }

    Write-ExportLog "Prerequisites validated"
    Write-ExportLog ""
    #endregion

    #region CONNECT TO SERVICES
    Write-ExportLog "Connecting to services..."

    # Determine authentication mode
    $useAppAuth = $AppId -and $TenantId -and $CertificateThumbprint

    # Connect to Exchange Online
    $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoSession -or $exoSession.State -ne "Connected")
    {
        Write-ExportLog "Connecting to Exchange Online..."
        if ($useAppAuth)
        {
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization "$TenantId.onmicrosoft.com" -ShowBanner:$false
        }
        else
        {
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
        if ($useAppAuth)
        {
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
            Connect-MgGraph -Scopes "eDiscovery.ReadWrite.All"
        }
    }
    else
    {
        Write-ExportLog "Using existing Microsoft Graph session"
    }

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
    #endregion

    #region CREATE EDISCOVERY CASE
    Write-ExportLog "Creating eDiscovery case: $script:FullCaseName"

    $caseParams = @{
        displayName  = $script:FullCaseName
        description  = "Archive mailbox export for migration - Created $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        externalId   = "ArchiveExport-$script:Timestamp"
    }

    $case = New-MgSecurityCaseEdiscoveryCase -BodyParameter $caseParams
    $caseId = $case.Id

    Write-ExportLog "Created eDiscovery case: $caseId"
    Write-ExportLog ""
    #endregion

    #region PHASE 1: CREATE ALL SEARCHES
    Write-ExportLog "Phase 1: Creating compliance searches for all mailboxes..."
    Write-ExportLog ""

    $results = @()
    $searchCounter = 0

    foreach ($mb in $mailboxInfo)
    {
        $searchCounter++
        $upn = $mb.UPN

        Write-ExportLog "[$searchCounter/$($mailboxInfo.Count)] Creating search for: $upn ($($mb.ArchiveSizeGB) GB)"

        $result = [PSCustomObject]@{
            UPN              = $upn
            DisplayName      = $mb.DisplayName
            ArchiveSizeGB    = $mb.ArchiveSizeGB
            ItemCount        = $mb.ItemCount
            SearchId         = $null
            SearchName       = $null
            ExportId         = $null
            PstPath          = $null
            PstSizeGB        = $null
            Status           = "Pending"
            ErrorMessage     = $null
        }

        try
        {
            # Get archive folder query
            $contentQuery = Get-ArchiveFolderQuery -MailboxUPN $upn

            if (-not $contentQuery)
            {
                $result.Status = "Failed"
                $result.ErrorMessage = "Could not build folder query"
                $results += $result
                continue
            }

            # Create search
            $searchName = "Archive_$($upn -replace '@', '_at_')_$script:Timestamp"
            $result.SearchName = $searchName

            # Ensure contentQuery is a string
            [string]$queryString = $contentQuery
            Write-ExportLog "  Query type: $($queryString.GetType().Name), Length: $($queryString.Length)"

            # Use direct parameters instead of BodyParameter to avoid serialization issues
            $search = New-MgSecurityCaseEdiscoveryCaseSearch `
                -EdiscoveryCaseId $caseId `
                -DisplayName $searchName `
                -ContentQuery $queryString `
                -DataSourceScopes "allTenantMailboxes"
            $result.SearchId = $search.Id

            Write-ExportLog "  Created search: $($search.Id)"

            # Add mailbox as data source
            $userSource = @{
                "@odata.type" = "microsoft.graph.security.userSource"
                email         = $upn
            }

            New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id -BodyParameter $userSource | Out-Null

            # Start search estimate (non-blocking)
            Invoke-MgEstimateSecurityCaseEdiscoveryCaseSearchStatistics -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id | Out-Null

            $result.Status = "SearchStarted"
            Write-ExportLog "  Search started"
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
                # Get search details - status is in lastEstimateStatisticsOperation, not directly on search
                $search = Get-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $caseId -EdiscoverySearchId $result.SearchId

                # The status comes from the lastEstimateStatisticsOperation property
                $searchStatus = $null
                if ($search.LastEstimateStatisticsOperation)
                {
                    $searchStatus = $search.LastEstimateStatisticsOperation.Status
                }

                # If status still not available, query the operation directly
                if (-not $searchStatus)
                {
                    try
                    {
                        $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $caseId |
                            Where-Object { $_.Action -eq "estimateStatistics" } |
                            Sort-Object -Property CreatedDateTime -Descending |
                            Select-Object -First 1

                        if ($operations)
                        {
                            $searchStatus = $operations.Status
                        }
                    }
                    catch
                    {
                        # Operations may not be available yet - continue checking
                        Write-Verbose "Operations query failed, will retry: $($_.Exception.Message)"
                    }
                }

                Write-ExportLog "  $($result.UPN): status = $searchStatus"

                if ($searchStatus -eq "succeeded" -or $searchStatus -eq "completed")
                {
                    Write-ExportLog "Search completed for: $($result.UPN)"

                    # Create export
                    $exportParams = @{
                        displayName       = "Export_$($result.SearchName)"
                        exportCriteria    = "searchHits"
                        exportformats     = "pst"
                        additionalOptions = "subfolderContents"
                        exportLocation    = "responsiveLocations"
                    }

                    $null = Export-MgSecurityCaseEdiscoveryCaseSearchResult -EdiscoveryCaseId $caseId -EdiscoverySearchId $result.SearchId -BodyParameter $exportParams

                    # Get export operation ID
                    Start-Sleep -Seconds 2  # Brief pause to let operation register
                    $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $caseId |
                        Where-Object { $_.Action -eq "exportResult" } |
                        Sort-Object -Property CreatedDateTime -Descending |
                        Select-Object -First 1

                    if ($operations)
                    {
                        $result.ExportId = $operations.Id
                        Write-ExportLog "  Export started: $($operations.Id)"
                    }

                    $result.Status = "ExportStarted"
                }
                elseif ($searchStatus -eq "failed")
                {
                    Write-ExportLog "Search failed for: $($result.UPN)" -Level ERROR
                    $result.Status = "SearchFailed"
                    $result.ErrorMessage = "Search failed"
                }
                # If status is "running", "notStarted", etc. - keep waiting
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
    if (-not $SkipDownload)
    {
        Write-ExportLog "Downloading PST files..."

        # Get MSAL token for download
        $downloadToken = $null

        if ($useAppAuth)
        {
            Import-Module MSAL.PS

            $cert = Get-ChildItem "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            if (-not $cert)
            {
                $cert = Get-ChildItem "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            }

            $tokenParams = @{
                ClientId            = $AppId
                TenantId            = $TenantId
                ClientCertificate   = $cert
                Scopes              = "b26e684c-5068-4120-a679-64a5d2c909d9/.default"
            }
            $tokenResult = Get-MsalToken @tokenParams
            $downloadToken = $tokenResult.AccessToken
        }
        else
        {
            # Interactive mode - automated download requires app registration
            # Direct user to download from Purview portal instead
            Write-ExportLog ""
            Write-ExportLog "============================================================" -Level WARNING
            Write-ExportLog "INTERACTIVE MODE: Manual download required" -Level WARNING
            Write-ExportLog "============================================================" -Level WARNING
            Write-ExportLog "Automated PST download requires app registration with" -Level WARNING
            Write-ExportLog "eDiscovery.Download.Read permission." -Level WARNING
            Write-ExportLog "" -Level WARNING
            Write-ExportLog "Download your PST files from:" -Level WARNING
            Write-ExportLog "  https://compliance.microsoft.com/ediscovery" -Level WARNING
            Write-ExportLog "" -Level WARNING
            Write-ExportLog "Look for case: $script:FullCaseName" -Level WARNING
            Write-ExportLog "============================================================" -Level WARNING
            Write-ExportLog ""
            $SkipDownload = $true
        }

        # Check if download was disabled due to token failure
        if ($SkipDownload -or -not $downloadToken)
        {
            Write-ExportLog "Download skipped - exports available in Purview portal"
            foreach ($result in $results | Where-Object { $_.ExportId })
            {
                $result.Status = "ExportCreated"
            }
        }
        else
        {
        foreach ($result in $results | Where-Object { $_.ExportId })
        {
            Write-ExportLog "Waiting for export: $($result.UPN)..."

            try
            {
                $operation = Wait-ExportCompletion -CaseId $caseId -OperationId $result.ExportId

                if (-not $operation)
                {
                    $result.Status = "ExportFailed"
                    $result.ErrorMessage = "Export did not complete"
                    continue
                }

                # Get download URLs
                $uri = "/v1.0/security/cases/ediscoveryCases/$caseId/operations/$($result.ExportId)"
                $exportInfo = Invoke-MgGraphRequest -Uri $uri

                $fileMetadata = $exportInfo.exportFileMetadata

                if (-not $fileMetadata)
                {
                    $result.Status = "NoFiles"
                    $result.ErrorMessage = "No export files found"
                    continue
                }

                # Download each file
                foreach ($file in $fileMetadata)
                {
                    $fileName = $file.fileName
                    $downloadUrl = $file.downloadUrl
                    $outputFile = Join-Path $OutputPath "$($result.UPN -replace '@', '_at_')_$fileName"

                    Write-ExportLog "  Downloading: $fileName"

                    $headers = @{
                        "Authorization"      = "Bearer $downloadToken"
                        "X-AllowWithAADToken" = "true"
                    }

                    Invoke-WebRequest -Uri $downloadUrl -OutFile $outputFile -Headers $headers

                    if (Test-Path $outputFile)
                    {
                        $fileSize = (Get-Item $outputFile).Length
                        $result.PstPath = $outputFile
                        $result.PstSizeGB = [math]::Round($fileSize / 1GB, 2)
                        $result.Status = "Completed"
                        Write-ExportLog "  Downloaded: $outputFile ($($result.PstSizeGB) GB)"
                    }
                }
            }
            catch
            {
                $result.Status = "DownloadFailed"
                $result.ErrorMessage = $_.Exception.Message
                Write-ExportLog "  Download error: $($_.Exception.Message)" -Level ERROR
            }
        }
        } # End of download else block
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

    #region GENERATE SUMMARY REPORT
    Write-ExportLog ""
    Write-ExportLog "Generating summary report..."

    $results | Export-Csv -Path $script:SummaryFile -NoTypeInformation

    Write-ExportLog ""
    Write-ExportLog "============================================================"
    Write-ExportLog "EXPORT COMPLETE"
    Write-ExportLog "============================================================"
    Write-ExportLog ""
    $totalCount = @($results).Count
    $completedCount = @($results | Where-Object { $_.Status -eq 'Completed' -or $_.Status -eq 'ExportCreated' -or $_.Status -eq 'ExportStarted' }).Count
    $failedCount = @($results | Where-Object { $_.Status -like '*Failed*' -or $_.Status -like '*Timeout*' }).Count
    $pendingCount = @($results | Where-Object { $_.Status -eq 'Pending' -or $_.Status -eq 'SearchStarted' }).Count

    Write-ExportLog "Summary:"
    Write-ExportLog "  Total mailboxes: $totalCount"
    Write-ExportLog "  Exports ready: $completedCount"
    Write-ExportLog "  Failed: $failedCount"
    Write-ExportLog "  Still pending: $pendingCount"
    Write-ExportLog ""
    Write-ExportLog "Output files:"
    Write-ExportLog "  Summary: $script:SummaryFile"
    Write-ExportLog "  Log: $script:LogFile"
    Write-ExportLog "  PST files: $OutputPath"
    Write-ExportLog ""
    Write-ExportLog "eDiscovery Case: $script:FullCaseName (ID: $caseId)"
    Write-ExportLog "  View in Purview: https://compliance.microsoft.com/ediscovery"
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
