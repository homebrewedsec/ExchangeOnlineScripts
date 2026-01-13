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
function Write-Log
{
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to console
    switch ($Level)
    {
        "WARNING" { Write-Warning $Message }
        "ERROR" { Write-Error $Message }
        default { Write-Output $Message }
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

    Write-Log "Getting archive folder IDs for: $MailboxUPN"

    try
    {
        # Get archive folder statistics
        $archiveFolders = Get-MailboxFolderStatistics -Identity $MailboxUPN -Archive -ErrorAction Stop

        if (-not $archiveFolders -or $archiveFolders.Count -eq 0)
        {
            Write-Log "No archive folders found for: $MailboxUPN" -Level WARNING
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
            Write-Log "No user folders found in archive for: $MailboxUPN" -Level WARNING
            return $null
        }

        Write-Log "Found $($userFolders.Count) user folders in archive"

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
                Write-Log "Could not convert folder ID for: $($folder.FolderPath)" -Level WARNING
            }
        }

        if ($folderQueries.Count -eq 0)
        {
            Write-Log "No valid folder IDs extracted for: $MailboxUPN" -Level WARNING
            return $null
        }

        # Join folder queries with OR
        $contentQuery = "(" + ($folderQueries -join " OR ") + ")"

        Write-Log "Built query with $($folderQueries.Count) folder IDs"

        return $contentQuery
    }
    catch
    {
        Write-Log "Error getting archive folders for $MailboxUPN : $($_.Exception.Message)" -Level ERROR
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
            $search = Get-MgSecurityCaseEdiscoveryCase -EdiscoveryCaseId $CaseId |
                Get-MgSecurityCaseEdiscoveryCaseSearch -EdiscoverySearchId $SearchId

            if ($search.Status -eq "succeeded" -or $search.Status -eq "completed")
            {
                return $true
            }
            elseif ($search.Status -eq "failed")
            {
                Write-Log "Search failed: $SearchId" -Level ERROR
                return $false
            }

            Write-Output "  Search status: $($search.Status) - waiting..."
            Start-Sleep -Seconds $script:SearchPollIntervalSeconds
        }
        catch
        {
            Write-Log "Error checking search status: $($_.Exception.Message)" -Level WARNING
            Start-Sleep -Seconds $script:SearchPollIntervalSeconds
        }
    }

    Write-Log "Search timed out after $TimeoutMinutes minutes" -Level ERROR
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
                Write-Log "Export failed: $OperationId" -Level ERROR
                return $null
            }

            $percentComplete = $operation.PercentProgress
            Write-Output "  Export progress: $percentComplete% - waiting..."
            Start-Sleep -Seconds $script:ExportPollIntervalSeconds
        }
        catch
        {
            Write-Log "Error checking export status: $($_.Exception.Message)" -Level WARNING
            Start-Sleep -Seconds $script:ExportPollIntervalSeconds
        }
    }

    Write-Log "Export timed out after $TimeoutMinutes minutes" -Level ERROR
    return $null
}
#endregion

#region MAIN SCRIPT

try
{
    Write-Log "============================================================"
    Write-Log "Exchange Online Archive Mailbox Export"
    Write-Log "============================================================"
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Log "Case Name: $script:FullCaseName"
    Write-Log "Output Path: $OutputPath"
    Write-Log ""

    #region VALIDATE PREREQUISITES
    Write-Log "Validating prerequisites..."

    # Check output directory
    if (-not (Test-Path $OutputPath))
    {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $OutputPath"
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
    Write-Log "Found $($upnList.Count) mailboxes in CSV"

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

    Write-Log "Prerequisites validated"
    Write-Log ""
    #endregion

    #region CONNECT TO SERVICES
    Write-Log "Connecting to services..."

    # Determine authentication mode
    $useAppAuth = $AppId -and $TenantId -and $CertificateThumbprint

    # Connect to Exchange Online
    $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoSession -or $exoSession.State -ne "Connected")
    {
        Write-Log "Connecting to Exchange Online..."
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
        Write-Log "Using existing Exchange Online session"
    }

    # Connect to Microsoft Graph
    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $graphContext)
    {
        Write-Log "Connecting to Microsoft Graph..."
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
        Write-Log "Using existing Microsoft Graph session"
    }

    Write-Log "Connected to all services"
    Write-Log ""
    #endregion

    #region ENUMERATE ARCHIVES AND GET SIZES
    Write-Log "Enumerating archive mailboxes and sizes..."

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
                Write-Log "Mailbox not found: $upn" -Level WARNING
                continue
            }

            if (-not $mailbox.ArchiveStatus -or $mailbox.ArchiveStatus -eq "None")
            {
                Write-Log "No archive enabled for: $upn" -Level WARNING
                continue
            }

            $archiveStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction SilentlyContinue

            if (-not $archiveStats)
            {
                Write-Log "Could not get archive stats for: $upn" -Level WARNING
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
            Write-Log "Error processing $upn : $($_.Exception.Message)" -Level WARNING
        }
    }

    Write-Progress -Activity "Enumerating Archives" -Completed

    if ($mailboxInfo.Count -eq 0)
    {
        throw "No valid archive mailboxes found to process"
    }

    # Sort by size descending (largest first)
    $mailboxInfo = $mailboxInfo | Sort-Object -Property ArchiveSizeBytes -Descending

    Write-Log "Found $($mailboxInfo.Count) archive mailboxes to process"
    Write-Log "Total archive size: $([math]::Round(($mailboxInfo | Measure-Object -Property ArchiveSizeBytes -Sum).Sum / 1GB, 2)) GB"
    Write-Log "Processing order (largest first):"
    foreach ($mb in $mailboxInfo | Select-Object -First 5)
    {
        Write-Log "  - $($mb.UPN): $($mb.ArchiveSizeGB) GB ($($mb.ItemCount) items)"
    }
    if ($mailboxInfo.Count -gt 5)
    {
        Write-Log "  ... and $($mailboxInfo.Count - 5) more"
    }
    Write-Log ""
    #endregion

    #region CREATE EDISCOVERY CASE
    Write-Log "Creating eDiscovery case: $script:FullCaseName"

    $caseParams = @{
        displayName  = $script:FullCaseName
        description  = "Archive mailbox export for migration - Created $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        externalId   = "ArchiveExport-$script:Timestamp"
    }

    $case = New-MgSecurityCaseEdiscoveryCase -BodyParameter $caseParams
    $caseId = $case.Id

    Write-Log "Created eDiscovery case: $caseId"
    Write-Log ""
    #endregion

    #region CREATE SEARCHES AND EXPORTS
    Write-Log "Creating compliance searches for each mailbox..."

    $results = @()
    $searchCounter = 0

    foreach ($mb in $mailboxInfo)
    {
        $searchCounter++
        $upn = $mb.UPN

        Write-Log "[$searchCounter/$($mailboxInfo.Count)] Processing: $upn ($($mb.ArchiveSizeGB) GB)"

        $result = [PSCustomObject]@{
            UPN              = $upn
            DisplayName      = $mb.DisplayName
            ArchiveSizeGB    = $mb.ArchiveSizeGB
            ItemCount        = $mb.ItemCount
            SearchId         = $null
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

            $searchParams = @{
                displayName      = $searchName
                contentQuery     = $contentQuery
                dataSourceScopes = "allTenantMailboxes"
            }

            $search = New-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $caseId -BodyParameter $searchParams
            $result.SearchId = $search.Id

            Write-Log "  Created search: $($search.Id)"

            # Add mailbox as data source
            $userSource = @{
                "@odata.type" = "microsoft.graph.security.userSource"
                email         = $upn
            }

            New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id -BodyParameter $userSource | Out-Null

            Write-Log "  Added data source: $upn"

            # Start search estimate
            Invoke-MgBetaEstimateSecurityCaseEdiscoveryCaseSearchStatistics -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id | Out-Null

            Write-Log "  Started search estimate"

            # Wait for search to complete
            Write-Log "  Waiting for search to complete..."
            $searchComplete = Wait-SearchCompletion -CaseId $caseId -SearchId $search.Id

            if (-not $searchComplete)
            {
                $result.Status = "SearchFailed"
                $result.ErrorMessage = "Search did not complete successfully"
                $results += $result
                continue
            }

            Write-Log "  Search completed"

            # Export to PST
            $exportParams = @{
                displayName       = "Export_$searchName"
                exportCriteria    = "searchHits"
                exportformats     = "pst"
                additionalOptions = "subfolderContents"
                exportLocation    = "responsiveLocations"
            }

            $null = Export-MgBetaSecurityCaseEdiscoveryCaseSearchResult -EdiscoveryCaseId $caseId -EdiscoverySearchId $search.Id -BodyParameter $exportParams

            # Get the operation ID from the response
            $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $caseId |
                Where-Object { $_.Action -eq "exportResult" } |
                Sort-Object -Property CreatedDateTime -Descending |
                Select-Object -First 1

            if ($operations)
            {
                $result.ExportId = $operations.Id
                Write-Log "  Started export: $($operations.Id)"
            }

            $result.Status = "ExportStarted"
        }
        catch
        {
            $result.Status = "Failed"
            $result.ErrorMessage = $_.Exception.Message
            Write-Log "  Error: $($_.Exception.Message)" -Level ERROR
        }

        $results += $result
    }

    Write-Log ""
    Write-Log "All searches created. Waiting for exports to complete..."
    Write-Log ""
    #endregion

    #region WAIT FOR EXPORTS AND DOWNLOAD
    if (-not $SkipDownload)
    {
        Write-Log "Downloading PST files..."

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
            Write-Log "Interactive download token acquisition..."
            Import-Module MSAL.PS

            $tokenResult = Get-MsalToken -ClientId "b26e684c-5068-4120-a679-64a5d2c909d9" -Scopes "eDiscovery.Download.Read" -Interactive
            $downloadToken = $tokenResult.AccessToken
        }

        foreach ($result in $results | Where-Object { $_.ExportId })
        {
            Write-Log "Waiting for export: $($result.UPN)..."

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

                    Write-Log "  Downloading: $fileName"

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
                        Write-Log "  Downloaded: $outputFile ($($result.PstSizeGB) GB)"
                    }
                }
            }
            catch
            {
                $result.Status = "DownloadFailed"
                $result.ErrorMessage = $_.Exception.Message
                Write-Log "  Download error: $($_.Exception.Message)" -Level ERROR
            }
        }
    }
    else
    {
        Write-Log "SkipDownload specified - PST files available in Purview portal"
        foreach ($result in $results | Where-Object { $_.ExportId })
        {
            $result.Status = "ExportCreated"
        }
    }
    #endregion

    #region GENERATE SUMMARY REPORT
    Write-Log ""
    Write-Log "Generating summary report..."

    $results | Export-Csv -Path $script:SummaryFile -NoTypeInformation

    Write-Log ""
    Write-Log "============================================================"
    Write-Log "EXPORT COMPLETE"
    Write-Log "============================================================"
    Write-Log ""
    Write-Log "Summary:"
    Write-Log "  Total mailboxes: $($mailboxInfo.Count)"
    Write-Log "  Completed: $(($results | Where-Object { $_.Status -eq 'Completed' }).Count)"
    Write-Log "  Failed: $(($results | Where-Object { $_.Status -like '*Failed*' }).Count)"
    Write-Log "  Pending: $(($results | Where-Object { $_.Status -eq 'Pending' -or $_.Status -eq 'ExportStarted' }).Count)"
    Write-Log ""
    Write-Log "Output files:"
    Write-Log "  Summary: $script:SummaryFile"
    Write-Log "  Log: $script:LogFile"
    Write-Log "  PST files: $OutputPath"
    Write-Log ""
    Write-Log "eDiscovery Case: $script:FullCaseName (ID: $caseId)"
    Write-Log "  View in Purview: https://compliance.microsoft.com/ediscovery"
    Write-Log ""
    #endregion
}
catch
{
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    exit 1
}
finally
{
    Write-Log "Script completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

#endregion
