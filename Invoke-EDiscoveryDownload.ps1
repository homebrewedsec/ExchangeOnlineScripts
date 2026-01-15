<#
.SYNOPSIS
    Downloads PST files from Microsoft Purview eDiscovery cases.

.DESCRIPTION
    This script automates the download of PST files from completed eDiscovery export operations.
    It uses the Microsoft Graph API to retrieve export metadata and the MicrosoftPurviewEDiscovery
    API to download files with proper authentication.

    Key Features:
    - Downloads all completed exports from an eDiscovery case
    - Supports case lookup by name or ID
    - Interactive authentication for download token (Microsoft requirement)
    - File verification to ensure valid downloads
    - Summary CSV report of all downloads

    IMPORTANT: Downloads require delegated (interactive) authentication. This is a Microsoft
    limitation - app-only authentication does NOT work for eDiscovery downloads.

.PARAMETER CaseId
    The eDiscovery case ID to download from. Either CaseId or CaseName is required.

.PARAMETER CaseName
    The eDiscovery case name (or partial name) to search for. Either CaseId or CaseName is required.

.PARAMETER ClientId
    Your Azure AD app registration client ID that has the eDiscovery.Download.Read delegated
    permission configured with admin consent. This is required.

.PARAMETER OutputPath
    Directory where PST files will be downloaded. Defaults to current directory.

.PARAMETER DownloadToken
    Pre-captured bearer token for fully automated downloads. Tokens expire after ~1 hour.
    Capture from browser dev tools when downloading manually from Purview portal.

.PARAMETER Force
    Overwrite existing files without prompting.

.PARAMETER Quiet
    Suppress verbose output. Only shows progress, errors, and final summary.

.PARAMETER ThrottleLimit
    Number of concurrent downloads (requires PowerShell 7+). Default is 1 (sequential).
    Recommended: 4-8 for faster downloads if network bandwidth allows.

.EXAMPLE
    .\Invoke-EDiscoveryDownload.ps1 -CaseName "ArchiveExport" -ClientId "c8fab561-2078-48cb-8f93-3789d1d72e8a"
    Downloads all PST files from the case matching "ArchiveExport".

.EXAMPLE
    .\Invoke-EDiscoveryDownload.ps1 -CaseId "abc123..." -ClientId "c8fab561-..." -OutputPath "C:\Downloads"
    Downloads to a specific directory.

.EXAMPLE
    .\Invoke-EDiscoveryDownload.ps1 -CaseName "MyCase" -ClientId "..." -DownloadToken "eyJ0..."
    Uses a pre-captured token for fully automated downloads (no interactive prompt).

.NOTES
    Author: Hudson Bush / Claude AI
    Version: 1.0

    Prerequisites:
    - Microsoft.Graph module: Install-Module Microsoft.Graph -Scope CurrentUser
    - MSAL.PS module: Install-Module MSAL.PS -Scope CurrentUser
    - App registration with eDiscovery.Download.Read (delegated) permission with admin consent
    - App registration redirect URI: https://login.microsoftonline.com/common/oauth2/nativeclient
    - User must be member of the eDiscovery case

    Reference:
    - https://learn.microsoft.com/en-us/purview/edisc-search-export
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CaseId,

    [Parameter(Mandatory = $false)]
    [string]$CaseName,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$OutputPath = (Get-Location).Path,

    [string]$DownloadToken,

    [switch]$Force,

    [switch]$Quiet,

    [int]$ThrottleLimit = 1
)

#region CONFIGURATION
$script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFile = Join-Path $OutputPath "EDiscoveryDownload_$script:Timestamp.log"
$script:SummaryFile = Join-Path $OutputPath "EDiscoveryDownload_Summary_$script:Timestamp.csv"

# MicrosoftPurviewEDiscovery resource for downloads
$script:PurviewResourceId = "b26e684c-5068-4120-a679-64a5d2c909d9"
$script:PurviewScope = "$script:PurviewResourceId/.default"
#endregion

#region LOGGING
function Write-Log
{
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "PROGRESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Always write to log file
    Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue

    # In quiet mode, only show errors and progress
    if ($Quiet -and $Level -notin @("ERROR", "PROGRESS"))
    {
        return
    }

    switch ($Level)
    {
        "ERROR"    { Write-Host $logEntry -ForegroundColor Red }
        "WARNING"  { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS"  { Write-Host $logEntry -ForegroundColor Green }
        "PROGRESS" { Write-Host $Message }
        default    { Write-Host $logEntry }
    }
}
#endregion

#region MAIN SCRIPT
try
{
    Write-Log "============================================================"
    Write-Log "eDiscovery PST Download Script"
    Write-Log "============================================================"
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    # Validate parameters
    if (-not $CaseId -and -not $CaseName)
    {
        throw "Either -CaseId or -CaseName parameter is required"
    }

    Write-Log "Output Path: $OutputPath"
    Write-Log "Client ID: $ClientId"
    if ($DownloadToken) { Write-Log "Using pre-captured download token" }
    Write-Log ""

    #region PREREQUISITES
    Write-Log "Checking prerequisites..."

    # Check output directory
    if (-not (Test-Path $OutputPath))
    {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $OutputPath"
    }

    # Check required modules
    $requiredModules = @("Microsoft.Graph", "MSAL.PS")
    foreach ($module in $requiredModules)
    {
        $installed = Get-Module -Name $module -ListAvailable
        if (-not $installed)
        {
            throw "Required module not found: $module. Install with: Install-Module $module -Scope CurrentUser"
        }
        Write-Log "  Module $module : OK"
    }

    Write-Log "Prerequisites validated"
    Write-Log ""
    #endregion

    #region CONNECT TO GRAPH
    Write-Log "Connecting to Microsoft Graph..."

    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $graphContext)
    {
        Write-Log "No existing Graph connection, connecting interactively..."
        Connect-MgGraph -Scopes "eDiscovery.ReadWrite.All" -NoWelcome
        $graphContext = Get-MgContext
    }

    Write-Log "  Connected as: $($graphContext.Account)"
    Write-Log "  Tenant: $($graphContext.TenantId)"
    $script:TenantId = $graphContext.TenantId
    Write-Log ""
    #endregion

    #region RESOLVE CASE
    Write-Log "Resolving eDiscovery case..."

    $resolvedCaseId = $null
    $resolvedCaseName = $null

    if ($CaseId)
    {
        Write-Log "  Using provided Case ID: $CaseId"
        $case = Get-MgSecurityCaseEdiscoveryCase -EdiscoveryCaseId $CaseId -ErrorAction Stop
        $resolvedCaseId = $case.Id
        $resolvedCaseName = $case.DisplayName
    }
    else
    {
        Write-Log "  Looking up case by name: $CaseName"
        $cases = Get-MgSecurityCaseEdiscoveryCase -All -ErrorAction Stop

        $matchingCases = @($cases | Where-Object { $_.DisplayName -like "*$CaseName*" })

        if ($matchingCases.Count -eq 0)
        {
            Write-Log "Available cases:" -Level WARNING
            foreach ($c in $cases | Select-Object -First 10)
            {
                Write-Log "  - $($c.DisplayName)" -Level WARNING
            }
            throw "No cases found matching: $CaseName"
        }

        if ($matchingCases.Count -gt 1)
        {
            Write-Log "Multiple cases found matching '$CaseName':" -Level WARNING
            foreach ($c in $matchingCases)
            {
                Write-Log "  - $($c.DisplayName) (ID: $($c.Id))" -Level WARNING
            }
            Write-Log "Using first match" -Level WARNING
        }

        $resolvedCaseId = $matchingCases[0].Id
        $resolvedCaseName = $matchingCases[0].DisplayName
    }

    Write-Log "  Case: $resolvedCaseName" -Level SUCCESS
    Write-Log "  Case ID: $resolvedCaseId"
    Write-Log ""
    #endregion

    #region GET EXPORT OPERATIONS
    Write-Log "Getting export operations..."

    $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $resolvedCaseId -ErrorAction Stop
    $exportOps = @($operations | Where-Object { $_.Action -eq "exportResult" -and $_.Status -eq "succeeded" })

    if ($exportOps.Count -eq 0)
    {
        throw "No completed export operations found in case"
    }

    Write-Log "  Found $($exportOps.Count) completed export operation(s)"
    Write-Log ""
    #endregion

    #region GET EXPORT FILE METADATA
    Write-Log "Retrieving download URLs from export metadata..."

    $downloadFiles = @()

    foreach ($op in $exportOps)
    {
        Write-Log "  Processing export: $($op.Id)"

        # Get full export details via REST API to access exportFileMetadata
        $uri = "https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases/$resolvedCaseId/operations/$($op.Id)"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        if ($response.exportFileMetadata)
        {
            foreach ($fileMeta in $response.exportFileMetadata)
            {
                if ($fileMeta.downloadUrl)
                {
                    $downloadFiles += [PSCustomObject]@{
                        ExportId    = $op.Id
                        FileName    = $fileMeta.fileName
                        Size        = $fileMeta.size
                        DownloadUrl = $fileMeta.downloadUrl
                        Status      = "Pending"
                        OutputPath  = $null
                        Error       = $null
                    }
                    Write-Log "    File: $($fileMeta.fileName) ($([math]::Round($fileMeta.size / 1MB, 2)) MB)"
                }
            }
        }
    }

    if ($downloadFiles.Count -eq 0)
    {
        throw "No download URLs found in export metadata. Exports may not be complete."
    }

    Write-Log "  Total files to download: $($downloadFiles.Count)" -Level SUCCESS
    Write-Log ""
    #endregion

    #region ACQUIRE DOWNLOAD TOKEN
    Write-Log "Acquiring download token for MicrosoftPurviewEDiscovery..."
    Write-Log "NOTE: Microsoft requires interactive (delegated) authentication for downloads"

    $token = $null

    if ($DownloadToken)
    {
        Write-Log "  Using pre-captured token"
        $token = $DownloadToken
    }
    else
    {
        Write-Log "  Acquiring token via interactive browser flow..."
        Write-Log "  Please sign in when the browser opens" -Level WARNING

        # Check if MSAL.NET assembly is already loaded
        $msalAssembly = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq "Microsoft.Identity.Client" }

        if (-not $msalAssembly)
        {
            Write-Log "  Loading MSAL.NET assembly..."

            # Search multiple locations for Microsoft.Identity.Client.dll
            $msalModule = Get-Module -Name MSAL.PS -ListAvailable | Select-Object -First 1
            $searchPaths = @(
                (Join-Path $msalModule.ModuleBase "Microsoft.Identity.Client.dll")
                (Join-Path $msalModule.ModuleBase "net45\Microsoft.Identity.Client.dll")
                (Join-Path $msalModule.ModuleBase "netcoreapp2.1\Microsoft.Identity.Client.dll")
                (Join-Path $msalModule.ModuleBase "lib\net45\Microsoft.Identity.Client.dll")
                (Join-Path $msalModule.ModuleBase "lib\netstandard2.0\Microsoft.Identity.Client.dll")
            )

            # Also search recursively in module directory
            $foundDlls = Get-ChildItem -Path $msalModule.ModuleBase -Filter "Microsoft.Identity.Client.dll" -Recurse -ErrorAction SilentlyContinue
            foreach ($dll in $foundDlls)
            {
                $searchPaths += $dll.FullName
            }

            $loaded = $false
            foreach ($path in $searchPaths | Select-Object -Unique)
            {
                if (Test-Path $path)
                {
                    Write-Log "  Found MSAL.NET at: $path"
                    try
                    {
                        Add-Type -Path $path -ErrorAction Stop
                        $loaded = $true
                        break
                    }
                    catch
                    {
                        Write-Log "  Could not load from $path : $($_.Exception.Message)" -Level WARNING
                    }
                }
            }

            if (-not $loaded)
            {
                # Try importing MSAL.PS module which may load the assembly
                Write-Log "  Trying to load via MSAL.PS module import..."
                Import-Module MSAL.PS -ErrorAction SilentlyContinue

                $msalAssembly = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq "Microsoft.Identity.Client" }
                if (-not $msalAssembly)
                {
                    throw "Could not load Microsoft.Identity.Client assembly. MSAL.PS module may be incomplete."
                }
            }
        }

        Write-Log "  MSAL.NET assembly loaded"

        # Build public client with explicit redirect URI
        $authority = "https://login.microsoftonline.com/$script:TenantId"
        $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

        $publicClientBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId)
        $publicClientBuilder = $publicClientBuilder.WithAuthority($authority)
        $publicClientBuilder = $publicClientBuilder.WithRedirectUri($redirectUri)
        $publicClient = $publicClientBuilder.Build()

        # Request token interactively
        $scopes = [System.Collections.Generic.List[string]]::new()
        $scopes.Add($script:PurviewScope)

        $tokenRequest = $publicClient.AcquireTokenInteractive($scopes)
        $authResult = $tokenRequest.ExecuteAsync().GetAwaiter().GetResult()

        if ($authResult -and $authResult.AccessToken)
        {
            $token = $authResult.AccessToken
            Write-Log "  Token acquired successfully" -Level SUCCESS

            # Decode and validate token claims
            $tokenParts = $token -split '\.'
            if ($tokenParts.Count -ge 2)
            {
                $payload = $tokenParts[1]
                $padding = 4 - ($payload.Length % 4)
                if ($padding -lt 4) { $payload += "=" * $padding }
                $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json

                Write-Log "  Token audience: $($decoded.aud)"
                Write-Log "  Token scope: $($decoded.scp)"
            }
        }
        else
        {
            throw "Token acquisition failed - no access token returned"
        }
    }

    Write-Log ""
    #endregion

    #region DOWNLOAD FILES
    Write-Log "============================================================"
    Write-Log "DOWNLOADING FILES"
    Write-Log "============================================================"
    Write-Log ""

    $headers = @{
        'Authorization' = "Bearer $token"
        'X-AllowWithAADToken' = "true"
    }

    # Check if parallel download is possible (PS7+ and ThrottleLimit > 1)
    $useParallel = ($ThrottleLimit -gt 1) -and ($PSVersionTable.PSVersion.Major -ge 7)

    if ($useParallel)
    {
        Write-Log "Parallel download enabled: $ThrottleLimit concurrent downloads" -Level PROGRESS

        # Parallel download using ForEach-Object -Parallel (PS7+)
        $downloadFiles | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $file = $_
            $outputFile = Join-Path $using:OutputPath $file.FileName
            $headers = $using:headers
            $forceOverwrite = $using:Force

            # Check if file exists
            if ((Test-Path $outputFile) -and -not $forceOverwrite)
            {
                Write-Host "SKIP: $($file.FileName) (exists)"
                $file.Status = "Skipped"
                $file.OutputPath = $outputFile
                return
            }

            try
            {
                $startTime = Get-Date
                Write-Host "START: $($file.FileName)"

                $ProgressPreference = 'SilentlyContinue'
                Invoke-WebRequest -Uri $file.DownloadUrl -Headers $headers -OutFile $outputFile -UseBasicParsing -ErrorAction Stop

                $duration = ((Get-Date) - $startTime).TotalSeconds

                if (Test-Path $outputFile)
                {
                    $fileInfo = Get-Item $outputFile
                    $stream = [System.IO.File]::OpenRead($outputFile)
                    $firstBytes = New-Object byte[] 100
                    $stream.Read($firstBytes, 0, 100) | Out-Null
                    $stream.Close()
                    $firstChars = [System.Text.Encoding]::UTF8.GetString($firstBytes)

                    if ($firstChars -match '<html|<!DOCTYPE')
                    {
                        Write-Host "FAIL: $($file.FileName) - HTML error response" -ForegroundColor Red
                        $file.Status = "Failed"
                        $file.Error = "Downloaded file is HTML error page"
                    }
                    elseif ($fileInfo.Length -lt 1000)
                    {
                        Write-Host "WARN: $($file.FileName) - $($fileInfo.Length) bytes" -ForegroundColor Yellow
                        $file.Status = "Warning"
                        $file.OutputPath = $outputFile
                    }
                    else
                    {
                        $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
                        $speed = [math]::Round($sizeMB / $duration, 1)
                        Write-Host "OK: $($file.FileName) - $sizeMB MB in $([math]::Round($duration, 1))s ($speed MB/s)" -ForegroundColor Green
                        $file.Status = "Success"
                        $file.OutputPath = $outputFile
                    }
                }
                else
                {
                    Write-Host "FAIL: $($file.FileName) - File not created" -ForegroundColor Red
                    $file.Status = "Failed"
                    $file.Error = "Output file not created"
                }
            }
            catch
            {
                Write-Host "FAIL: $($file.FileName) - $($_.Exception.Message)" -ForegroundColor Red
                $file.Status = "Failed"
                $file.Error = $_.Exception.Message
            }
        }
    }
    else
    {
        # Sequential download
        $downloadCounter = 0

        foreach ($file in $downloadFiles)
        {
            $downloadCounter++
            Write-Log "[$downloadCounter/$($downloadFiles.Count)] Downloading: $($file.FileName)" -Level PROGRESS

            $outputFile = Join-Path $OutputPath $file.FileName

            if ((Test-Path $outputFile) -and -not $Force)
            {
                Write-Log "  File already exists, skipping (use -Force to overwrite)" -Level WARNING
                $file.Status = "Skipped"
                $file.OutputPath = $outputFile
                continue
            }

            try
            {
                $startTime = Get-Date
                $prevProgressPref = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'

                Invoke-WebRequest -Uri $file.DownloadUrl -Headers $headers -OutFile $outputFile -UseBasicParsing -ErrorAction Stop

                $ProgressPreference = $prevProgressPref
                $duration = ((Get-Date) - $startTime).TotalSeconds

                if (Test-Path $outputFile)
                {
                    $fileInfo = Get-Item $outputFile
                    $stream = [System.IO.File]::OpenRead($outputFile)
                    $firstBytes = New-Object byte[] 100
                    $stream.Read($firstBytes, 0, 100) | Out-Null
                    $stream.Close()
                    $firstChars = [System.Text.Encoding]::UTF8.GetString($firstBytes)

                    if ($firstChars -match '<html|<!DOCTYPE')
                    {
                        Write-Log "  ERROR: Downloaded file is HTML (error response)" -Level ERROR
                        $file.Status = "Failed"
                        $file.Error = "Downloaded file is HTML error page"
                    }
                    elseif ($fileInfo.Length -lt 1000)
                    {
                        Write-Log "  WARNING: File is suspiciously small ($($fileInfo.Length) bytes)" -Level WARNING
                        $file.Status = "Warning"
                        $file.OutputPath = $outputFile
                    }
                    else
                    {
                        $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
                        $speed = [math]::Round($sizeMB / $duration, 1)
                        Write-Log "  OK: $sizeMB MB in $([math]::Round($duration, 1))s ($speed MB/s)" -Level PROGRESS
                        $file.Status = "Success"
                        $file.OutputPath = $outputFile
                    }
                }
                else
                {
                    Write-Log "  ERROR: Output file not created" -Level ERROR
                    $file.Status = "Failed"
                    $file.Error = "Output file not created"
                }
            }
            catch
            {
                Write-Log "  ERROR: $($_.Exception.Message)" -Level ERROR
                $file.Status = "Failed"
                $file.Error = $_.Exception.Message
            }
        }
    }

    Write-Log ""
    #endregion

    #region SUMMARY
    # Calculate counts from status
    $successCount = @($downloadFiles | Where-Object { $_.Status -eq "Success" -or $_.Status -eq "Warning" }).Count
    $failCount = @($downloadFiles | Where-Object { $_.Status -eq "Failed" }).Count
    $skipCount = @($downloadFiles | Where-Object { $_.Status -eq "Skipped" }).Count

    # Export summary CSV
    $downloadFiles | Select-Object ExportId, FileName, @{N='SizeMB';E={[math]::Round($_.Size / 1MB, 2)}}, Status, OutputPath, Error |
        Export-Csv -Path $script:SummaryFile -NoTypeInformation

    Write-Log ""
    $summaryMsg = "Done. $successCount downloaded, $failCount failed"
    if ($skipCount -gt 0) { $summaryMsg += ", $skipCount skipped" }
    Write-Log "$summaryMsg. Summary: $script:SummaryFile" -Level PROGRESS
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
