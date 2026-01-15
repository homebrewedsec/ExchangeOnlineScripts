<#
.SYNOPSIS
    Diagnostic script to test multiple eDiscovery PST download methods with verbose logging.

.DESCRIPTION
    This script attempts to download eDiscovery export files using multiple methods to identify
    what works in your specific environment. It includes comprehensive logging for troubleshooting.

    Methods tested:
    1. MSAL Interactive + exportFileMetadata (current recommended approach)
    2. MSAL Interactive + getDownloadUrl (legacy/deprecated API)
    3. Azure Blob Direct via SAS URL (if available)
    4. .NET HttpClient with chunked/streaming download
    5. BITS Transfer (handles interruptions/resumes)
    6. Browser-based token capture guidance

    The script logs every API call, response, and error to help identify issues.

.PARAMETER CaseId
    The eDiscovery case ID to download from. Required.

.PARAMETER ExportOperationId
    Specific export operation ID. If not provided, script will auto-detect completed exports.

.PARAMETER OutputPath
    Directory for downloads and logs. Defaults to current directory.

.PARAMETER AppId
    Azure AD application ID for Graph operations. Optional for interactive auth.

.PARAMETER TenantId
    Azure AD tenant ID. Required for app auth, optional for interactive.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app authentication to Graph.

.PARAMETER ClientSecret
    Client secret for app authentication to Graph.

.PARAMETER DownloadToken
    Pre-captured bearer token for downloads. Use this for fully automated runs.
    Capture from browser dev tools when downloading from Purview portal.
    NOTE: Tokens expire after ~1 hour.

.PARAMETER SkipMethods
    Array of method numbers to skip (e.g., 1,2,3). Useful for targeted testing.

.EXAMPLE
    .\Test-EDiscoveryDownload.ps1 -CaseName "ArchiveExport" -AppId "xxx" -TenantId "yyy" -ClientSecret "zzz"
    Uses client secret for Graph API, then prompts for device code auth for downloads.

.EXAMPLE
    .\Test-EDiscoveryDownload.ps1 -CaseId "abc123-def456" -AppId "xxx" -TenantId "yyy" -ClientSecret "zzz"
    Tests downloads for specific case ID with app auth for Graph operations.

.EXAMPLE
    .\Test-EDiscoveryDownload.ps1 -CaseName "MyCase" -DownloadToken "eyJ0eXAiOiJKV1Q..."
    Uses a pre-captured token for fully automated downloads (no interactive auth).

.EXAMPLE
    .\Test-EDiscoveryDownload.ps1 -CaseName "MyCase" -SkipMethods 5,6
    Tests specific case, skipping BITS and browser token methods.

.NOTES
    Author: Hudson Bush / Claude AI
    Version: 1.0

    Prerequisites:
    - Microsoft.Graph module
    - MSAL.PS module (for download token acquisition)
    - App registration with eDiscovery.ReadWrite.All (Graph) and eDiscovery.Download.Read (MicrosoftPurviewEDiscovery)
    - User must be member of the eDiscovery case

    References:
    - https://michev.info/blog/post/5806/using-the-graph-api-to-export-ediscovery-premium-datasets
    - https://learn.microsoft.com/en-us/purview/edisc-search-export
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CaseId,

    [Parameter(Mandatory = $false)]
    [string]$CaseName,

    [string]$ExportOperationId,

    [string]$OutputPath = (Get-Location).Path,

    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ClientSecret,

    [string]$DownloadToken,

    [int[]]$SkipMethods = @()
)

#region CONFIGURATION
$script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFile = Join-Path $OutputPath "EDiscoveryDownloadTest_$script:Timestamp.log"
$script:ResultsFile = Join-Path $OutputPath "EDiscoveryDownloadTest_Results_$script:Timestamp.json"

# MicrosoftPurviewEDiscovery resource for downloads
$script:PurviewResourceId = "b26e684c-5068-4120-a679-64a5d2c909d9"
$script:PurviewScope = "$script:PurviewResourceId/.default"

# Regional proxy endpoints for network testing
$script:ProxyEndpoints = @(
    @{ Region = "NAM"; Url = "https://nam.proxyservice.ediscovery.svc.cloud.microsoft" }
    @{ Region = "EUR"; Url = "https://eur.proxyservice.ediscovery.svc.cloud.microsoft" }
    @{ Region = "APC"; Url = "https://apc.proxyservice.ediscovery.svc.cloud.microsoft" }
)

# Results tracking
$script:Results = @{
    Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    CaseId         = $null  # Will be populated after resolution
    CaseName       = $CaseName
    ExportOperationId = $ExportOperationId
    Prerequisites  = @{}
    ExportMetadata = $null
    Methods        = @{}
    NetworkDiagnostics = @{}
    Recommendations = @()
}
#endregion

#region LOGGING
function Write-Log
{
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message = "",

        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS")]
        [string]$Level = "INFO",

        [string]$Method = "GENERAL"
    )

    # Handle empty messages (used for spacing)
    if ([string]::IsNullOrEmpty($Message))
    {
        Write-Host ""
        Add-Content -Path $script:LogFile -Value "" -ErrorAction SilentlyContinue
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logEntry = "[$timestamp] [$Method] [$Level] $Message"

    # Console output with colors
    switch ($Level)
    {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Gray }
        default   { Write-Host $logEntry }
    }

    # File output
    Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
}

function Write-LogSection
{
    param([string]$Title)

    $separator = "=" * 80
    Write-Log $separator
    Write-Log $Title
    Write-Log $separator
}
#endregion

#region PREREQUISITE CHECKS
function Test-Prerequisites
{
    Write-LogSection "PHASE 1: PREREQUISITE CHECKS"

    $prereqResults = @{
        Modules = @{}
        ServicePrincipal = $null
        OutputDirectory = $null
        GraphConnection = $null
    }

    # Check required modules
    $requiredModules = @("Microsoft.Graph", "MSAL.PS")
    foreach ($module in $requiredModules)
    {
        Write-Log "Checking module: $module" -Method "PREREQ"
        $installed = Get-Module -Name $module -ListAvailable
        if ($installed)
        {
            $version = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()
            Write-Log "  Found: $module v$version" -Level SUCCESS -Method "PREREQ"
            $prereqResults.Modules[$module] = @{ Installed = $true; Version = $version }
        }
        else
        {
            Write-Log "  NOT FOUND: $module - Install with: Install-Module $module -Scope CurrentUser" -Level ERROR -Method "PREREQ"
            $prereqResults.Modules[$module] = @{ Installed = $false; Version = $null }
        }
    }

    # Check optional modules
    $optionalModules = @("ExchangeOnlineManagement", "BitsTransfer")
    foreach ($module in $optionalModules)
    {
        Write-Log "Checking optional module: $module" -Method "PREREQ"
        $installed = Get-Module -Name $module -ListAvailable
        if ($installed)
        {
            $version = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()
            Write-Log "  Found: $module v$version" -Level SUCCESS -Method "PREREQ"
            $prereqResults.Modules[$module] = @{ Installed = $true; Version = $version; Optional = $true }
        }
        else
        {
            Write-Log "  Not found (optional): $module" -Level WARNING -Method "PREREQ"
            $prereqResults.Modules[$module] = @{ Installed = $false; Version = $null; Optional = $true }
        }
    }

    # Check output directory
    Write-Log "Checking output directory: $OutputPath" -Method "PREREQ"
    if (-not (Test-Path $OutputPath))
    {
        try
        {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            Write-Log "  Created directory: $OutputPath" -Level SUCCESS -Method "PREREQ"
            $prereqResults.OutputDirectory = @{ Exists = $true; Created = $true; Path = $OutputPath }
        }
        catch
        {
            Write-Log "  Failed to create directory: $($_.Exception.Message)" -Level ERROR -Method "PREREQ"
            $prereqResults.OutputDirectory = @{ Exists = $false; Error = $_.Exception.Message }
        }
    }
    else
    {
        Write-Log "  Directory exists: $OutputPath" -Level SUCCESS -Method "PREREQ"
        $prereqResults.OutputDirectory = @{ Exists = $true; Created = $false; Path = $OutputPath }
    }

    # Fail if required modules missing
    $missingRequired = $requiredModules | Where-Object { -not $prereqResults.Modules[$_].Installed }
    if ($missingRequired)
    {
        Write-Log "Missing required modules: $($missingRequired -join ', ')" -Level ERROR -Method "PREREQ"
        Write-Log "Install with: Install-Module $($missingRequired -join ',') -Scope CurrentUser" -Level ERROR -Method "PREREQ"
        $script:Results.Prerequisites = $prereqResults
        throw "Missing required modules. Cannot continue."
    }

    $script:Results.Prerequisites = $prereqResults
    Write-Log "Prerequisite checks completed" -Level SUCCESS -Method "PREREQ"
    return $prereqResults
}
#endregion

#region GRAPH CONNECTION
function Connect-ToGraph
{
    Write-LogSection "PHASE 2: MICROSOFT GRAPH CONNECTION"

    $connectionResult = @{
        Method = $null
        Success = $false
        Context = $null
        Error = $null
    }

    # Determine auth method
    $useAppAuth = ($AppId -and $TenantId) -and ($CertificateThumbprint -or $ClientSecret)

    # Check existing connection - but if app auth requested, reconnect with app credentials
    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($existingContext -and -not $useAppAuth)
    {
        Write-Log "Using existing Graph connection" -Level SUCCESS -Method "GRAPH"
        Write-Log "  Account: $($existingContext.Account)" -Method "GRAPH"
        Write-Log "  TenantId: $($existingContext.TenantId)" -Method "GRAPH"
        Write-Log "  Scopes: $($existingContext.Scopes -join ', ')" -Method "GRAPH"
        $connectionResult.Method = "Existing"
        $connectionResult.Success = $true
        $connectionResult.Context = @{
            Account = $existingContext.Account
            TenantId = $existingContext.TenantId
            Scopes = $existingContext.Scopes
        }
        return $connectionResult
    }
    elseif ($existingContext -and $useAppAuth)
    {
        Write-Log "Disconnecting existing session to use app authentication..." -Method "GRAPH"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    try
    {
        if ($useAppAuth)
        {
            Write-Log "Connecting with app authentication..." -Method "GRAPH"

            if ($ClientSecret)
            {
                Write-Log "  Using client secret" -Method "GRAPH"
                $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
                $connectionResult.Method = "ClientSecret"
            }
            else
            {
                Write-Log "  Using certificate: $CertificateThumbprint" -Method "GRAPH"
                $cert = Get-ChildItem "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if (-not $cert)
                {
                    $cert = Get-ChildItem "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                }
                if (-not $cert)
                {
                    throw "Certificate not found: $CertificateThumbprint"
                }
                Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert -NoWelcome
                $connectionResult.Method = "Certificate"
            }
        }
        else
        {
            Write-Log "Connecting with interactive authentication..." -Method "GRAPH"
            Write-Log "  Required scopes: eDiscovery.ReadWrite.All" -Method "GRAPH"
            Connect-MgGraph -Scopes "eDiscovery.ReadWrite.All" -NoWelcome
            $connectionResult.Method = "Interactive"
        }

        $context = Get-MgContext
        Write-Log "Connected successfully" -Level SUCCESS -Method "GRAPH"
        Write-Log "  Account: $($context.Account)" -Method "GRAPH"
        Write-Log "  TenantId: $($context.TenantId)" -Method "GRAPH"

        $connectionResult.Success = $true
        $connectionResult.Context = @{
            Account = $context.Account
            TenantId = $context.TenantId
            Scopes = $context.Scopes
        }

        # Store tenant ID for later use
        if (-not $script:TenantId -and $context.TenantId)
        {
            $script:TenantId = $context.TenantId
        }
    }
    catch
    {
        Write-Log "Graph connection failed: $($_.Exception.Message)" -Level ERROR -Method "GRAPH"
        $connectionResult.Error = $_.Exception.Message
    }

    $script:Results.Prerequisites.GraphConnection = $connectionResult
    return $connectionResult
}
#endregion

#region SERVICE PRINCIPAL CHECK
function Test-PurviewServicePrincipal
{
    Write-LogSection "CHECKING PURVIEW SERVICE PRINCIPAL"

    $spResult = @{
        Exists = $false
        DisplayName = $null
        AppId = $script:PurviewResourceId
        Error = $null
    }

    try
    {
        Write-Log "Checking for MicrosoftPurviewEDiscovery service principal..." -Method "SP-CHECK"
        Write-Log "  Resource ID: $script:PurviewResourceId" -Method "SP-CHECK"

        $sp = Get-MgServicePrincipal -Filter "appId eq '$script:PurviewResourceId'" -ErrorAction Stop

        if ($sp)
        {
            Write-Log "Service principal found:" -Level SUCCESS -Method "SP-CHECK"
            Write-Log "  DisplayName: $($sp.DisplayName)" -Method "SP-CHECK"
            Write-Log "  Id: $($sp.Id)" -Method "SP-CHECK"
            $spResult.Exists = $true
            $spResult.DisplayName = $sp.DisplayName
            $spResult.Id = $sp.Id
        }
        else
        {
            Write-Log "Service principal NOT found in tenant" -Level WARNING -Method "SP-CHECK"
            Write-Log "To create it, run:" -Level WARNING -Method "SP-CHECK"
            Write-Log "  New-MgServicePrincipal -AppId '$script:PurviewResourceId'" -Level WARNING -Method "SP-CHECK"
        }
    }
    catch
    {
        Write-Log "Error checking service principal: $($_.Exception.Message)" -Level ERROR -Method "SP-CHECK"
        $spResult.Error = $_.Exception.Message
    }

    $script:Results.Prerequisites.ServicePrincipal = $spResult
    return $spResult
}
#endregion

#region RESOLVE CASE ID
function Resolve-CaseId
{
    Write-LogSection "RESOLVING CASE"

    if ($script:CaseId)
    {
        Write-Log "Using provided Case ID: $script:CaseId" -Method "RESOLVE"
        $script:Results.CaseId = $script:CaseId
        return $script:CaseId
    }

    if (-not $script:CaseName)
    {
        throw "Either CaseId or CaseName must be provided"
    }

    Write-Log "Looking up case by name: $script:CaseName" -Method "RESOLVE"

    try
    {
        # Get all cases and filter by name
        $cases = Get-MgSecurityCaseEdiscoveryCase -All -ErrorAction Stop

        Write-Log "Found $($cases.Count) total eDiscovery cases" -Method "RESOLVE"

        # Find matching case(s)
        $matchingCases = @($cases | Where-Object { $_.DisplayName -like "*$script:CaseName*" })

        if ($matchingCases.Count -eq 0)
        {
            Write-Log "No cases found matching: $script:CaseName" -Level ERROR -Method "RESOLVE"
            Write-Log "Available cases:" -Method "RESOLVE"
            foreach ($c in $cases | Select-Object -First 10)
            {
                Write-Log "  - $($c.DisplayName) (ID: $($c.Id))" -Method "RESOLVE"
            }
            throw "Case not found: $script:CaseName"
        }

        if ($matchingCases.Count -gt 1)
        {
            Write-Log "Multiple cases found matching '$script:CaseName':" -Level WARNING -Method "RESOLVE"
            foreach ($c in $matchingCases)
            {
                Write-Log "  - $($c.DisplayName) (ID: $($c.Id), Status: $($c.Status))" -Method "RESOLVE"
            }
            Write-Log "Using first match: $($matchingCases[0].DisplayName)" -Method "RESOLVE"
        }

        $selectedCase = $matchingCases[0]
        Write-Log "Found case: $($selectedCase.DisplayName)" -Level SUCCESS -Method "RESOLVE"
        Write-Log "  Case ID: $($selectedCase.Id)" -Method "RESOLVE"
        Write-Log "  Status: $($selectedCase.Status)" -Method "RESOLVE"

        $script:CaseId = $selectedCase.Id
        $script:Results.CaseId = $selectedCase.Id
        $script:Results.CaseName = $selectedCase.DisplayName
        return $selectedCase.Id
    }
    catch
    {
        Write-Log "Error resolving case: $($_.Exception.Message)" -Level ERROR -Method "RESOLVE"
        throw
    }
}
#endregion

#region GET EXPORT METADATA
function Get-ExportMetadata
{
    Write-LogSection "PHASE 3: RETRIEVING EXPORT METADATA"

    $metadataResult = @{
        CaseFound = $false
        CaseName = $null
        ExportOperations = @()
        SelectedExport = $null
        ExportFileMetadata = @()
        DownloadUrls = @()
        Error = $null
    }

    try
    {
        # Resolve case ID from name if needed
        $resolvedCaseId = Resolve-CaseId

        # Get case details
        Write-Log "Getting eDiscovery case details: $resolvedCaseId" -Method "METADATA"
        $case = Get-MgSecurityCaseEdiscoveryCase -EdiscoveryCaseId $resolvedCaseId -ErrorAction Stop

        if (-not $case)
        {
            throw "Case not found: $resolvedCaseId"
        }

        Write-Log "Case found:" -Level SUCCESS -Method "METADATA"
        Write-Log "  DisplayName: $($case.DisplayName)" -Method "METADATA"
        Write-Log "  Status: $($case.Status)" -Method "METADATA"
        Write-Log "  CreatedDateTime: $($case.CreatedDateTime)" -Method "METADATA"

        $metadataResult.CaseFound = $true
        $metadataResult.CaseName = $case.DisplayName
        $metadataResult.CaseStatus = $case.Status

        # Get export operations
        Write-Log "Getting export operations for case..." -Method "METADATA"
        $operations = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $resolvedCaseId -ErrorAction Stop

        $exportOps = @($operations | Where-Object { $_.Action -eq "exportResult" })
        Write-Log "Found $($exportOps.Count) export operation(s)" -Method "METADATA"

        foreach ($op in $exportOps)
        {
            Write-Log "  Operation: $($op.Id)" -Method "METADATA"
            Write-Log "    Status: $($op.Status)" -Method "METADATA"
            Write-Log "    Progress: $($op.PercentProgress)%" -Method "METADATA"
            Write-Log "    Created: $($op.CreatedDateTime)" -Method "METADATA"

            $metadataResult.ExportOperations += @{
                Id = $op.Id
                Status = $op.Status
                PercentProgress = $op.PercentProgress
                CreatedDateTime = $op.CreatedDateTime
            }
        }

        # Select export operation
        $selectedOp = $null
        if ($ExportOperationId)
        {
            $selectedOp = $exportOps | Where-Object { $_.Id -eq $ExportOperationId }
            if (-not $selectedOp)
            {
                throw "Export operation not found: $ExportOperationId"
            }
        }
        else
        {
            # Auto-select most recent succeeded export
            $selectedOp = $exportOps | Where-Object { $_.Status -eq "succeeded" } |
                Sort-Object CreatedDateTime -Descending |
                Select-Object -First 1

            if (-not $selectedOp)
            {
                # Try any completed export
                $selectedOp = $exportOps | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
            }
        }

        if (-not $selectedOp)
        {
            throw "No export operations found in case"
        }

        Write-Log "Selected export operation: $($selectedOp.Id)" -Level SUCCESS -Method "METADATA"
        Write-Log "  Status: $($selectedOp.Status)" -Method "METADATA"
        $script:ExportOperationId = $selectedOp.Id
        $metadataResult.SelectedExport = @{
            Id = $selectedOp.Id
            Status = $selectedOp.Status
        }

        # Get full export details via REST API to access exportFileMetadata
        Write-Log "Retrieving exportFileMetadata via REST API..." -Method "METADATA"
        $uri = "https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases/$resolvedCaseId/operations/$($selectedOp.Id)"
        Write-Log "  URI: $uri" -Level DEBUG -Method "METADATA"

        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        Write-Log "  Response received" -Level DEBUG -Method "METADATA"

        # Log full response for debugging
        $responseJson = $response | ConvertTo-Json -Depth 10
        Write-Log "  Full response:" -Level DEBUG -Method "METADATA"
        foreach ($line in ($responseJson -split "`n"))
        {
            Write-Log "    $line" -Level DEBUG -Method "METADATA"
        }

        # Check for exportFileMetadata
        if ($response.exportFileMetadata)
        {
            Write-Log "Found exportFileMetadata:" -Level SUCCESS -Method "METADATA"
            foreach ($fileMeta in $response.exportFileMetadata)
            {
                Write-Log "  File: $($fileMeta.fileName)" -Method "METADATA"
                Write-Log "    Size: $($fileMeta.size) bytes" -Method "METADATA"
                Write-Log "    URL present: $([bool]$fileMeta.downloadUrl)" -Method "METADATA"

                $metadataResult.ExportFileMetadata += @{
                    FileName = $fileMeta.fileName
                    Size = $fileMeta.size
                    DownloadUrl = $fileMeta.downloadUrl
                }

                if ($fileMeta.downloadUrl)
                {
                    $metadataResult.DownloadUrls += $fileMeta.downloadUrl
                }
            }
        }
        else
        {
            Write-Log "exportFileMetadata not found in response" -Level WARNING -Method "METADATA"
            Write-Log "This may indicate export is not complete or a permissions issue" -Level WARNING -Method "METADATA"
        }

        # Also check AdditionalProperties
        if ($response.AdditionalProperties -and $response.AdditionalProperties.exportFileMetadata)
        {
            Write-Log "Found exportFileMetadata in AdditionalProperties" -Method "METADATA"
            foreach ($fileMeta in $response.AdditionalProperties.exportFileMetadata)
            {
                $metadataResult.ExportFileMetadata += @{
                    FileName = $fileMeta.fileName
                    Size = $fileMeta.size
                    DownloadUrl = $fileMeta.downloadUrl
                    Source = "AdditionalProperties"
                }

                if ($fileMeta.downloadUrl)
                {
                    $metadataResult.DownloadUrls += $fileMeta.downloadUrl
                }
            }
        }

        # Try beta API as well
        Write-Log "Also trying beta API for more details..." -Method "METADATA"
        $betaUri = "https://graph.microsoft.com/beta/security/cases/ediscoveryCases/$resolvedCaseId/operations/$($selectedOp.Id)"
        try
        {
            $betaResponse = Invoke-MgGraphRequest -Method GET -Uri $betaUri -ErrorAction Stop
            Write-Log "Beta API response received" -Level DEBUG -Method "METADATA"

            $betaJson = $betaResponse | ConvertTo-Json -Depth 10
            foreach ($line in ($betaJson -split "`n"))
            {
                Write-Log "  [BETA] $line" -Level DEBUG -Method "METADATA"
            }

            # Check for downloadUrl property directly
            if ($betaResponse.downloadUrl)
            {
                Write-Log "Found downloadUrl in beta response" -Level SUCCESS -Method "METADATA"
                $metadataResult.DownloadUrls += $betaResponse.downloadUrl
            }

            if ($betaResponse.exportFileMetadata)
            {
                Write-Log "Found exportFileMetadata in beta response" -Level SUCCESS -Method "METADATA"
                foreach ($fileMeta in $betaResponse.exportFileMetadata)
                {
                    if ($fileMeta.downloadUrl -and $fileMeta.downloadUrl -notin $metadataResult.DownloadUrls)
                    {
                        $metadataResult.DownloadUrls += $fileMeta.downloadUrl
                        $metadataResult.ExportFileMetadata += @{
                            FileName = $fileMeta.fileName
                            Size = $fileMeta.size
                            DownloadUrl = $fileMeta.downloadUrl
                            Source = "BetaAPI"
                        }
                    }
                }
            }
        }
        catch
        {
            Write-Log "Beta API request failed: $($_.Exception.Message)" -Level WARNING -Method "METADATA"
        }

        Write-Log "Total download URLs found: $($metadataResult.DownloadUrls.Count)" -Method "METADATA"
    }
    catch
    {
        Write-Log "Error retrieving metadata: $($_.Exception.Message)" -Level ERROR -Method "METADATA"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG -Method "METADATA"
        $metadataResult.Error = $_.Exception.Message
    }

    $script:Results.ExportMetadata = $metadataResult
    return $metadataResult
}
#endregion

#region MSAL TOKEN ACQUISITION
function Get-PurviewDownloadToken
{
    Write-LogSection "PHASE 4: MSAL TOKEN ACQUISITION"

    $tokenResult = @{
        Success = $false
        Method = $null
        AccessToken = $null
        ExpiresOn = $null
        Audience = $null
        Scopes = $null
        Error = $null
    }

    # Import MSAL.PS
    try
    {
        Import-Module MSAL.PS -ErrorAction Stop
        Write-Log "MSAL.PS module loaded" -Level SUCCESS -Method "TOKEN"
    }
    catch
    {
        Write-Log "Failed to import MSAL.PS: $($_.Exception.Message)" -Level ERROR -Method "TOKEN"
        $tokenResult.Error = "Failed to import MSAL.PS module"
        $script:Results.Methods.TokenAcquisition = $tokenResult
        return $tokenResult
    }

    Write-Log "Acquiring token for MicrosoftPurviewEDiscovery..." -Method "TOKEN"
    Write-Log "  Resource ID: $script:PurviewResourceId" -Method "TOKEN"
    Write-Log "  Scope: $script:PurviewScope" -Method "TOKEN"

    # Check if pre-captured token was provided
    if ($DownloadToken)
    {
        Write-Log "Using pre-captured download token" -Level SUCCESS -Method "TOKEN"
        Write-Log "  Token length: $($DownloadToken.Length) characters" -Method "TOKEN"

        $script:PurviewToken = $DownloadToken
        $tokenResult.Success = $true
        $tokenResult.Method = "Pre-captured"
        $tokenResult.AccessToken = $DownloadToken

        $script:Results.Methods.TokenAcquisition = $tokenResult
        return $tokenResult
    }

    Write-Log "NOTE: Microsoft only supports DELEGATED permissions for downloads" -Level WARNING -Method "TOKEN"
    Write-Log "      App-only authentication CANNOT work - this is a Microsoft limitation" -Level WARNING -Method "TOKEN"

    try
    {
        # Determine client ID to use - for delegated auth, use Azure PowerShell default
        # The app registration's client ID won't work unless it has delegated permissions configured
        $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"  # Azure PowerShell default - has delegated permissions
        Write-Log "  Using ClientId: $clientId (Azure PowerShell - supports delegated auth)" -Method "TOKEN"

        # Try device code flow first (works better for scripts/servers)
        Write-Log "Attempting device code flow authentication..." -Method "TOKEN"
        Write-Log "This will display a code - enter it at https://microsoft.com/devicelogin" -Level WARNING -Method "TOKEN"

        $token = $null
        try
        {
            $token = Get-MsalToken -ClientId $clientId -TenantId $script:TenantId -Scopes @($script:PurviewScope) -DeviceCode -ErrorAction Stop
        }
        catch
        {
            Write-Log "Device code flow failed: $($_.Exception.Message)" -Level WARNING -Method "TOKEN"
            Write-Log "Falling back to interactive browser authentication..." -Method "TOKEN"

            # Fallback to interactive
            $token = Get-MsalToken -ClientId $clientId -TenantId $script:TenantId -Scopes @($script:PurviewScope) -Interactive -ErrorAction Stop
        }

        if ($token -and $token.AccessToken)
        {
            Write-Log "Token acquired successfully" -Level SUCCESS -Method "TOKEN"
            Write-Log "  Token length: $($token.AccessToken.Length) characters" -Method "TOKEN"
            Write-Log "  ExpiresOn: $($token.ExpiresOn)" -Method "TOKEN"

            # Decode and log token claims (without exposing full token)
            try
            {
                $tokenParts = $token.AccessToken -split '\.'
                if ($tokenParts.Count -ge 2)
                {
                    $payload = $tokenParts[1]
                    # Add padding if needed
                    $padding = 4 - ($payload.Length % 4)
                    if ($padding -lt 4) { $payload += "=" * $padding }
                    $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json

                    Write-Log "Token claims:" -Method "TOKEN"
                    Write-Log "  aud (audience): $($decoded.aud)" -Method "TOKEN"
                    Write-Log "  iss (issuer): $($decoded.iss)" -Method "TOKEN"
                    Write-Log "  scp (scopes): $($decoded.scp)" -Method "TOKEN"
                    Write-Log "  upn (user): $($decoded.upn)" -Method "TOKEN"

                    $tokenResult.Audience = $decoded.aud
                    $tokenResult.Scopes = $decoded.scp
                }
            }
            catch
            {
                Write-Log "Could not decode token claims: $($_.Exception.Message)" -Level WARNING -Method "TOKEN"
            }

            $tokenResult.Success = $true
            $tokenResult.Method = "MSAL.PS Interactive"
            $tokenResult.AccessToken = $token.AccessToken
            $tokenResult.ExpiresOn = $token.ExpiresOn.ToString()

            # Store for use by download methods
            $script:PurviewToken = $token.AccessToken
        }
        else
        {
            Write-Log "Token acquisition returned empty token" -Level ERROR -Method "TOKEN"
            $tokenResult.Error = "Token acquisition returned empty token"
        }
    }
    catch
    {
        Write-Log "Token acquisition failed: $($_.Exception.Message)" -Level ERROR -Method "TOKEN"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG -Method "TOKEN"
        $tokenResult.Error = $_.Exception.Message

        # Common error guidance
        if ($_.Exception.Message -match "AADSTS")
        {
            Write-Log "This appears to be an Azure AD error. Common causes:" -Level WARNING -Method "TOKEN"
            Write-Log "  - App registration doesn't have eDiscovery.Download.Read permission" -Level WARNING -Method "TOKEN"
            Write-Log "  - Admin consent not granted for the permission" -Level WARNING -Method "TOKEN"
            Write-Log "  - MicrosoftPurviewEDiscovery service principal not in tenant" -Level WARNING -Method "TOKEN"
        }
    }

    $script:Results.Methods.TokenAcquisition = $tokenResult
    return $tokenResult
}
#endregion

#region METHOD 1: INVOKE-WEBREQUEST WITH EXPORTFILEMETADATA
function Test-Method1-InvokeWebRequest
{
    if (1 -in $SkipMethods)
    {
        Write-Log "Method 1 skipped by user" -Level WARNING -Method "METHOD-1"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 1: INVOKE-WEBREQUEST + EXPORTFILEMETADATA"

    $result = @{
        Method = "Invoke-WebRequest with exportFileMetadata URLs"
        Success = $false
        Downloads = @()
        Error = $null
    }

    if (-not $script:PurviewToken)
    {
        Write-Log "No Purview token available - cannot proceed" -Level ERROR -Method "METHOD-1"
        $result.Error = "No Purview token available"
        $script:Results.Methods.Method1 = $result
        return $result
    }

    if (-not $script:Results.ExportMetadata.DownloadUrls -or $script:Results.ExportMetadata.DownloadUrls.Count -eq 0)
    {
        Write-Log "No download URLs found in export metadata" -Level ERROR -Method "METHOD-1"
        $result.Error = "No download URLs available"
        $script:Results.Methods.Method1 = $result
        return $result
    }

    Write-Log "Testing download with Invoke-WebRequest..." -Method "METHOD-1"
    Write-Log "Using X-AllowWithAADToken header (CRITICAL for downloads)" -Method "METHOD-1"

    $headers = @{
        'Authorization' = "Bearer $script:PurviewToken"
        'X-AllowWithAADToken' = "true"
        'Content-Type' = 'application/json'
    }

    # Log headers (without exposing full token)
    Write-Log "Request headers:" -Level DEBUG -Method "METHOD-1"
    Write-Log "  Authorization: Bearer [token-length: $($script:PurviewToken.Length)]" -Level DEBUG -Method "METHOD-1"
    Write-Log "  X-AllowWithAADToken: true" -Level DEBUG -Method "METHOD-1"

    $downloadIndex = 0
    foreach ($url in $script:Results.ExportMetadata.DownloadUrls)
    {
        $downloadIndex++
        Write-Log "Download $downloadIndex of $($script:Results.ExportMetadata.DownloadUrls.Count)" -Method "METHOD-1"

        # Get filename from metadata or URL
        $fileName = $null
        if ($script:Results.ExportMetadata.ExportFileMetadata[$downloadIndex - 1])
        {
            $fileName = $script:Results.ExportMetadata.ExportFileMetadata[$downloadIndex - 1].FileName
        }
        if (-not $fileName)
        {
            $fileName = "download_$downloadIndex.zip"
        }

        $outputFile = Join-Path $OutputPath "Method1_$fileName"

        Write-Log "  URL: $($url.Substring(0, [Math]::Min(100, $url.Length)))..." -Level DEBUG -Method "METHOD-1"
        Write-Log "  Output: $outputFile" -Method "METHOD-1"

        $downloadResult = @{
            FileName = $fileName
            OutputFile = $outputFile
            Success = $false
            StatusCode = $null
            ContentLength = $null
            Error = $null
        }

        try
        {
            $startTime = Get-Date
            Write-Log "  Starting download..." -Method "METHOD-1"

            $response = Invoke-WebRequest -Uri $url -Headers $headers -OutFile $outputFile -PassThru -ErrorAction Stop

            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds

            Write-Log "  Response received" -Level SUCCESS -Method "METHOD-1"
            Write-Log "    StatusCode: $($response.StatusCode)" -Method "METHOD-1"
            Write-Log "    Content-Length: $($response.Headers.'Content-Length')" -Method "METHOD-1"
            Write-Log "    Content-Type: $($response.Headers.'Content-Type')" -Method "METHOD-1"
            Write-Log "    Duration: $duration seconds" -Method "METHOD-1"

            $downloadResult.StatusCode = $response.StatusCode
            $downloadResult.ContentLength = $response.Headers.'Content-Length'
            $downloadResult.ContentType = $response.Headers.'Content-Type'
            $downloadResult.Duration = $duration

            # Verify file was created and is not HTML
            if (Test-Path $outputFile)
            {
                $fileInfo = Get-Item $outputFile
                Write-Log "  File created: $($fileInfo.Length) bytes" -Method "METHOD-1"
                $downloadResult.FileSize = $fileInfo.Length

                # Check if file is HTML (error response)
                $firstBytes = [System.IO.File]::ReadAllBytes($outputFile) | Select-Object -First 100
                $firstChars = [System.Text.Encoding]::UTF8.GetString($firstBytes)

                if ($firstChars -match '<html|<!DOCTYPE')
                {
                    Write-Log "  WARNING: File appears to be HTML (likely error response)" -Level WARNING -Method "METHOD-1"
                    Write-Log "  First 100 chars: $firstChars" -Level DEBUG -Method "METHOD-1"
                    $downloadResult.Error = "Downloaded file is HTML (error response)"
                }
                elseif ($fileInfo.Length -lt 1000)
                {
                    Write-Log "  WARNING: File is suspiciously small ($($fileInfo.Length) bytes)" -Level WARNING -Method "METHOD-1"
                    $content = Get-Content $outputFile -Raw -ErrorAction SilentlyContinue
                    Write-Log "  Content: $content" -Level DEBUG -Method "METHOD-1"
                    $downloadResult.Error = "File too small - may be error response"
                }
                else
                {
                    Write-Log "  Download appears successful!" -Level SUCCESS -Method "METHOD-1"
                    $downloadResult.Success = $true
                    $result.Success = $true
                }
            }
            else
            {
                Write-Log "  Output file not created" -Level ERROR -Method "METHOD-1"
                $downloadResult.Error = "Output file not created"
            }
        }
        catch
        {
            Write-Log "  Download failed: $($_.Exception.Message)" -Level ERROR -Method "METHOD-1"
            Write-Log "  Stack trace: $($_.ScriptStackTrace)" -Level DEBUG -Method "METHOD-1"

            # Log response details if available
            if ($_.Exception.Response)
            {
                Write-Log "  Response StatusCode: $($_.Exception.Response.StatusCode)" -Level DEBUG -Method "METHOD-1"
                try
                {
                    $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                    $responseBody = $reader.ReadToEnd()
                    Write-Log "  Response Body: $($responseBody.Substring(0, [Math]::Min(500, $responseBody.Length)))" -Level DEBUG -Method "METHOD-1"
                }
                catch { }
            }

            $downloadResult.Error = $_.Exception.Message
        }

        $result.Downloads += $downloadResult
    }

    $script:Results.Methods.Method1 = $result
    return $result
}
#endregion

#region METHOD 2: LEGACY GETDOWNLOADURL
function Test-Method2-GetDownloadUrl
{
    if (2 -in $SkipMethods)
    {
        Write-Log "Method 2 skipped by user" -Level WARNING -Method "METHOD-2"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 2: LEGACY GETDOWNLOADURL API"

    $result = @{
        Method = "Legacy getDownloadUrl API (deprecated)"
        Success = $false
        DownloadUrl = $null
        Error = $null
    }

    Write-Log "NOTE: This API was deprecated as of April 30, 2023" -Level WARNING -Method "METHOD-2"
    Write-Log "Testing if it still works in your tenant..." -Method "METHOD-2"

    try
    {
        # Try to call getDownloadUrl
        $uri = "https://graph.microsoft.com/beta/security/cases/ediscoveryCases/$script:CaseId/operations/$script:ExportOperationId/microsoft.graph.security.ediscoveryExportOperation/getDownloadUrl"
        Write-Log "Calling: $uri" -Level DEBUG -Method "METHOD-2"

        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        Write-Log "Response received:" -Level DEBUG -Method "METHOD-2"
        Write-Log ($response | ConvertTo-Json -Depth 5) -Level DEBUG -Method "METHOD-2"

        if ($response -and $response.value)
        {
            Write-Log "Download URL retrieved!" -Level SUCCESS -Method "METHOD-2"
            $result.DownloadUrl = $response.value
            Write-Log "URL (first 100 chars): $($response.value.Substring(0, [Math]::Min(100, $response.value.Length)))..." -Level DEBUG -Method "METHOD-2"

            # Try to download using this URL
            if ($script:PurviewToken)
            {
                $outputFile = Join-Path $OutputPath "Method2_download.zip"
                $headers = @{
                    'Authorization' = "Bearer $script:PurviewToken"
                    'X-AllowWithAADToken' = "true"
                }

                Write-Log "Attempting download to: $outputFile" -Method "METHOD-2"
                try
                {
                    Invoke-WebRequest -Uri $response.value -Headers $headers -OutFile $outputFile -ErrorAction Stop

                    if (Test-Path $outputFile)
                    {
                        $fileInfo = Get-Item $outputFile
                        Write-Log "File downloaded: $($fileInfo.Length) bytes" -Level SUCCESS -Method "METHOD-2"
                        $result.Success = $true
                        $result.FileSize = $fileInfo.Length
                    }
                }
                catch
                {
                    Write-Log "Download failed: $($_.Exception.Message)" -Level ERROR -Method "METHOD-2"
                    $result.Error = $_.Exception.Message
                }
            }
        }
        elseif ($response -is [string])
        {
            Write-Log "Download URL retrieved (string format)" -Level SUCCESS -Method "METHOD-2"
            $result.DownloadUrl = $response
        }
        else
        {
            Write-Log "getDownloadUrl returned empty response" -Level WARNING -Method "METHOD-2"
            $result.Error = "Empty response from getDownloadUrl"
        }
    }
    catch
    {
        Write-Log "getDownloadUrl API call failed: $($_.Exception.Message)" -Level ERROR -Method "METHOD-2"
        $result.Error = $_.Exception.Message

        if ($_.Exception.Message -match "deprecated|retired")
        {
            Write-Log "API appears to be fully deprecated" -Level WARNING -Method "METHOD-2"
        }
    }

    $script:Results.Methods.Method2 = $result
    return $result
}
#endregion

#region METHOD 3: AZURE BLOB DIRECT
function Test-Method3-AzureBlobDirect
{
    if (3 -in $SkipMethods)
    {
        Write-Log "Method 3 skipped by user" -Level WARNING -Method "METHOD-3"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 3: AZURE BLOB DIRECT (SAS URL)"

    $result = @{
        Method = "Direct Azure Blob download via SAS URL"
        Success = $false
        SasUrlFound = $false
        Error = $null
    }

    Write-Log "Checking if export has direct Azure blob SAS URL..." -Method "METHOD-3"
    Write-Log "This bypasses the eDiscovery proxy service" -Method "METHOD-3"

    # Look for SAS-style URLs in the download URLs
    $sasUrls = @()
    foreach ($url in $script:Results.ExportMetadata.DownloadUrls)
    {
        if ($url -match '\.blob\.' -and $url -match '\?.*sig=')
        {
            Write-Log "Found potential SAS URL" -Level SUCCESS -Method "METHOD-3"
            $sasUrls += $url
        }
    }

    if ($sasUrls.Count -eq 0)
    {
        Write-Log "No Azure blob SAS URLs found in export metadata" -Level WARNING -Method "METHOD-3"
        Write-Log "Your tenant may not support direct blob export" -Level WARNING -Method "METHOD-3"
        $result.Error = "No SAS URLs available"
        $script:Results.Methods.Method3 = $result
        return $result
    }

    $result.SasUrlFound = $true
    Write-Log "Found $($sasUrls.Count) SAS URL(s)" -Method "METHOD-3"

    foreach ($sasUrl in $sasUrls)
    {
        $outputFile = Join-Path $OutputPath "Method3_blob_download.zip"
        Write-Log "Attempting direct blob download..." -Method "METHOD-3"
        Write-Log "  Output: $outputFile" -Method "METHOD-3"

        try
        {
            # SAS URLs don't need auth headers - the signature is in the URL
            Invoke-WebRequest -Uri $sasUrl -OutFile $outputFile -ErrorAction Stop

            if (Test-Path $outputFile)
            {
                $fileInfo = Get-Item $outputFile
                Write-Log "Download successful: $($fileInfo.Length) bytes" -Level SUCCESS -Method "METHOD-3"
                $result.Success = $true
                $result.FileSize = $fileInfo.Length
            }
        }
        catch
        {
            Write-Log "Direct blob download failed: $($_.Exception.Message)" -Level ERROR -Method "METHOD-3"
            $result.Error = $_.Exception.Message
        }
    }

    $script:Results.Methods.Method3 = $result
    return $result
}
#endregion

#region METHOD 4: HTTPCLIENT CHUNKED
function Test-Method4-HttpClientChunked
{
    if (4 -in $SkipMethods)
    {
        Write-Log "Method 4 skipped by user" -Level WARNING -Method "METHOD-4"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 4: .NET HTTPCLIENT CHUNKED DOWNLOAD"

    $result = @{
        Method = ".NET HttpClient with streaming/chunked download"
        Success = $false
        Downloads = @()
        Error = $null
    }

    if (-not $script:PurviewToken)
    {
        Write-Log "No Purview token available" -Level ERROR -Method "METHOD-4"
        $result.Error = "No Purview token"
        $script:Results.Methods.Method4 = $result
        return $result
    }

    if (-not $script:Results.ExportMetadata.DownloadUrls -or $script:Results.ExportMetadata.DownloadUrls.Count -eq 0)
    {
        Write-Log "No download URLs available" -Level ERROR -Method "METHOD-4"
        $result.Error = "No download URLs"
        $script:Results.Methods.Method4 = $result
        return $result
    }

    Write-Log "Using System.Net.Http.HttpClient for streaming download..." -Method "METHOD-4"
    Write-Log "This handles large files better than Invoke-WebRequest" -Method "METHOD-4"

    Add-Type -AssemblyName System.Net.Http

    $downloadIndex = 0
    foreach ($url in $script:Results.ExportMetadata.DownloadUrls)
    {
        $downloadIndex++
        $outputFile = Join-Path $OutputPath "Method4_download_$downloadIndex.zip"

        Write-Log "Download $downloadIndex to: $outputFile" -Method "METHOD-4"

        $downloadResult = @{
            OutputFile = $outputFile
            Success = $false
            BytesDownloaded = 0
            Error = $null
        }

        $handler = $null
        $client = $null

        try
        {
            $handler = [System.Net.Http.HttpClientHandler]::new()
            $handler.AllowAutoRedirect = $true

            $client = [System.Net.Http.HttpClient]::new($handler)
            $client.Timeout = [System.TimeSpan]::FromMinutes(30)

            # Set headers
            $client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $script:PurviewToken)
            $client.DefaultRequestHeaders.Add("X-AllowWithAADToken", "true")

            Write-Log "  Sending request..." -Method "METHOD-4"
            $response = $client.GetAsync($url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result

            Write-Log "  StatusCode: $($response.StatusCode)" -Method "METHOD-4"
            Write-Log "  Content-Length: $($response.Content.Headers.ContentLength)" -Method "METHOD-4"

            if ($response.IsSuccessStatusCode)
            {
                $stream = $response.Content.ReadAsStreamAsync().Result
                $fileStream = [System.IO.File]::Create($outputFile)

                $buffer = New-Object byte[] 81920
                $totalBytesRead = 0
                $lastProgress = 0

                Write-Log "  Downloading with chunked streaming..." -Method "METHOD-4"

                try
                {
                    while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0)
                    {
                        $fileStream.Write($buffer, 0, $bytesRead)
                        $totalBytesRead += $bytesRead

                        # Progress every 10MB
                        $progressMB = [Math]::Floor($totalBytesRead / 10MB)
                        if ($progressMB -gt $lastProgress)
                        {
                            Write-Log "    Downloaded: $([Math]::Round($totalBytesRead / 1MB, 2)) MB" -Level DEBUG -Method "METHOD-4"
                            $lastProgress = $progressMB
                        }
                    }
                }
                finally
                {
                    $fileStream.Close()
                    $stream.Close()
                }

                Write-Log "  Total downloaded: $([Math]::Round($totalBytesRead / 1MB, 2)) MB" -Method "METHOD-4"
                $downloadResult.BytesDownloaded = $totalBytesRead

                if (Test-Path $outputFile)
                {
                    $fileInfo = Get-Item $outputFile
                    if ($fileInfo.Length -gt 1000)
                    {
                        Write-Log "  Download successful!" -Level SUCCESS -Method "METHOD-4"
                        $downloadResult.Success = $true
                        $result.Success = $true
                    }
                    else
                    {
                        Write-Log "  File too small - may be error" -Level WARNING -Method "METHOD-4"
                    }
                }
            }
            else
            {
                $errorContent = $response.Content.ReadAsStringAsync().Result
                Write-Log "  Request failed: $($response.StatusCode)" -Level ERROR -Method "METHOD-4"
                Write-Log "  Response: $($errorContent.Substring(0, [Math]::Min(500, $errorContent.Length)))" -Level DEBUG -Method "METHOD-4"
                $downloadResult.Error = "HTTP $($response.StatusCode)"
            }
        }
        catch
        {
            Write-Log "  Error: $($_.Exception.Message)" -Level ERROR -Method "METHOD-4"
            $downloadResult.Error = $_.Exception.Message
        }
        finally
        {
            if ($client) { $client.Dispose() }
            if ($handler) { $handler.Dispose() }
        }

        $result.Downloads += $downloadResult
    }

    $script:Results.Methods.Method4 = $result
    return $result
}
#endregion

#region METHOD 5: BITS TRANSFER
function Test-Method5-BitsTransfer
{
    if (5 -in $SkipMethods)
    {
        Write-Log "Method 5 skipped by user" -Level WARNING -Method "METHOD-5"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 5: BITS TRANSFER"

    $result = @{
        Method = "Background Intelligent Transfer Service (BITS)"
        Success = $false
        BitsSupported = $false
        Error = $null
    }

    Write-Log "Testing BITS transfer for more reliable large file downloads..." -Method "METHOD-5"
    Write-Log "BITS can resume interrupted downloads automatically" -Method "METHOD-5"

    # Check if BITS module is available
    if (-not (Get-Module -ListAvailable -Name BitsTransfer))
    {
        Write-Log "BitsTransfer module not available" -Level WARNING -Method "METHOD-5"
        $result.Error = "BitsTransfer module not available"
        $script:Results.Methods.Method5 = $result
        return $result
    }

    $result.BitsSupported = $true
    Import-Module BitsTransfer -ErrorAction SilentlyContinue

    if (-not $script:Results.ExportMetadata.DownloadUrls -or $script:Results.ExportMetadata.DownloadUrls.Count -eq 0)
    {
        Write-Log "No download URLs available" -Level ERROR -Method "METHOD-5"
        $result.Error = "No download URLs"
        $script:Results.Methods.Method5 = $result
        return $result
    }

    # BITS doesn't support custom headers directly, so this method may not work
    # but we'll try it anyway as some configurations might allow it
    Write-Log "NOTE: BITS may not support custom auth headers required by eDiscovery" -Level WARNING -Method "METHOD-5"

    $downloadIndex = 0
    foreach ($url in $script:Results.ExportMetadata.DownloadUrls)
    {
        $downloadIndex++
        $outputFile = Join-Path $OutputPath "Method5_bits_$downloadIndex.zip"

        Write-Log "Attempting BITS transfer to: $outputFile" -Method "METHOD-5"

        try
        {
            # Try direct BITS transfer (may work if URL has embedded auth like SAS)
            $bitsJob = Start-BitsTransfer -Source $url -Destination $outputFile -Asynchronous -ErrorAction Stop

            Write-Log "  BITS job created: $($bitsJob.JobId)" -Method "METHOD-5"

            # Wait for completion with timeout
            $timeout = 300  # 5 minutes for test
            $elapsed = 0
            while ($bitsJob.JobState -eq "Transferring" -or $bitsJob.JobState -eq "Connecting" -and $elapsed -lt $timeout)
            {
                $percent = if ($bitsJob.BytesTotal -gt 0) { [Math]::Round(($bitsJob.BytesTransferred / $bitsJob.BytesTotal) * 100, 2) } else { 0 }
                Write-Log "    Progress: $percent% ($($bitsJob.BytesTransferred) / $($bitsJob.BytesTotal) bytes)" -Level DEBUG -Method "METHOD-5"
                Start-Sleep -Seconds 5
                $elapsed += 5
            }

            if ($bitsJob.JobState -eq "Transferred")
            {
                Complete-BitsTransfer -BitsJob $bitsJob
                Write-Log "  BITS transfer completed!" -Level SUCCESS -Method "METHOD-5"

                if (Test-Path $outputFile)
                {
                    $fileInfo = Get-Item $outputFile
                    Write-Log "  File size: $($fileInfo.Length) bytes" -Method "METHOD-5"
                    $result.Success = $true
                }
            }
            else
            {
                Write-Log "  BITS job ended with state: $($bitsJob.JobState)" -Level WARNING -Method "METHOD-5"
                if ($bitsJob.Error)
                {
                    Write-Log "  Error: $($bitsJob.Error.Description)" -Level ERROR -Method "METHOD-5"
                }
                Remove-BitsTransfer -BitsJob $bitsJob -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            Write-Log "  BITS transfer failed: $($_.Exception.Message)" -Level ERROR -Method "METHOD-5"
            $result.Error = $_.Exception.Message
        }
    }

    $script:Results.Methods.Method5 = $result
    return $result
}
#endregion

#region METHOD 6: BROWSER TOKEN GUIDANCE
function Test-Method6-BrowserTokenGuidance
{
    if (6 -in $SkipMethods)
    {
        Write-Log "Method 6 skipped by user" -Level WARNING -Method "METHOD-6"
        return @{ Skipped = $true }
    }

    Write-LogSection "METHOD 6: BROWSER TOKEN CAPTURE GUIDANCE"

    $result = @{
        Method = "Manual browser token capture"
        GuidanceProvided = $true
        Success = $false
    }

    Write-Log "If automated methods fail, you can capture a token from the browser..." -Method "METHOD-6"
    Write-Log "" -Method "METHOD-6"
    Write-Log "STEPS TO CAPTURE TOKEN MANUALLY:" -Level WARNING -Method "METHOD-6"
    Write-Log "1. Open Microsoft Purview portal: https://compliance.microsoft.com" -Method "METHOD-6"
    Write-Log "2. Navigate to your eDiscovery case" -Method "METHOD-6"
    Write-Log "3. Open browser Developer Tools (F12)" -Method "METHOD-6"
    Write-Log "4. Go to Network tab" -Method "METHOD-6"
    Write-Log "5. Click 'Download' on an export" -Method "METHOD-6"
    Write-Log "6. Look for a request to *.proxyservice.ediscovery.*" -Method "METHOD-6"
    Write-Log "7. Copy the 'Authorization' header value (after 'Bearer ')" -Method "METHOD-6"
    Write-Log "8. Use the token with this script or Invoke-WebRequest" -Method "METHOD-6"
    Write-Log "" -Method "METHOD-6"
    Write-Log "Example command with captured token:" -Method "METHOD-6"
    Write-Log @"
`$headers = @{
    'Authorization' = 'Bearer YOUR_CAPTURED_TOKEN_HERE'
    'X-AllowWithAADToken' = 'true'
}
`$url = 'YOUR_DOWNLOAD_URL_HERE'
Invoke-WebRequest -Uri `$url -Headers `$headers -OutFile 'download.zip'
"@ -Level DEBUG -Method "METHOD-6"

    # If we have download URLs, provide them
    if ($script:Results.ExportMetadata.DownloadUrls.Count -gt 0)
    {
        Write-Log "" -Method "METHOD-6"
        Write-Log "Download URLs from your export:" -Method "METHOD-6"
        foreach ($url in $script:Results.ExportMetadata.DownloadUrls)
        {
            Write-Log "  $($url.Substring(0, [Math]::Min(150, $url.Length)))..." -Method "METHOD-6"
        }
    }

    $script:Results.Methods.Method6 = $result
    return $result
}
#endregion

#region NETWORK DIAGNOSTICS
function Test-NetworkDiagnostics
{
    Write-LogSection "NETWORK DIAGNOSTICS"

    $netResults = @{
        ProxyEndpoints = @{}
        TlsVersion = $null
        SystemProxy = $null
    }

    # Check system proxy settings
    Write-Log "Checking system proxy settings..." -Method "NETWORK"
    try
    {
        $proxySettings = [System.Net.WebRequest]::DefaultWebProxy
        if ($proxySettings)
        {
            $testUri = [System.Uri]"https://compliance.microsoft.com"
            $proxyUri = $proxySettings.GetProxy($testUri)
            if ($proxyUri -ne $testUri)
            {
                Write-Log "  System proxy detected: $proxyUri" -Level WARNING -Method "NETWORK"
                $netResults.SystemProxy = $proxyUri.ToString()
            }
            else
            {
                Write-Log "  No proxy for Microsoft endpoints" -Method "NETWORK"
                $netResults.SystemProxy = "None"
            }
        }
    }
    catch
    {
        Write-Log "  Could not check proxy: $($_.Exception.Message)" -Level WARNING -Method "NETWORK"
    }

    # Check TLS version
    Write-Log "Checking TLS configuration..." -Method "NETWORK"
    $netResults.TlsVersion = [System.Net.ServicePointManager]::SecurityProtocol.ToString()
    Write-Log "  Security Protocol: $($netResults.TlsVersion)" -Method "NETWORK"

    # Test connectivity to proxy endpoints
    Write-Log "Testing connectivity to eDiscovery proxy endpoints..." -Method "NETWORK"
    foreach ($endpoint in $script:ProxyEndpoints)
    {
        Write-Log "  Testing $($endpoint.Region): $($endpoint.Url)" -Method "NETWORK"

        $endpointResult = @{
            Region = $endpoint.Region
            Url = $endpoint.Url
            Reachable = $false
            StatusCode = $null
            ResponseHeaders = @{}
            Error = $null
        }

        try
        {
            # Just test connectivity, not auth
            $response = Invoke-WebRequest -Uri $endpoint.Url -Method HEAD -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
            Write-Log "    Reachable: StatusCode $($response.StatusCode)" -Level SUCCESS -Method "NETWORK"
            $endpointResult.Reachable = $true
            $endpointResult.StatusCode = $response.StatusCode
        }
        catch [System.Net.WebException]
        {
            $webEx = $_.Exception
            if ($webEx.Response)
            {
                $statusCode = [int]$webEx.Response.StatusCode
                Write-Log "    Response: HTTP $statusCode (endpoint exists but returned error)" -Level WARNING -Method "NETWORK"
                $endpointResult.StatusCode = $statusCode
                $endpointResult.Reachable = $true  # Endpoint is reachable even if auth fails

                # Check for SSL inspection signatures
                foreach ($header in $webEx.Response.Headers.AllKeys)
                {
                    $endpointResult.ResponseHeaders[$header] = $webEx.Response.Headers[$header]
                    if ($header -match "X-.*-Proxy|Via|X-Forwarded")
                    {
                        Write-Log "    Possible proxy header: $header = $($webEx.Response.Headers[$header])" -Level WARNING -Method "NETWORK"
                    }
                }
            }
            else
            {
                Write-Log "    Not reachable: $($_.Exception.Message)" -Level ERROR -Method "NETWORK"
                $endpointResult.Error = $_.Exception.Message
            }
        }
        catch
        {
            Write-Log "    Error: $($_.Exception.Message)" -Level ERROR -Method "NETWORK"
            $endpointResult.Error = $_.Exception.Message
        }

        $netResults.ProxyEndpoints[$endpoint.Region] = $endpointResult
    }

    $script:Results.NetworkDiagnostics = $netResults
    return $netResults
}
#endregion

#region RESULTS SUMMARY
function Write-ResultsSummary
{
    Write-LogSection "RESULTS SUMMARY"

    $recommendations = @()

    # Summarize each method
    Write-Log "Method Results:" -Method "SUMMARY"

    $methodResults = @(
        @{ Num = 1; Name = "Invoke-WebRequest + exportFileMetadata"; Result = $script:Results.Methods.Method1 }
        @{ Num = 2; Name = "Legacy getDownloadUrl API"; Result = $script:Results.Methods.Method2 }
        @{ Num = 3; Name = "Azure Blob Direct"; Result = $script:Results.Methods.Method3 }
        @{ Num = 4; Name = "HttpClient Chunked"; Result = $script:Results.Methods.Method4 }
        @{ Num = 5; Name = "BITS Transfer"; Result = $script:Results.Methods.Method5 }
        @{ Num = 6; Name = "Browser Token Guidance"; Result = $script:Results.Methods.Method6 }
    )

    foreach ($method in $methodResults)
    {
        if ($method.Result.Skipped)
        {
            Write-Log "  Method $($method.Num): $($method.Name) - SKIPPED" -Method "SUMMARY"
        }
        elseif ($method.Result.Success)
        {
            Write-Log "  Method $($method.Num): $($method.Name) - SUCCESS" -Level SUCCESS -Method "SUMMARY"
        }
        else
        {
            Write-Log "  Method $($method.Num): $($method.Name) - FAILED" -Level WARNING -Method "SUMMARY"
            if ($method.Result.Error)
            {
                Write-Log "    Error: $($method.Result.Error)" -Level DEBUG -Method "SUMMARY"
            }
        }
    }

    # Generate recommendations based on results
    Write-Log "" -Method "SUMMARY"
    Write-Log "Recommendations:" -Method "SUMMARY"

    if (-not $script:Results.Methods.TokenAcquisition.Success)
    {
        $recommendations += "Token acquisition failed - verify app registration has eDiscovery.Download.Read permission with admin consent"
        Write-Log "  - Fix token acquisition first - all download methods depend on it" -Level WARNING -Method "SUMMARY"
    }

    if ($script:Results.ExportMetadata.DownloadUrls.Count -eq 0)
    {
        $recommendations += "No download URLs found - verify export is complete and you have case membership"
        Write-Log "  - No download URLs found in export metadata" -Level WARNING -Method "SUMMARY"
    }

    if ($script:Results.NetworkDiagnostics.SystemProxy -and $script:Results.NetworkDiagnostics.SystemProxy -ne "None")
    {
        $recommendations += "System proxy detected - may need to configure proxy bypass for eDiscovery endpoints"
        Write-Log "  - System proxy detected - may interfere with downloads" -Level WARNING -Method "SUMMARY"
    }

    $successfulMethods = $methodResults | Where-Object { $_.Result.Success }
    if ($successfulMethods)
    {
        Write-Log "  - Working method(s) found: $($successfulMethods.Name -join ', ')" -Level SUCCESS -Method "SUMMARY"
        $recommendations += "Use Method $($successfulMethods[0].Num) ($($successfulMethods[0].Name)) for production downloads"
    }
    else
    {
        Write-Log "  - No automated download method succeeded" -Level ERROR -Method "SUMMARY"
        $recommendations += "Try browser token capture (Method 6) as a workaround"
        $recommendations += "Share log file for further analysis"
    }

    $script:Results.Recommendations = $recommendations

    # Save results to JSON
    Write-Log "" -Method "SUMMARY"
    Write-Log "Saving detailed results to: $script:ResultsFile" -Method "SUMMARY"
    $script:Results | ConvertTo-Json -Depth 10 | Out-File $script:ResultsFile -Encoding UTF8

    Write-Log "" -Method "SUMMARY"
    Write-Log "Output files:" -Method "SUMMARY"
    Write-Log "  Log file: $script:LogFile" -Method "SUMMARY"
    Write-Log "  Results JSON: $script:ResultsFile" -Method "SUMMARY"
    Write-Log "" -Method "SUMMARY"
    Write-Log "Please share the log file for troubleshooting assistance." -Level WARNING -Method "SUMMARY"
}
#endregion

#region MAIN EXECUTION
try
{
    Write-Log "============================================================"
    Write-Log "eDiscovery PST Download Test Script"
    Write-Log "============================================================"
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    # Store parameters in script scope for access by functions
    $script:CaseId = $CaseId
    $script:CaseName = $CaseName
    $script:TenantId = $TenantId

    if ($CaseId)
    {
        Write-Log "Case ID: $CaseId"
    }
    elseif ($CaseName)
    {
        Write-Log "Case Name: $CaseName (will lookup ID)"
    }
    else
    {
        throw "Either -CaseId or -CaseName parameter is required"
    }

    Write-Log "Output Path: $OutputPath"
    Write-Log "App ID: $AppId"
    Write-Log "Tenant ID: $TenantId"
    if ($ClientSecret) { Write-Log "Auth: Client Secret" } elseif ($CertificateThumbprint) { Write-Log "Auth: Certificate" }
    Write-Log "Skip Methods: $(if ($SkipMethods) { $SkipMethods -join ', ' } else { 'None' })"
    Write-Log ""

    # Phase 1: Prerequisites
    Test-Prerequisites

    # Phase 2: Graph Connection
    $graphResult = Connect-ToGraph
    if (-not $graphResult.Success)
    {
        throw "Failed to connect to Microsoft Graph"
    }

    # Check service principal
    Test-PurviewServicePrincipal

    # Phase 3: Get Export Metadata
    $metadataResult = Get-ExportMetadata
    if (-not $metadataResult.CaseFound)
    {
        throw "Failed to get export metadata: $($metadataResult.Error)"
    }

    # Phase 4: Token Acquisition
    $tokenResult = Get-PurviewDownloadToken

    # Phase 5: Test Download Methods
    Test-Method1-InvokeWebRequest
    Test-Method2-GetDownloadUrl
    Test-Method3-AzureBlobDirect
    Test-Method4-HttpClientChunked
    Test-Method5-BitsTransfer
    Test-Method6-BrowserTokenGuidance

    # Phase 6: Network Diagnostics
    Test-NetworkDiagnostics

    # Summary
    Write-ResultsSummary
}
catch
{
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level ERROR -Method "MAIN"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG -Method "MAIN"

    $script:Results.FatalError = $_.Exception.Message
    $script:Results | ConvertTo-Json -Depth 10 | Out-File $script:ResultsFile -Encoding UTF8 -ErrorAction SilentlyContinue
}
finally
{
    Write-Log ""
    Write-Log "Script completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Log "============================================================"
}
#endregion
