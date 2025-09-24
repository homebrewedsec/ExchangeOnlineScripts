<#
.SYNOPSIS
Exports detailed reports of Office 365 mailboxes with delegation permissions and forwarding configurations.

.DESCRIPTION
This script generates comprehensive reports of mailbox delegation permissions (Full Access, Send As, Send on Behalf)
and forwarding configurations. Supports filtering by CSV input or mailbox types. Outputs CSV reports to the scriptlogs directory.

.PARAMETER InputCsvPath
Optional path to CSV file containing UPN column to filter specific mailboxes for reporting.
If provided, only mailboxes specified in the CSV will be processed regardless of type.

.PARAMETER OutputPath
Directory path where the CSV report will be saved. Defaults to current working directory.

.PARAMETER IncludeUsers
Switch to include user mailboxes in addition to shared mailboxes. Ignored if InputCsvPath is provided.

.PARAMETER IncludeOrgData
Switch to include organizational data (Manager, Office, Department) from Active Directory or Microsoft Graph.
When not specified, these fields will be empty, making the script faster and simpler.

.PARAMETER ReverseLookup
Switch to perform reverse lookup mode. When combined with InputCsvPath, finds all mailboxes, distribution groups, and dynamic
distribution groups that the specified users have permissions on or receive forwards from. Output shows each user and what
objects they have access to. For distribution groups, checks AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromSendersOrMembers,
GrantSendOnBehalfTo, and ModeratedBy permissions.

.PARAMETER SendAsMethod
Specifies which cmdlet to use first for retrieving Send As permissions. Valid values are "MailboxPermission" (default) or
"RecipientPermission". Different Exchange Online environments may require different cmdlets. The script will try the
specified cmdlet first, then fallback to the alternative if the first fails.

.PARAMETER Verbose
Provides detailed output during execution

.EXAMPLE
.\Invoke-MailboxDelegationReport.ps1
Exports delegation report for all shared, room, and equipment mailboxes to current directory.

.EXAMPLE
.\Invoke-MailboxDelegationReport.ps1 -InputCsvPath "C:\temp\targetmailboxes.csv" -OutputPath "C:\Reports"
Exports delegation report for mailboxes specified in the CSV file to the C:\Reports directory.

.EXAMPLE
.\Invoke-MailboxDelegationReport.ps1 -IncludeUsers -OutputPath "C:\Powershell\scriptlogs"
Exports delegation report for shared, room, equipment, and user mailboxes to specified directory.

.EXAMPLE
.\Invoke-MailboxDelegationReport.ps1 -IncludeOrgData
Exports delegation report with organizational data (Manager, Office, Department) included.

.EXAMPLE
.\Invoke-MailboxDelegationReport.ps1 -InputCsvPath "C:\temp\users.csv" -ReverseLookup
Performs reverse lookup to find all mailboxes, distribution groups, and dynamic distribution groups that users in the CSV have permissions on.

.NOTES
Author: Hudson Bush, Seguri - hudson@seguri.io
Requires: Exchange Online Management module (required), Active Directory or Microsoft Graph modules (optional for -IncludeOrgData)
Prerequisites: User must run Connect-ExchangeOnline first. For organizational data, AD module or Connect-MgGraph required.
Output: CSV file saved to specified output directory (default: current directory)
#>

[CmdletBinding()]
param(
    [string]$InputCsvPath,
    [string]$OutputPath = (Get-Location).Path,
    [switch]$IncludeUsers,
    [switch]$IncludeOrgData,
    [switch]$ReverseLookup,
    [ValidateSet("MailboxPermission", "RecipientPermission")]
    [string]$SendAsMethod = "MailboxPermission"
)


# Function to resolve GUID to user identity
function Resolve-UserIdentity {
    param(
        [string]$Identity,
        [hashtable]$IdentityCache = @{}
    )

    # Return if already cached
    if ($IdentityCache.ContainsKey($Identity)) {
        return $IdentityCache[$Identity]
    }

    # If it's already a UPN or display name (contains @ or doesn't look like GUID), return as-is
    if ($Identity -match '@' -or $Identity -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        $IdentityCache[$Identity] = $Identity
        return $Identity
    }

    # Try to resolve GUID to user
    try {
        # Try as mailbox first
        $user = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
        if ($user) {
            $resolvedIdentity = $user.UserPrincipalName
            $IdentityCache[$Identity] = $resolvedIdentity
            return $resolvedIdentity
        }

        # Try as mail user
        $mailUser = Get-MailUser -Identity $Identity -ErrorAction SilentlyContinue
        if ($mailUser) {
            $resolvedIdentity = $mailUser.UserPrincipalName
            $IdentityCache[$Identity] = $resolvedIdentity
            return $resolvedIdentity
        }

        # Try as recipient (broader catch-all)
        $recipient = Get-Recipient -Identity $Identity -ErrorAction SilentlyContinue
        if ($recipient) {
            $resolvedIdentity = if ($recipient.PrimarySmtpAddress) { $recipient.PrimarySmtpAddress } else { $recipient.DisplayName }
            $IdentityCache[$Identity] = $resolvedIdentity
            return $resolvedIdentity
        }
    }
    catch {
        Write-Verbose "Could not resolve identity: $Identity - $($_.Exception.Message)"
    }

    # If resolution fails, return original identity
    $IdentityCache[$Identity] = $Identity
    return $Identity
}

# Generate output file path
$ReportType = if ($ReverseLookup) { "MailboxesWithSpecifiedDelegates" } else { "MailboxDelegationReport" }
$OutputFile = Join-Path $OutputPath "${ReportType}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Determine mailbox types and mode based on parameters
if ($ReverseLookup)
{
    if (-not $InputCsvPath)
    {
        Write-Error "ReverseLookup mode requires InputCsvPath parameter with user UPNs."
        exit 1
    }
    $MailboxTypes = @("SharedMailbox", "RoomMailbox", "EquipmentMailbox") # Always get all shared/room/equipment for reverse lookup
    Write-Output "Starting Reverse Lookup Report..."
    Write-Output "Finding mailboxes that users from CSV have permissions on: $InputCsvPath"
}
elseif ($InputCsvPath)
{
    $MailboxTypes = @() # Will be determined from CSV content
    Write-Output "Starting Mailbox Delegation Report Export..."
    Write-Output "Processing mailboxes from CSV: $InputCsvPath"
}
else
{
    # Default: Shared, Room, and Equipment mailboxes
    $MailboxTypes = @("SharedMailbox", "RoomMailbox", "EquipmentMailbox")

    # Add user mailboxes if requested
    if ($IncludeUsers)
    {
        $MailboxTypes += "UserMailbox"
    }

    Write-Output "Starting Mailbox Delegation Report Export..."
    Write-Output "Target Mailbox Types: $($MailboxTypes -join ', ')"
}

try
{
    # Check for existing Exchange Online session
    Write-Output "Checking Exchange Online connection..."
    try
    {
        $sessionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $sessionInfo -or $sessionInfo.State -ne "Connected")
        {
            Write-Error "No active Exchange Online session found. Please run Connect-ExchangeOnline first."
            Write-Output "Example: Connect-ExchangeOnline -UserPrincipalName youradmin@domain.com"
            exit 1
        }
        else
        {
            Write-Output "Found active Exchange Online session: $($sessionInfo.Name)"
        }
    }
    catch
    {
        Write-Error "Exchange Online module not available or no active session. Please run Connect-ExchangeOnline first."
        exit 1
    }

    # Initialize organizational data variables
    $ADUsers = @()
    $script:SkipOrgData = -not $IncludeOrgData

    # Check for organizational data modules only if requested
    if ($IncludeOrgData)
    {
        Write-Output "Checking for organizational data sources..."

        # Try Active Directory first
        try
        {
            Import-Module ActiveDirectory -ErrorAction Stop
            Write-Output "Active Directory module loaded successfully"

            Write-Output "Retrieving Active Directory user data..."
            $ADUsers = Get-ADUser -Filter * -Properties Manager, DisplayName, Office, Department, Title
            Write-Output "Retrieved $($ADUsers.Count) AD users"
        }
        catch
        {
            Write-Output "Active Directory module not available. Trying Microsoft Graph..."

            # Try Microsoft Graph as fallback
            try
            {
                $graphContext = Get-MgContext -ErrorAction SilentlyContinue
                if (-not $graphContext)
                {
                    Write-Output "No active Microsoft Graph session found. Please run Connect-MgGraph first."
                    Write-Output "Organizational data will not be included in the report."
                    $script:SkipOrgData = $true
                }
                else
                {
                    Write-Output "Using Microsoft Graph for organizational data..."
                    # Note: Graph user retrieval would go here if implemented
                    Write-Output "Microsoft Graph organizational data retrieval not yet implemented."
                    $script:SkipOrgData = $true
                }
            }
            catch
            {
                Write-Output "Neither Active Directory nor Microsoft Graph available for organizational data."
                Write-Output "Organizational fields (Manager, Office, Department) will be empty."
                $script:SkipOrgData = $true
            }
        }
    }
    else
    {
        Write-Verbose "Organizational data not requested (-IncludeOrgData not specified)"
    }

    # Determine mailbox collection approach
    $targetMailboxes = @()
    $targetUserUPNs = @()

    if ($ReverseLookup)
    {
        # For reverse lookup, get user UPNs from CSV and all shared mailboxes
        Write-Output "Processing CSV input file for user UPNs: $InputCsvPath"
        $csvData = Import-Csv $InputCsvPath

        if (-not $csvData -or $csvData.Count -eq 0)
        {
            Write-Error "CSV file is empty or could not be read: $InputCsvPath"
            exit 1
        }

        if (-not ($csvData | Get-Member -Name "upn" -MemberType NoteProperty))
        {
            Write-Error "CSV file must contain a 'upn' column with user principal names."
            exit 1
        }

        $targetUserUPNs = $csvData | Select-Object -ExpandProperty upn
        Write-Output "Found $($targetUserUPNs.Count) user UPNs in CSV file"
        Write-Output "Target UPNs: $($targetUserUPNs -join ', ')"

        # Get all mailboxes and distribution groups for reverse lookup
        Write-Output "Retrieving all mailboxes for reverse lookup..."
        $targetMailboxes = Get-Mailbox -ResultSize Unlimited
        Write-Output "Retrieved $($targetMailboxes.Count) mailboxes"

        Write-Output "Retrieving all distribution groups for reverse lookup..."
        $targetDistributionGroups = @(Get-DistributionGroup -ResultSize Unlimited)
        $dgText = if ($targetDistributionGroups.Count -eq 1) { "distribution group" } else { "distribution groups" }
        Write-Output "Retrieved $($targetDistributionGroups.Count) $dgText"

        Write-Output "Retrieving all dynamic distribution groups for reverse lookup..."
        $targetDynamicDistributionGroups = @(Get-DynamicDistributionGroup -ResultSize Unlimited)
        $ddgText = if ($targetDynamicDistributionGroups.Count -eq 1) { "dynamic distribution group" } else { "dynamic distribution groups" }
        Write-Output "Retrieved $($targetDynamicDistributionGroups.Count) $ddgText"
    }
    elseif ($InputCsvPath -and (Test-Path $InputCsvPath))
    {
        Write-Output "Processing CSV input file: $InputCsvPath"
        $csvData = Import-Csv $InputCsvPath

        if (-not ($csvData | Get-Member -Name "upn" -MemberType NoteProperty))
        {
            Write-Error "CSV file must contain a 'upn' column with mailbox user principal names."
            exit 1
        }

        $upnList = $csvData | Select-Object -ExpandProperty upn
        Write-Output "Found $($upnList.Count) UPNs in CSV file"

        # Get mailboxes matching UPNs from CSV
        foreach ($upn in $upnList)
        {
            try
            {
                $mailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
                if ($mailbox)
                {
                    $targetMailboxes += $mailbox
                }
            }
            catch
            {
                Write-Warning "Could not find mailbox for UPN: $upn"
            }
        }
    }
    else
    {
        Write-Output "Retrieving mailboxes by type filter..."

        # Build filter string for mailbox types
        if ($MailboxTypes.Count -eq 1)
        {
            $filter = "RecipientTypeDetails -eq '$($MailboxTypes[0])'"
        }
        else
        {
            $typeFilters = $MailboxTypes | ForEach-Object { "RecipientTypeDetails -eq '$_'" }
            $filter = "($($typeFilters -join ' -or '))"
        }

        $targetMailboxes = Get-Mailbox -ResultSize Unlimited -Filter $filter
    }

    $mbText = if ($targetMailboxes.Count -eq 1) { "mailbox" } else { "mailboxes" }
    Write-Output "Processing $($targetMailboxes.Count) $mbText..."

    # For reverse lookup, cache all permissions upfront to reduce API calls
    # For standard lookup, use individual calls since we're processing fewer mailboxes
    if ($ReverseLookup)
    {
        Write-Output "Caching all mailbox permissions to improve performance..."
    $mbFullAccessText = if ($targetMailboxes.Count -eq 1) { "mailbox" } else { "mailboxes" }
    Write-Output "Retrieving Full Access permissions for $($targetMailboxes.Count) $mbFullAccessText..."
    $allFullAccessPermissions = @{}
    $permissionCount = 0
    foreach ($mailbox in $targetMailboxes) {
        try {
            $fullAccessPerms = Get-MailboxPermission $mailbox.PrimarySmtpAddress |
                Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
            if ($fullAccessPerms) {
                $allFullAccessPermissions[$mailbox.PrimarySmtpAddress] = $fullAccessPerms
                $allFullAccessPermissions[$mailbox.Identity] = $fullAccessPerms
                # Handle single permission vs array
                if ($fullAccessPerms -is [array]) {
                    $permissionCount += $fullAccessPerms.Count
                } else {
                    $permissionCount += 1
                }
            }
        }
        catch {
            Write-Warning "Could not get Full Access permissions for $($mailbox.DisplayName): $($_.Exception.Message)"
        }
        Write-Progress -Activity "Caching Full Access Permissions" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($targetMailboxes.IndexOf($mailbox) / $targetMailboxes.Count) * 100)
    }
    Write-Output "Cached $permissionCount Full Access permissions"

    $mbSendAsText = if ($targetMailboxes.Count -eq 1) { "mailbox" } else { "mailboxes" }
    Write-Output "Retrieving Send As permissions for $($targetMailboxes.Count) $mbSendAsText..."
    $allSendAsPermissions = @{}
    $sendAsCount = 0
    foreach ($mailbox in $targetMailboxes) {
        try {
            if ($SendAsMethod -eq "MailboxPermission") {
                $sendAsPerms = Get-MailboxPermission $mailbox.PrimarySmtpAddress |
                    Where-Object { $_.AccessRights -eq 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
            }
            else {
                $sendAsPerms = Get-RecipientPermission $mailbox.PrimarySmtpAddress |
                    Where-Object { $_.AccessRights -eq 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
            }

            if ($sendAsPerms) {
                $allSendAsPermissions[$mailbox.PrimarySmtpAddress] = $sendAsPerms
                $allSendAsPermissions[$mailbox.Identity] = $sendAsPerms
                # Handle single permission vs array
                if ($sendAsPerms -is [array]) {
                    $sendAsCount += $sendAsPerms.Count
                } else {
                    $sendAsCount += 1
                }
            }
        }
        catch {
            Write-Warning "Could not get Send As permissions for $($mailbox.DisplayName): $($_.Exception.Message)"
        }
        Write-Progress -Activity "Caching Send As Permissions" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($targetMailboxes.IndexOf($mailbox) / $targetMailboxes.Count) * 100)
    }
    Write-Output "Cached $sendAsCount Send As permissions"
    }
    else
    {
        # For standard lookup, initialize empty cache (permissions will be retrieved individually)
        $allFullAccessPermissions = @{}
        $allSendAsPermissions = @{}
    }

    # Initialize results array and identity cache
    $results = @()
    $script:IdentityCache = @{}

    # Process each mailbox
    foreach ($mailbox in $targetMailboxes)
    {
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($results.Count / $targetMailboxes.Count) * 100)

        try
        {
            # Reset variables for each mailbox
            $fullAccessDelegates = @()
            $sendAsDelegates = @()
            $sendOnBehalfDelegates = @()
            $forwardingRules = @()
            $manager = ""
            $adUser = $null

            # Get AD user information (only if organizational data requested)
            $adUser = $null
            if (-not $script:SkipOrgData)
            {
                $adUser = $ADUsers | Where-Object { $_.UserPrincipalName -eq $mailbox.UserPrincipalName }
                if ($adUser -and $adUser.Manager)
                {
                    $managerUser = $ADUsers | Where-Object { $_.DistinguishedName -eq $adUser.Manager }
                    if ($managerUser)
                    {
                        $manager = $managerUser.DisplayName
                    }
                }
            }

            # Get delegation permissions (from cache in reverse lookup mode, or directly for standard mode)
            if ($ReverseLookup)
            {
                Write-Verbose "Getting cached permissions for $($mailbox.DisplayName)"
                # Get Full Access delegates from cache
                $fullAccessDelegates = @()
                $mailboxKeys = @($mailbox.PrimarySmtpAddress, $mailbox.Identity, $mailbox.DistinguishedName, $mailbox.ExchangeGuid)
                foreach ($key in $mailboxKeys) {
                    if ($allFullAccessPermissions[$key]) {
                        $fullAccessDelegates = $allFullAccessPermissions[$key]
                        break
                    }
                }
                # Get Send As delegates from cache
                $sendAsDelegates = @()
                foreach ($key in $mailboxKeys) {
                    if ($allSendAsPermissions[$key]) {
                        $sendAsDelegates = $allSendAsPermissions[$key]
                        break
                    }
                }
            }
            else
            {
                Write-Verbose "Getting permissions directly for $($mailbox.DisplayName)"

                # Get Full Access delegates directly
                try {
                    $fullAccessDelegates = Get-MailboxPermission $mailbox.PrimarySmtpAddress |
                        Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
                }
                catch {
                    Write-Warning "Could not retrieve Full Access permissions for $($mailbox.DisplayName): $($_.Exception.Message)"
                    $fullAccessDelegates = @()
                }

                # Get Send As delegates directly
                try {
                    if ($SendAsMethod -eq "MailboxPermission") {
                        $sendAsDelegates = Get-MailboxPermission $mailbox.PrimarySmtpAddress |
                            Where-Object { $_.AccessRights -eq 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
                    }
                    else {
                        $sendAsDelegates = Get-RecipientPermission $mailbox.PrimarySmtpAddress |
                            Where-Object { $_.AccessRights -eq 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false }
                    }
                }
                catch {
                    Write-Warning "Could not retrieve Send As permissions for $($mailbox.DisplayName): $($_.Exception.Message)"
                    $sendAsDelegates = @()
                }
            }

            if ($mailbox.GrantSendOnBehalfTo)
            {
                $sendOnBehalfDelegates = $mailbox.GrantSendOnBehalfTo
            }

            # Get inbox rules for forwarding
            try
            {
                $inboxRules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
                $forwardingRules = $inboxRules | Where-Object {
                    $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo
                }
            }
            catch
            {
                Write-Warning "Could not retrieve inbox rules for $($mailbox.DisplayName): $($_.Exception.Message)"
            }

            # Create result objects based on mode
            if ($ReverseLookup)
            {
                # In reverse lookup mode, find if any of our target users have permissions on this mailbox

                # Check Full Access delegates
                foreach ($delegate in $fullAccessDelegates)
                {
                    $delegateUser = $delegate.User
                    Write-Verbose "Checking Full Access delegate: $delegateUser"
                    $containsMatch = $targetUserUPNs -contains $delegateUser
                    $wildcardMatch = $targetUserUPNs | Where-Object { $_ -like "*$delegateUser*" }
                    Write-Verbose "Contains match: $containsMatch, Wildcard match: $($null -ne $wildcardMatch)"
                    if ($containsMatch -or $wildcardMatch)
                    {
                        $results += [PSCustomObject]@{
                            UserUPN = $delegateUser
                            UserDisplayName = $delegateUser
                            MailboxName = $mailbox.DisplayName
                            MailboxUPN = $mailbox.UserPrincipalName
                            MailboxType = $mailbox.RecipientTypeDetails
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            PermissionType = "FullAccess"
                            AccessRights = ($delegate.AccessRights -join ", ")
                            ForwardingAddress = $mailbox.ForwardingAddress
                            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        }
                    }
                }

                # Check Send As delegates
                foreach ($delegate in $sendAsDelegates)
                {
                    $delegateName = $delegate.Trustee
                    Write-Verbose "Checking Send As delegate: $delegateName"
                    if ($targetUserUPNs -contains $delegateName -or $targetUserUPNs | Where-Object { $_ -like "*$delegateName*" })
                    {
                        $results += [PSCustomObject]@{
                            UserUPN = $delegateName
                            UserDisplayName = $delegateName
                            MailboxName = $mailbox.DisplayName
                            MailboxUPN = $mailbox.UserPrincipalName
                            MailboxType = $mailbox.RecipientTypeDetails
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            PermissionType = "SendAs"
                            AccessRights = ($delegate.AccessRights -join ", ")
                            ForwardingAddress = $mailbox.ForwardingAddress
                            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        }
                    }
                }

                # Check Send On Behalf delegates
                foreach ($delegate in $sendOnBehalfDelegates)
                {
                    if ($targetUserUPNs -contains $delegate -or $targetUserUPNs | Where-Object { $_ -like "*$delegate*" })
                    {
                        $results += [PSCustomObject]@{
                            UserUPN = $delegate
                            UserDisplayName = $delegate
                            MailboxName = $mailbox.DisplayName
                            MailboxUPN = $mailbox.UserPrincipalName
                            MailboxType = $mailbox.RecipientTypeDetails
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            PermissionType = "SendOnBehalf"
                            AccessRights = "SendOnBehalf"
                            ForwardingAddress = $mailbox.ForwardingAddress
                            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        }
                    }
                }

                # Check forwarding rules for target users
                foreach ($rule in $forwardingRules)
                {
                    $forwardingDests = @()
                    if ($rule.ForwardTo) { $forwardingDests += $rule.ForwardTo }
                    if ($rule.ForwardAsAttachmentTo) { $forwardingDests += $rule.ForwardAsAttachmentTo }
                    if ($rule.RedirectTo) { $forwardingDests += $rule.RedirectTo }

                    foreach ($dest in $forwardingDests)
                    {
                        if ($targetUserUPNs -contains $dest -or $targetUserUPNs | Where-Object { $_ -like "*$dest*" })
                        {
                            $results += [PSCustomObject]@{
                                UserUPN = $dest
                                UserDisplayName = $dest
                                MailboxName = $mailbox.DisplayName
                                MailboxUPN = $mailbox.UserPrincipalName
                                MailboxType = $mailbox.RecipientTypeDetails
                                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                                PermissionType = "ForwardingRule"
                                AccessRights = "Receives forwards via rule: $($rule.Name)"
                                ForwardingAddress = $mailbox.ForwardingAddress
                                ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                                DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                            }
                        }
                    }
                }
            }
            else
            {
                # Regular mode - existing logic
                if ($fullAccessDelegates -or $sendAsDelegates -or $sendOnBehalfDelegates -or $mailbox.ForwardingAddress -or $mailbox.ForwardingSmtpAddress -or $forwardingRules)
                {
                    # Process Full Access delegates
                    foreach ($delegate in $fullAccessDelegates)
                    {
                        $results += [PSCustomObject]@{
                            MailboxName = $mailbox.DisplayName
                            MailboxUPN = $mailbox.UserPrincipalName
                            MailboxType = $mailbox.RecipientTypeDetails
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            Office = if ($adUser) { $adUser.Office } else { $mailbox.Office }
                            Department = if ($adUser) { $adUser.Department } else { "" }
                            Manager = $manager
                            PermissionType = "FullAccess"
                            Delegate = $delegate.User
                            DelegateAccessRights = ($delegate.AccessRights -join ", ")
                            ForwardingType = ""
                            ForwardingDestination = ""
                            ForwardingAddress = $mailbox.ForwardingAddress
                            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                            RuleName = ""
                            RuleCondition = ""
                        }
                    }

                # Process Send As delegates
                foreach ($delegate in $sendAsDelegates)
                {
                    # Both Get-MailboxPermission and Get-RecipientPermission use Trustee property for Send As permissions
                    $delegateName = $delegate.Trustee

                    $results += [PSCustomObject]@{
                        MailboxName = $mailbox.DisplayName
                        MailboxUPN = $mailbox.UserPrincipalName
                        MailboxType = $mailbox.RecipientTypeDetails
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        Office = if ($adUser) { $adUser.Office } else { $mailbox.Office }
                        Department = if ($adUser) { $adUser.Department } else { "" }
                        Manager = $manager
                        PermissionType = "SendAs"
                        Delegate = $delegateName
                        DelegateAccessRights = ($delegate.AccessRights -join ", ")
                        ForwardingType = ""
                        ForwardingDestination = ""
                        ForwardingAddress = $mailbox.ForwardingAddress
                        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        RuleName = ""
                        RuleCondition = ""
                    }
                }

                # Process Send On Behalf delegates
                foreach ($delegate in $sendOnBehalfDelegates)
                {
                    $results += [PSCustomObject]@{
                        MailboxName = $mailbox.DisplayName
                        MailboxUPN = $mailbox.UserPrincipalName
                        MailboxType = $mailbox.RecipientTypeDetails
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        Office = if ($adUser) { $adUser.Office } else { $mailbox.Office }
                        Department = if ($adUser) { $adUser.Department } else { "" }
                        Manager = $manager
                        PermissionType = "SendOnBehalf"
                        Delegate = $delegate
                        DelegateAccessRights = "SendOnBehalf"
                        ForwardingType = ""
                        ForwardingDestination = ""
                        ForwardingAddress = $mailbox.ForwardingAddress
                        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        RuleName = ""
                        RuleCondition = ""
                    }
                }

                # Process forwarding rules
                foreach ($rule in $forwardingRules)
                {
                    $forwardingDest = ""
                    $forwardingType = ""

                    if ($rule.ForwardTo)
                    {
                        $forwardingType = "ForwardTo"
                        $forwardingDest = ($rule.ForwardTo -join ", ")
                    }
                    elseif ($rule.ForwardAsAttachmentTo)
                    {
                        $forwardingType = "ForwardAsAttachment"
                        $forwardingDest = ($rule.ForwardAsAttachmentTo -join ", ")
                    }
                    elseif ($rule.RedirectTo)
                    {
                        $forwardingType = "RedirectTo"
                        $forwardingDest = ($rule.RedirectTo -join ", ")
                    }

                    $results += [PSCustomObject]@{
                        MailboxName = $mailbox.DisplayName
                        MailboxUPN = $mailbox.UserPrincipalName
                        MailboxType = $mailbox.RecipientTypeDetails
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        Office = if ($adUser) { $adUser.Office } else { $mailbox.Office }
                        Department = if ($adUser) { $adUser.Department } else { "" }
                        Manager = $manager
                        PermissionType = "InboxRule"
                        Delegate = ""
                        DelegateAccessRights = ""
                        ForwardingType = $forwardingType
                        ForwardingDestination = $forwardingDest
                        ForwardingAddress = $mailbox.ForwardingAddress
                        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        RuleName = $rule.Name
                        RuleCondition = $rule.Description
                    }
                }

                # If mailbox has native forwarding but no other delegations, add entry
                if (($mailbox.ForwardingAddress -or $mailbox.ForwardingSmtpAddress) -and
                    -not $fullAccessDelegates -and -not $sendAsDelegates -and -not $sendOnBehalfDelegates -and -not $forwardingRules)
                {
                    $results += [PSCustomObject]@{
                        MailboxName = $mailbox.DisplayName
                        MailboxUPN = $mailbox.UserPrincipalName
                        MailboxType = $mailbox.RecipientTypeDetails
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        Office = if ($adUser) { $adUser.Office } else { $mailbox.Office }
                        Department = if ($adUser) { $adUser.Department } else { "" }
                        Manager = $manager
                        PermissionType = "MailboxForwarding"
                        Delegate = ""
                        DelegateAccessRights = ""
                        ForwardingType = "MailboxForwarding"
                        ForwardingDestination = if ($mailbox.ForwardingSmtpAddress) { $mailbox.ForwardingSmtpAddress } else { $mailbox.ForwardingAddress }
                        ForwardingAddress = $mailbox.ForwardingAddress
                        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                        RuleName = ""
                        RuleCondition = ""
                    }
                }
            }
        }
        }
        catch
        {
            Write-Warning "Error processing mailbox $($mailbox.DisplayName): $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "Processing Mailboxes" -Completed

    # Process distribution groups in reverse lookup mode
    if ($ReverseLookup -and $targetDistributionGroups)
    {
        $dgProcessText = if ($targetDistributionGroups.Count -eq 1) { "distribution group" } else { "distribution groups" }
        Write-Output "Processing $($targetDistributionGroups.Count) $dgProcessText for reverse lookup..."

        foreach ($distributionGroup in $targetDistributionGroups)
        {
            Write-Progress -Activity "Processing Distribution Groups" -Status "Processing $($distributionGroup.DisplayName)" -PercentComplete (($targetDistributionGroups.IndexOf($distributionGroup) / $targetDistributionGroups.Count) * 100)

            try
            {
                # Check AcceptMessagesOnlyFrom
                if ($distributionGroup.AcceptMessagesOnlyFrom)
                {
                    Write-Verbose "Checking AcceptMessagesOnlyFrom for $($distributionGroup.DisplayName): $($distributionGroup.AcceptMessagesOnlyFrom -join ', ')"
                    foreach ($senderIdentity in $distributionGroup.AcceptMessagesOnlyFrom)
                    {
                        $resolvedSender = Resolve-UserIdentity -Identity $senderIdentity -IdentityCache $script:IdentityCache
                        Write-Verbose "Resolved $senderIdentity to $resolvedSender"
                        Write-Verbose "Checking if $resolvedSender matches any of: $($targetUserUPNs -join ', ')"
                        if ($targetUserUPNs -contains $resolvedSender)
                        {
                            Write-Output "MATCH FOUND: User $resolvedSender has AcceptMessagesOnlyFrom permission on distribution group $($distributionGroup.DisplayName)"
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedSender
                                UserDisplayName = $resolvedSender
                                MailboxName = $distributionGroup.DisplayName
                                MailboxUPN = $distributionGroup.PrimarySmtpAddress
                                MailboxType = "DistributionGroup"
                                PrimarySmtpAddress = $distributionGroup.PrimarySmtpAddress
                                PermissionType = "AcceptMessagesOnlyFrom"
                                AccessRights = "Can send messages to this distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check AcceptMessagesOnlyFromSendersOrMembers
                if ($distributionGroup.AcceptMessagesOnlyFromSendersOrMembers)
                {
                    foreach ($senderIdentity in $distributionGroup.AcceptMessagesOnlyFromSendersOrMembers)
                    {
                        $resolvedSender = Resolve-UserIdentity -Identity $senderIdentity -IdentityCache $script:IdentityCache
                        if ($targetUserUPNs -contains $resolvedSender)
                        {
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedSender
                                UserDisplayName = $resolvedSender
                                MailboxName = $distributionGroup.DisplayName
                                MailboxUPN = $distributionGroup.PrimarySmtpAddress
                                MailboxType = "DistributionGroup"
                                PrimarySmtpAddress = $distributionGroup.PrimarySmtpAddress
                                PermissionType = "AcceptMessagesOnlyFromSendersOrMembers"
                                AccessRights = "Can send messages to this distribution group (sender or member)"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check GrantSendOnBehalfTo
                if ($distributionGroup.GrantSendOnBehalfTo)
                {
                    foreach ($delegate in $distributionGroup.GrantSendOnBehalfTo)
                    {
                        $resolvedDelegate = Resolve-UserIdentity -Identity $delegate -IdentityCache $script:IdentityCache
                        if ($targetUserUPNs -contains $resolvedDelegate)
                        {
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedDelegate
                                UserDisplayName = $resolvedDelegate
                                MailboxName = $distributionGroup.DisplayName
                                MailboxUPN = $distributionGroup.PrimarySmtpAddress
                                MailboxType = "DistributionGroup"
                                PrimarySmtpAddress = $distributionGroup.PrimarySmtpAddress
                                PermissionType = "GrantSendOnBehalfTo"
                                AccessRights = "Can send on behalf of this distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check ModeratedBy
                if ($distributionGroup.ModeratedBy)
                {
                    foreach ($moderator in $distributionGroup.ModeratedBy)
                    {
                        $resolvedModerator = Resolve-UserIdentity -Identity $moderator -IdentityCache $script:IdentityCache
                        if ($targetUserUPNs -contains $resolvedModerator)
                        {
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedModerator
                                UserDisplayName = $resolvedModerator
                                MailboxName = $distributionGroup.DisplayName
                                MailboxUPN = $distributionGroup.PrimarySmtpAddress
                                MailboxType = "DistributionGroup"
                                PrimarySmtpAddress = $distributionGroup.PrimarySmtpAddress
                                PermissionType = "ModeratedBy"
                                AccessRights = "Moderator for this distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }
            }
            catch
            {
                Write-Warning "Error processing distribution group $($distributionGroup.DisplayName): $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Processing Distribution Groups" -Completed
    }

    # Process dynamic distribution groups in reverse lookup mode
    if ($ReverseLookup -and $targetDynamicDistributionGroups)
    {
        $ddgProcessText = if ($targetDynamicDistributionGroups.Count -eq 1) { "dynamic distribution group" } else { "dynamic distribution groups" }
        Write-Output "Processing $($targetDynamicDistributionGroups.Count) $ddgProcessText for reverse lookup..."

        foreach ($dynamicGroup in $targetDynamicDistributionGroups)
        {
            Write-Progress -Activity "Processing Dynamic Distribution Groups" -Status "Processing $($dynamicGroup.DisplayName)" -PercentComplete (($targetDynamicDistributionGroups.IndexOf($dynamicGroup) / $targetDynamicDistributionGroups.Count) * 100)

            try
            {
                # Check AcceptMessagesOnlyFrom
                if ($dynamicGroup.AcceptMessagesOnlyFrom)
                {
                    Write-Verbose "Checking AcceptMessagesOnlyFrom for $($dynamicGroup.DisplayName): $($dynamicGroup.AcceptMessagesOnlyFrom -join ', ')"
                    foreach ($senderIdentity in $dynamicGroup.AcceptMessagesOnlyFrom)
                    {
                        $resolvedSender = Resolve-UserIdentity -Identity $senderIdentity -IdentityCache $script:IdentityCache
                        Write-Verbose "Resolved $senderIdentity to $resolvedSender"
                        Write-Verbose "Checking if $resolvedSender matches any of: $($targetUserUPNs -join ', ')"
                        if ($targetUserUPNs -contains $resolvedSender)
                        {
                            Write-Output "MATCH FOUND: User $resolvedSender has AcceptMessagesOnlyFrom permission on dynamic distribution group $($dynamicGroup.DisplayName)"
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedSender
                                UserDisplayName = $resolvedSender
                                MailboxName = $dynamicGroup.DisplayName
                                MailboxUPN = $dynamicGroup.PrimarySmtpAddress
                                MailboxType = "DynamicDistributionGroup"
                                PrimarySmtpAddress = $dynamicGroup.PrimarySmtpAddress
                                PermissionType = "AcceptMessagesOnlyFrom"
                                AccessRights = "Can send messages to this dynamic distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check AcceptMessagesOnlyFromSendersOrMembers
                if ($dynamicGroup.AcceptMessagesOnlyFromSendersOrMembers)
                {
                    Write-Verbose "Checking AcceptMessagesOnlyFromSendersOrMembers for $($dynamicGroup.DisplayName): $($dynamicGroup.AcceptMessagesOnlyFromSendersOrMembers -join ', ')"
                    foreach ($senderIdentity in $dynamicGroup.AcceptMessagesOnlyFromSendersOrMembers)
                    {
                        $resolvedSender = Resolve-UserIdentity -Identity $senderIdentity -IdentityCache $script:IdentityCache
                        Write-Verbose "Resolved $senderIdentity to $resolvedSender"
                        Write-Verbose "Checking if $resolvedSender matches any of: $($targetUserUPNs -join ', ')"
                        if ($targetUserUPNs -contains $resolvedSender)
                        {
                            Write-Output "MATCH FOUND: User $resolvedSender has AcceptMessagesOnlyFromSendersOrMembers permission on dynamic distribution group $($dynamicGroup.DisplayName)"
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedSender
                                UserDisplayName = $resolvedSender
                                MailboxName = $dynamicGroup.DisplayName
                                MailboxUPN = $dynamicGroup.PrimarySmtpAddress
                                MailboxType = "DynamicDistributionGroup"
                                PrimarySmtpAddress = $dynamicGroup.PrimarySmtpAddress
                                PermissionType = "AcceptMessagesOnlyFromSendersOrMembers"
                                AccessRights = "Can send messages to this dynamic distribution group (sender or member)"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check GrantSendOnBehalfTo
                if ($dynamicGroup.GrantSendOnBehalfTo)
                {
                    Write-Verbose "Checking GrantSendOnBehalfTo for $($dynamicGroup.DisplayName): $($dynamicGroup.GrantSendOnBehalfTo -join ', ')"
                    foreach ($delegate in $dynamicGroup.GrantSendOnBehalfTo)
                    {
                        $resolvedDelegate = Resolve-UserIdentity -Identity $delegate -IdentityCache $script:IdentityCache
                        Write-Verbose "Resolved $delegate to $resolvedDelegate"
                        Write-Verbose "Checking if $resolvedDelegate matches any of: $($targetUserUPNs -join ', ')"
                        if ($targetUserUPNs -contains $resolvedDelegate)
                        {
                            Write-Output "MATCH FOUND: User $resolvedDelegate has GrantSendOnBehalfTo permission on dynamic distribution group $($dynamicGroup.DisplayName)"
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedDelegate
                                UserDisplayName = $resolvedDelegate
                                MailboxName = $dynamicGroup.DisplayName
                                MailboxUPN = $dynamicGroup.PrimarySmtpAddress
                                MailboxType = "DynamicDistributionGroup"
                                PrimarySmtpAddress = $dynamicGroup.PrimarySmtpAddress
                                PermissionType = "GrantSendOnBehalfTo"
                                AccessRights = "Can send on behalf of this dynamic distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }

                # Check ModeratedBy
                if ($dynamicGroup.ModeratedBy)
                {
                    Write-Verbose "Checking ModeratedBy for $($dynamicGroup.DisplayName): $($dynamicGroup.ModeratedBy -join ', ')"
                    foreach ($moderator in $dynamicGroup.ModeratedBy)
                    {
                        $resolvedModerator = Resolve-UserIdentity -Identity $moderator -IdentityCache $script:IdentityCache
                        Write-Verbose "Resolved $moderator to $resolvedModerator"
                        Write-Verbose "Checking if $resolvedModerator matches any of: $($targetUserUPNs -join ', ')"
                        if ($targetUserUPNs -contains $resolvedModerator)
                        {
                            Write-Output "MATCH FOUND: User $resolvedModerator has ModeratedBy permission on dynamic distribution group $($dynamicGroup.DisplayName)"
                            $results += [PSCustomObject]@{
                                UserUPN = $resolvedModerator
                                UserDisplayName = $resolvedModerator
                                MailboxName = $dynamicGroup.DisplayName
                                MailboxUPN = $dynamicGroup.PrimarySmtpAddress
                                MailboxType = "DynamicDistributionGroup"
                                PrimarySmtpAddress = $dynamicGroup.PrimarySmtpAddress
                                PermissionType = "ModeratedBy"
                                AccessRights = "Moderator for this dynamic distribution group"
                                ForwardingAddress = ""
                                ForwardingSmtpAddress = ""
                                DeliverToMailboxAndForward = ""
                            }
                        }
                    }
                }
            }
            catch
            {
                Write-Warning "Error processing dynamic distribution group $($dynamicGroup.DisplayName): $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Processing Dynamic Distribution Groups" -Completed
    }

    # Export results to CSV
    if ($results.Count -gt 0)
    {
        Write-Output "Exporting $($results.Count) records to CSV..."
        if ($ReverseLookup)
        {
            $results | Sort-Object UserUPN, MailboxName, PermissionType | Export-Csv $OutputFile -NoTypeInformation
        }
        else
        {
            $results | Sort-Object MailboxName, PermissionType, Delegate | Export-Csv $OutputFile -NoTypeInformation
        }
        Write-Output "CSV report saved to: $OutputFile"

        # Display summary statistics
        Write-Output "Report Summary:"
        Write-Output "  Total Records: $($results.Count)"
        $mbSummaryText = if ($targetMailboxes.Count -eq 1) { "Mailbox" } else { "Mailboxes" }
        Write-Output "  $mbSummaryText Processed: $($targetMailboxes.Count)"
        if ($ReverseLookup -and $targetDistributionGroups) {
            $dgSummaryText = if ($targetDistributionGroups.Count -eq 1) { "Distribution Group" } else { "Distribution Groups" }
            Write-Output "  $dgSummaryText Processed: $($targetDistributionGroups.Count)"
        }
        if ($ReverseLookup -and $targetDynamicDistributionGroups) {
            $ddgSummaryText = if ($targetDynamicDistributionGroups.Count -eq 1) { "Dynamic Distribution Group" } else { "Dynamic Distribution Groups" }
            Write-Output "  $ddgSummaryText Processed: $($targetDynamicDistributionGroups.Count)"
        }
        if ($MailboxTypes.Count -gt 0) { Write-Output "  Mailbox Types: $($MailboxTypes -join ', ')" }
        if ($InputCsvPath) { Write-Output "  CSV Filter Applied: $InputCsvPath" }
        Write-Output "  Output Location: $OutputFile"

        Write-Output "Permission Type Summary:"
        $results | Group-Object PermissionType | Select-Object Name, Count | Sort-Object Name | ForEach-Object {
            Write-Output "  $($_.Name): $($_.Count)"
        }
    }
    else
    {
        Write-Output "No delegation permissions or forwarding configurations found for the specified criteria."
    }

}
catch
{
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
}
finally
{
    # Leave Exchange Online session active for user
    Write-Output "Script completed. Exchange Online session maintained."
}

Write-Output "Mailbox Delegation Report Export completed."