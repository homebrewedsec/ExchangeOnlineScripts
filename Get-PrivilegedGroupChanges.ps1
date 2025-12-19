<#
.SYNOPSIS
Monitors privileged Active Directory groups for membership changes, password changes, and login activity.

.DESCRIPTION
This script queries Active Directory for privileged groups in specified OUs, checks for recent
changes to group membership, monitors password changes for members of those groups, and optionally
queries domain controller event logs for login activity by privileged users.

The script generates CSV reports and optionally sends HTML email notifications when activity is detected.
By default, empty reports (no changes) are suppressed unless -AlwaysReport is specified.

.PARAMETER SearchBaseOUs
Required. Array of distinguished names for OUs/containers to search for privileged groups.
Example: "OU=Admins,DC=corp,DC=contoso,DC=com","CN=Builtin,DC=corp,DC=contoso,DC=com"

.PARAMETER DomainControllerSearchBase
Required. Distinguished name of the OU containing domain controllers.
Example: "OU=Domain Controllers,DC=corp,DC=contoso,DC=com"

.PARAMETER LookbackMinutes
Optional. How far back to look for changes in minutes. Default is 60 (1 hour).

.PARAMETER EventIDs
Optional. Array of security event IDs to monitor. Defaults to common authentication events:
4624 (Logon), 4625 (Failed logon), 4648 (Explicit credentials), 4672 (Special privileges),
4768-4771 (Kerberos), 4774, 4776 (Credential validation).

.PARAMETER SmtpServer
Optional. SMTP server for sending email notifications. If not specified, only CSV output is generated.

.PARAMETER SmtpTo
Optional. Array of email recipients. Required if SmtpServer is specified.

.PARAMETER SmtpFrom
Optional. Email sender address. Required if SmtpServer is specified.

.PARAMETER SmtpSubject
Optional. Email subject line. Default is "Privileged Group Activity".

.PARAMETER OutputPath
Optional. Directory for CSV output files. Default is the current directory.

.PARAMETER SkipEventLogs
Optional switch. Skip event log queries on domain controllers for faster execution.

.PARAMETER AlwaysReport
Optional switch. Generate output even if no changes are detected.

.EXAMPLE
.\Get-PrivilegedGroupChanges.ps1 -SearchBaseOUs "OU=Admins,DC=corp,DC=contoso,DC=com","CN=Builtin,DC=corp,DC=contoso,DC=com" -DomainControllerSearchBase "OU=Domain Controllers,DC=corp,DC=contoso,DC=com"

Basic usage - generates CSV report of changes in the last 60 minutes.

.EXAMPLE
.\Get-PrivilegedGroupChanges.ps1 -SearchBaseOUs "OU=Admins,DC=corp,DC=contoso,DC=com" -DomainControllerSearchBase "OU=Domain Controllers,DC=corp,DC=contoso,DC=com" -LookbackMinutes 1440 -SmtpServer "mail.contoso.com" -SmtpTo "security@contoso.com" -SmtpFrom "monitoring@contoso.com"

Check the last 24 hours and send email notification.

.EXAMPLE
.\Get-PrivilegedGroupChanges.ps1 -SearchBaseOUs "OU=Admins,DC=corp,DC=contoso,DC=com" -DomainControllerSearchBase "OU=Domain Controllers,DC=corp,DC=contoso,DC=com" -SkipEventLogs

Skip event log queries for faster execution (only reports group and password changes).

.NOTES
Author: Hudson Bush, Seguri - hudson@seguri.io
Requires: ActiveDirectory PowerShell module
Prerequisites: Read access to Active Directory, Event log read access on domain controllers (unless -SkipEventLogs)
Output: CSV file in OutputPath, optional HTML email
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$SearchBaseOUs,

    [Parameter(Mandatory)]
    [string]$DomainControllerSearchBase,

    [int]$LookbackMinutes = 60,

    [string[]]$EventIDs = @("4624", "4625", "4648", "4634", "4672", "4647", "4778", "4768", "4769", "4770", "4771", "4774", "4776"),

    [string]$SmtpServer,

    [string[]]$SmtpTo,

    [string]$SmtpFrom,

    [string]$SmtpSubject = "Privileged Group Activity",

    [string]$OutputPath = (Get-Location).Path,

    [switch]$SkipEventLogs,

    [switch]$AlwaysReport
)

# Validate SMTP parameters if SmtpServer is specified
if ($SmtpServer)
{
    if (-not $SmtpTo -or -not $SmtpFrom)
    {
        Write-Error "When SmtpServer is specified, SmtpTo and SmtpFrom are required."
        exit 1
    }
}

# Validate AD module is available
Write-Verbose "Checking for ActiveDirectory module..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory))
{
    Write-Error "ActiveDirectory PowerShell module is not installed. Please install RSAT or run from a domain controller."
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

try
{
    # Phase 1: Calculate lookback time and collect groups
    Write-Output "Phase 1: Collecting privileged groups from specified OUs..."
    $LookbackTime = (Get-Date).AddMinutes(-$LookbackMinutes)
    Write-Verbose "Looking back to: $LookbackTime"

    $Groups = @()
    foreach ($OU in $SearchBaseOUs)
    {
        Write-Verbose "Searching OU: $OU"
        try
        {
            $Groups += Get-ADGroup -Filter * -SearchBase $OU -Properties WhenChanged, Members
        }
        catch
        {
            Write-Warning "Could not query OU '$OU': $($_.Exception.Message)"
        }
    }

    if ($Groups.Count -eq 0)
    {
        Write-Warning "No groups found in specified OUs."
        if (-not $AlwaysReport)
        {
            Write-Output "No groups to monitor. Exiting."
            exit 0
        }
    }

    Write-Output "Found $($Groups.Count) groups to monitor."

    # Phase 2: Check for group membership changes
    Write-Output "Phase 2: Checking for group membership changes..."
    Write-Output "  Lookback time: $LookbackTime"
    $GroupChanges = @()
    $PrivilegedUsers = @()

    foreach ($Group in $Groups)
    {
        Write-Verbose "  Checking $($Group.Name): WhenChanged=$($Group.WhenChanged)"
        if ($Group.WhenChanged -gt $LookbackTime)
        {
            Write-Output "  [CHANGE DETECTED] $($Group.Name) - modified $($Group.WhenChanged)"
            $GroupChanges += [PSCustomObject]@{
                ActivityType    = "GroupChange"
                Name            = $Group.Name
                Timestamp       = $Group.WhenChanged
                Details         = "Group modified"
                MemberCount     = ($Group.Members | Measure-Object).Count
            }
        }

        # Collect unique privileged users from all groups
        if ($Group.Members)
        {
            foreach ($MemberDN in $Group.Members)
            {
                try
                {
                    $MemberObj = Get-ADObject $MemberDN -ErrorAction SilentlyContinue
                    if ($MemberObj -and $MemberObj.ObjectClass -eq "user")
                    {
                        $PrivilegedUsers += $MemberDN
                    }
                }
                catch
                {
                    Write-Verbose "Could not resolve member: $MemberDN"
                }
            }
        }
    }

    # Get unique privileged users
    $PrivilegedUsers = $PrivilegedUsers | Select-Object -Unique
    Write-Output "Found $($PrivilegedUsers.Count) unique privileged users."

    # Phase 3: Get full user objects and check password changes
    Write-Output "Phase 3: Checking for password changes..."
    $PasswordChanges = @()
    $UserObjects = @()
    $EnabledUserObjects = @()

    foreach ($UserDN in $PrivilegedUsers)
    {
        try
        {
            $User = Get-ADUser $UserDN -Properties PasswordLastSet, DisplayName, SamAccountName, Enabled
            $UserObjects += $User

            if ($User.Enabled)
            {
                $EnabledUserObjects += $User
            }

            if ($User.PasswordLastSet -gt $LookbackTime)
            {
                $PasswordChanges += [PSCustomObject]@{
                    ActivityType    = "PasswordChange"
                    Name            = $User.DisplayName
                    Timestamp       = $User.PasswordLastSet
                    Details         = "Password changed for $($User.SamAccountName) (Enabled: $($User.Enabled))"
                    MemberCount     = $null
                }
            }
        }
        catch
        {
            Write-Verbose "Could not get user details for: $UserDN"
        }
    }

    $disabledCount = $UserObjects.Count - $EnabledUserObjects.Count
    Write-Output "Found $($PasswordChanges.Count) password changes."
    Write-Output "Found $($EnabledUserObjects.Count) enabled users, $disabledCount disabled (will skip disabled for event logs)."

    # Phase 4: Query event logs on domain controllers (unless skipped)
    $PrivilegedLogins = @()

    if (-not $SkipEventLogs)
    {
        Write-Output "Phase 4: Querying event logs on domain controllers..."

        try
        {
            $DomainControllers = Get-ADComputer -SearchBase $DomainControllerSearchBase -Filter *
        }
        catch
        {
            Write-Warning "Could not query domain controllers: $($_.Exception.Message)"
            $DomainControllers = @()
        }

        $dcCount = $DomainControllers.Count
        $dcCurrent = 0

        # Build a hashtable for quick user lookup by SamAccountName (enabled users only)
        $UserLookup = @{}
        foreach ($User in $EnabledUserObjects)
        {
            $UserLookup[$User.SamAccountName.ToLower()] = $User
        }

        foreach ($DC in $DomainControllers)
        {
            $dcCurrent++
            Write-Progress -Activity "Querying Domain Controllers" -Status "Checking $($DC.Name) ($dcCurrent of $dcCount)" -PercentComplete (($dcCurrent / $dcCount) * 100)

            try
            {
                # Single query per DC - much faster than one query per user
                $Events = Get-WinEvent -ComputerName $DC.Name -FilterHashtable @{
                    LogName   = 'Security'
                    Id        = $EventIDs
                    StartTime = $LookbackTime
                } -ErrorAction SilentlyContinue

                if ($Events)
                {
                    Write-Verbose "  Found $($Events.Count) events on $($DC.Name), filtering for privileged users..."

                    foreach ($LogEntry in $Events)
                    {
                        # Check if any of our privileged users are in this event
                        foreach ($SamAccountName in $UserLookup.Keys)
                        {
                            if ($LogEntry.Message -like "*$SamAccountName*")
                            {
                                $User = $UserLookup[$SamAccountName]
                                $PrivilegedLogins += [PSCustomObject]@{
                                    ActivityType    = "LoginEvent"
                                    Name            = $User.DisplayName
                                    Timestamp       = $LogEntry.TimeCreated
                                    Details         = "Event $($LogEntry.Id) on $($DC.Name)"
                                    MemberCount     = $null
                                }
                                break  # Found a match, no need to check other users for this event
                            }
                        }
                    }
                }
            }
            catch
            {
                Write-Verbose "Could not query events on $($DC.Name): $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Querying Domain Controllers" -Completed
        Write-Output "Found $($PrivilegedLogins.Count) login events."
    }
    else
    {
        Write-Output "Phase 4: Skipping event log queries (-SkipEventLogs specified)."
    }

    # Phase 5: Generate output
    Write-Output "Phase 5: Generating output..."

    # Combine all results
    $AllResults = @()
    $AllResults += $GroupChanges
    $AllResults += $PasswordChanges
    $AllResults += $PrivilegedLogins

    # Check if there's anything to report
    if ($AllResults.Count -eq 0 -and -not $AlwaysReport)
    {
        Write-Output "No changes detected in the last $LookbackMinutes minutes. No report generated."
        Write-Output "Use -AlwaysReport to generate output even when no changes are found."
        exit 0
    }

    # Generate CSV output
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFile = Join-Path $OutputPath "PrivilegedGroupActivity_$Timestamp.csv"

    if ($AllResults.Count -gt 0)
    {
        $AllResults | Sort-Object Timestamp -Descending | Export-Csv $OutputFile -NoTypeInformation
        Write-Output "CSV report saved to: $OutputFile"
    }
    elseif ($AlwaysReport)
    {
        # Create empty report with headers
        [PSCustomObject]@{
            ActivityType = "NoActivity"
            Name         = "No changes detected"
            Timestamp    = Get-Date
            Details      = "No privileged group activity in the last $LookbackMinutes minutes"
            MemberCount  = $null
        } | Export-Csv $OutputFile -NoTypeInformation
        Write-Output "Empty report saved to: $OutputFile"
    }

    # Send email if SMTP is configured
    if ($SmtpServer)
    {
        Write-Output "Sending email notification..."

        $Body = "<html><head><style>"
        $Body += "body { font-family: Arial, sans-serif; }"
        $Body += "h2 { color: #333; }"
        $Body += "table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }"
        $Body += "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }"
        $Body += "th { background-color: #4472C4; color: white; }"
        $Body += "tr:nth-child(even) { background-color: #f2f2f2; }"
        $Body += "</style></head><body>"
        $Body += "<h1>Privileged Group Activity Report</h1>"
        $Body += "<p>Lookback period: $LookbackMinutes minutes (since $LookbackTime)</p>"

        if ($GroupChanges.Count -gt 0)
        {
            $Body += "<h2>Group Changes ($($GroupChanges.Count))</h2>"
            $Body += $GroupChanges | Select-Object Name, Timestamp, MemberCount | ConvertTo-Html -Fragment
        }

        if ($PasswordChanges.Count -gt 0)
        {
            $Body += "<h2>Password Changes ($($PasswordChanges.Count))</h2>"
            $Body += $PasswordChanges | Select-Object Name, Timestamp, Details | ConvertTo-Html -Fragment
        }

        if ($PrivilegedLogins.Count -gt 0)
        {
            $Body += "<h2>Login Events ($($PrivilegedLogins.Count))</h2>"
            $Body += $PrivilegedLogins | Select-Object Name, Timestamp, Details -Unique | ConvertTo-Html -Fragment
        }

        if ($AllResults.Count -eq 0)
        {
            $Body += "<p>No privileged group activity detected in the specified time period.</p>"
        }

        $Body += "<hr><p><em>Generated: $(Get-Date)</em></p>"
        $Body += "</body></html>"

        try
        {
            Send-MailMessage -SmtpServer $SmtpServer -From $SmtpFrom -To $SmtpTo -Subject $SmtpSubject -Body $Body -BodyAsHtml -Priority High
            Write-Output "Email sent successfully to: $($SmtpTo -join ', ')"
        }
        catch
        {
            Write-Warning "Failed to send email: $($_.Exception.Message)"
        }
    }

    # Summary
    Write-Output ""
    Write-Output "=== Summary ==="
    Write-Output "Groups monitored: $($Groups.Count)"
    Write-Output "Privileged users: $($PrivilegedUsers.Count)"
    Write-Output "Group changes: $($GroupChanges.Count)"
    Write-Output "Password changes: $($PasswordChanges.Count)"
    Write-Output "Login events: $($PrivilegedLogins.Count)"
    Write-Output "Total activities: $($AllResults.Count)"
}
catch
{
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
    exit 1
}
finally
{
    Write-Output "Script completed at $(Get-Date)"
}
