#Requires -Version 7.0
<#
.SYNOPSIS
    Helper functions for the 2X user offboarding automation.
#>

Set-StrictMode -Version Latest

# ── Logging ───────────────────────────────────────────────────────────────────

$script:LogFile = $null

function Initialize-OffboardingLog {
    [CmdletBinding()]
    param(
        [string]$LogDirectory,
        [string]$UserPrincipalName
    )
    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $timestamp       = Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeUpn         = $UserPrincipalName -replace '[^a-zA-Z0-9_@.-]', '_'
    $script:LogFile  = Join-Path $LogDirectory "Offboarding_${safeUpn}_${timestamp}.log"
    Write-OffboardingLog -Message "Offboarding log started for $UserPrincipalName" -Level INFO
}

function Write-OffboardingLog {
    [CmdletBinding()]
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'STEP')]
        [string]$Level = 'INFO',
        [switch]$NoConsole
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp] [$Level] $Message"

    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8
    }

    if (-not $NoConsole) {
        $color = switch ($Level) {
            'SUCCESS' { 'Green'   }
            'ERROR'   { 'Red'     }
            'WARN'    { 'Yellow'  }
            'STEP'    { 'Cyan'    }
            default   { 'White'   }
        }
        Write-Host $entry -ForegroundColor $color
    }
}

# ── Module / connection helpers ───────────────────────────────────────────────

function Assert-RequiredModules {
    [CmdletBinding()]
    param([string[]]$ModuleNames)

    $missing = @()
    foreach ($name in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $name)) {
            $missing += $name
        }
    }
    if ($missing.Count -gt 0) {
        Write-OffboardingLog -Message "Missing modules: $($missing -join ', ')" -Level ERROR
        Write-OffboardingLog -Message "Install with: Install-Module $($missing -join ', ') -Scope CurrentUser" -Level INFO
        throw "Required PowerShell modules are not installed."
    }
    Write-OffboardingLog -Message "All required modules are available." -Level INFO
}

function Connect-OffboardingServices {
    [CmdletBinding()]
    param(
        [string[]]$GraphScopes,
        [string]$TenantId,
        [switch]$WhatIf
    )
    if ($WhatIf) {
        $tenantMsg = if ($TenantId) { " (tenant: $TenantId)" } else { ' (tenant: resolved via login)' }
        Write-OffboardingLog -Message "[WHATIF] Would connect to Microsoft Graph and Exchange Online$tenantMsg." -Level INFO
        return
    }

    Write-OffboardingLog -Message "Connecting to Microsoft Graph..." -Level STEP
    try {
        $mgParams = @{ Scopes = $GraphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $mgParams['TenantId'] = $TenantId }
        Connect-MgGraph @mgParams
        Write-OffboardingLog -Message "Connected to Microsoft Graph." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to connect to Microsoft Graph: $_" -Level ERROR
        throw
    }

    Write-OffboardingLog -Message "Connecting to Exchange Online..." -Level STEP
    try {
        $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($TenantId) { $exoParams['Organization'] = $TenantId }
        Connect-ExchangeOnline @exoParams
        Write-OffboardingLog -Message "Connected to Exchange Online." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to connect to Exchange Online: $_" -Level ERROR
        throw
    }
}

function Disconnect-OffboardingServices {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    Write-OffboardingLog -Message "Disconnected from all services." -Level INFO
}

# ── Step 1: Disable account + revoke sessions ─────────────────────────────────

function Disable-EntraUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$User,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 1 — Disabling account and revoking sessions for $($User.UserPrincipalName)" -Level STEP

    if ($WhatIf) {
        Write-OffboardingLog -Message "[WHATIF] Would disable account $($User.UserPrincipalName) in Entra ID." -Level INFO
        Write-OffboardingLog -Message "[WHATIF] Would revoke all active sign-in sessions." -Level INFO
        return
    }

    try {
        Update-MgUser -UserId $User.Id -AccountEnabled:$false -ErrorAction Stop
        Write-OffboardingLog -Message "Account disabled in Entra ID." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to disable account: $_" -Level ERROR
        throw
    }

    try {
        $null = Revoke-MgUserSignInSession -UserId $User.Id -ErrorAction Stop
        Write-OffboardingLog -Message "All active sign-in sessions revoked." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to revoke sessions: $_" -Level WARN
    }
}

# ── Step 2: Convert mailbox to Shared ─────────────────────────────────────────

function Convert-ToSharedMailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 2 — Converting mailbox to Shared for $UserPrincipalName" -Level STEP

    if ($WhatIf) {
        Write-OffboardingLog -Message "[WHATIF] Would convert mailbox to Shared type." -Level INFO
        return
    }

    try {
        Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
        Write-OffboardingLog -Message "Mailbox converted to Shared." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to convert mailbox to Shared: $_" -Level ERROR
        throw
    }
}

# ── Step 3: OOO + forwarding ──────────────────────────────────────────────────

function Set-OffboardingMailConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$DisplayName,
        [string]$ManagerEmail,
        [string]$ManagerName,
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$ForwardToManager,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 3 — Configuring OOO and mail forwarding for $UserPrincipalName" -Level STEP

    $hasManager = -not [string]::IsNullOrWhiteSpace($ManagerEmail)

    if ($hasManager) {
        $internalMsg = $Config.OOOInternal  -f $DisplayName, $ManagerName, $ManagerEmail, $Config.ITContactEmail, $Config.CompanyName
        $externalMsg = $Config.OOOExternal  -f $DisplayName, $ManagerName, $ManagerEmail, $Config.ITContactEmail, $Config.CompanyName
    }
    else {
        $internalMsg = $Config.OOOInternalNoManager -f $DisplayName, $null, $null, $Config.ITContactEmail, $Config.CompanyName
        $externalMsg = $Config.OOOExternalNoManager -f $DisplayName, $null, $null, $Config.ITContactEmail, $Config.CompanyName
    }

    if ($WhatIf) {
        Write-OffboardingLog -Message "[WHATIF] Would enable OOO (internal + external all)." -Level INFO
        if ($ForwardToManager -and $hasManager) {
            Write-OffboardingLog -Message "[WHATIF] Would forward mail to manager: $ManagerEmail (deliver to mailbox AND forward)." -Level INFO
        }
        return
    }

    try {
        Set-MailboxAutoReplyConfiguration `
            -Identity        $UserPrincipalName `
            -AutoReplyState  Enabled `
            -InternalMessage $internalMsg `
            -ExternalMessage $externalMsg `
            -ExternalAudience All `
            -ErrorAction Stop

        Write-OffboardingLog -Message "Out-of-Office reply configured." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to set OOO: $_" -Level WARN
    }

    if ($ForwardToManager -and $hasManager) {
        try {
            Set-Mailbox `
                -Identity                  $UserPrincipalName `
                -ForwardingSMTPAddress     $ManagerEmail `
                -DeliverToMailboxAndForward $true `
                -ErrorAction Stop

            Write-OffboardingLog -Message "Mail forwarding enabled → $ManagerEmail (copy kept in shared mailbox)." -Level SUCCESS
        }
        catch {
            Write-OffboardingLog -Message "Failed to configure mail forwarding: $_" -Level WARN
        }
    }
    elseif ($ForwardToManager -and -not $hasManager) {
        Write-OffboardingLog -Message "ForwardToManager requested but no manager email found — skipping forwarding." -Level WARN
    }
}

# ── Step 4: Remove M365 licenses ──────────────────────────────────────────────

function Remove-UserLicenses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$User,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 4 — Removing M365 licenses for $($User.UserPrincipalName)" -Level STEP

    try {
        $licenses = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
    }
    catch {
        Write-OffboardingLog -Message "Failed to retrieve licenses: $_" -Level ERROR
        return
    }

    if (-not $licenses -or $licenses.Count -eq 0) {
        Write-OffboardingLog -Message "No licenses assigned — nothing to remove." -Level INFO
        return
    }

    $skuIds = $licenses | Select-Object -ExpandProperty SkuId
    Write-OffboardingLog -Message "Found $($licenses.Count) license(s): $($licenses.SkuPartNumber -join ', ')" -Level INFO

    if ($WhatIf) {
        foreach ($lic in $licenses) {
            Write-OffboardingLog -Message "[WHATIF] Would remove license: $($lic.SkuPartNumber) ($($lic.SkuId))" -Level INFO
        }
        return
    }

    try {
        Set-MgUserLicense -UserId $User.Id -AddLicenses @() -RemoveLicenses $skuIds -ErrorAction Stop
        Write-OffboardingLog -Message "All $($licenses.Count) license(s) removed and returned to pool." -Level SUCCESS
    }
    catch {
        Write-OffboardingLog -Message "Failed to remove licenses: $_" -Level ERROR
        throw
    }
}

# ── Step 5: Remove Intune / Entra / AutoPilot devices ─────────────────────────

function Remove-UserDevices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$User,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 5 — Removing devices for $($User.UserPrincipalName)" -Level STEP

    # ── Intune managed devices ────────────────────────────────────────────────
    try {
        $intuneDevices = Get-MgDeviceManagementManagedDevice `
            -Filter "userPrincipalName eq '$($User.UserPrincipalName)'" `
            -ErrorAction Stop

        if ($intuneDevices.Count -gt 0) {
            Write-OffboardingLog -Message "Found $($intuneDevices.Count) Intune managed device(s)." -Level INFO
            foreach ($device in $intuneDevices) {
                if ($WhatIf) {
                    Write-OffboardingLog -Message "[WHATIF] Would retire/wipe Intune device: $($device.DeviceName) ($($device.Id))" -Level INFO
                }
                else {
                    try {
                        Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $device.Id -ErrorAction Stop
                        Write-OffboardingLog -Message "Removed Intune device: $($device.DeviceName)" -Level SUCCESS
                    }
                    catch {
                        Write-OffboardingLog -Message "Failed to remove Intune device $($device.DeviceName): $_" -Level WARN
                    }
                }
            }
        }
        else {
            Write-OffboardingLog -Message "No Intune managed devices found." -Level INFO
        }
    }
    catch {
        Write-OffboardingLog -Message "Failed to query Intune devices: $_" -Level WARN
    }

    # ── Entra ID registered devices ───────────────────────────────────────────
    try {
        $entraDevices = Get-MgUserRegisteredDevice -UserId $User.Id -ErrorAction Stop

        if ($entraDevices.Count -gt 0) {
            Write-OffboardingLog -Message "Found $($entraDevices.Count) Entra ID registered device(s)." -Level INFO
            foreach ($device in $entraDevices) {
                if ($WhatIf) {
                    Write-OffboardingLog -Message "[WHATIF] Would remove Entra registered device: $($device.Id)" -Level INFO
                }
                else {
                    try {
                        Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
                        Write-OffboardingLog -Message "Removed Entra device: $($device.Id)" -Level SUCCESS
                    }
                    catch {
                        Write-OffboardingLog -Message "Failed to remove Entra device $($device.Id): $_" -Level WARN
                    }
                }
            }
        }
        else {
            Write-OffboardingLog -Message "No Entra ID registered devices found." -Level INFO
        }
    }
    catch {
        Write-OffboardingLog -Message "Failed to query Entra devices: $_" -Level WARN
    }

    # ── AutoPilot devices ─────────────────────────────────────────────────────
    try {
        $allAutoPilot = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -ErrorAction Stop
        $userAutoPilot = $allAutoPilot | Where-Object { $_.UserPrincipalName -eq $User.UserPrincipalName }

        if ($userAutoPilot.Count -gt 0) {
            Write-OffboardingLog -Message "Found $($userAutoPilot.Count) AutoPilot device(s)." -Level INFO
            foreach ($device in $userAutoPilot) {
                if ($WhatIf) {
                    Write-OffboardingLog -Message "[WHATIF] Would remove AutoPilot device: $($device.SerialNumber) ($($device.Id))" -Level INFO
                }
                else {
                    try {
                        Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity `
                            -WindowsAutopilotDeviceIdentityId $device.Id `
                            -ErrorAction Stop
                        Write-OffboardingLog -Message "Removed AutoPilot device: $($device.SerialNumber)" -Level SUCCESS
                    }
                    catch {
                        Write-OffboardingLog -Message "Failed to remove AutoPilot device $($device.SerialNumber): $_" -Level WARN
                    }
                }
            }
        }
        else {
            Write-OffboardingLog -Message "No AutoPilot devices found for this user." -Level INFO
        }
    }
    catch {
        Write-OffboardingLog -Message "Failed to query AutoPilot devices: $_" -Level WARN
    }
}

# ── Step 6: Remove group memberships (skip dynamic groups) ────────────────────

function Remove-UserGroupMemberships {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$User,
        [switch]$WhatIf
    )

    Write-OffboardingLog -Message "STEP 6 — Removing group memberships for $($User.UserPrincipalName)" -Level STEP

    try {
        $memberships = Get-MgUserMemberOf -UserId $User.Id -All -ErrorAction Stop
    }
    catch {
        Write-OffboardingLog -Message "Failed to retrieve group memberships: $_" -Level ERROR
        return
    }

    $groups = $memberships | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }

    if (-not $groups -or $groups.Count -eq 0) {
        Write-OffboardingLog -Message "User has no group memberships." -Level INFO
        return
    }

    Write-OffboardingLog -Message "Found $($groups.Count) group membership(s) — evaluating eligibility..." -Level INFO

    $removed  = 0
    $skipped  = 0
    $failed   = 0

    foreach ($groupRef in $groups) {
        try {
            $group = Get-MgGroup -GroupId $groupRef.Id -Property 'Id,DisplayName,MembershipRule,GroupTypes' -ErrorAction Stop
        }
        catch {
            Write-OffboardingLog -Message "Could not retrieve group $($groupRef.Id): $_" -Level WARN
            $failed++
            continue
        }

        # Skip dynamic membership groups
        if (-not [string]::IsNullOrWhiteSpace($group.MembershipRule)) {
            Write-OffboardingLog -Message "SKIP dynamic group: $($group.DisplayName)" -Level INFO
            $skipped++
            continue
        }

        # Skip Microsoft 365 group owner role — can't remove via member API
        $isM365 = $group.GroupTypes -contains 'Unified'

        if ($WhatIf) {
            Write-OffboardingLog -Message "[WHATIF] Would remove from group: $($group.DisplayName) (M365=$isM365)" -Level INFO
            $removed++
            continue
        }

        try {
            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $User.Id -ErrorAction Stop
            Write-OffboardingLog -Message "Removed from group: $($group.DisplayName)" -Level SUCCESS
            $removed++
        }
        catch {
            # Handle already-not-a-member gracefully
            if ($_.Exception.Message -match '404|does not exist') {
                Write-OffboardingLog -Message "Already not a member of: $($group.DisplayName)" -Level INFO
            }
            else {
                Write-OffboardingLog -Message "Failed to remove from $($group.DisplayName): $_" -Level WARN
                $failed++
            }
        }
    }

    Write-OffboardingLog -Message "Group removal summary — Removed: $removed | Skipped (dynamic): $skipped | Failed: $failed" -Level $(if ($failed -gt 0) { 'WARN' } else { 'SUCCESS' })
}

# ── Summary report ────────────────────────────────────────────────────────────

function Write-OffboardingSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$DisplayName,
        [string]$ManagerEmail,
        [string]$CompanyName,
        [string]$TenantId,
        [switch]$ForwardToManager,
        [switch]$WhatIf,
        [DateTime]$StartTime
    )

    $elapsed    = (Get-Date) - $StartTime
    $mode       = if ($WhatIf) { 'WHATIF (dry run)' } else { 'LIVE' }
    $tenantLine = if ($TenantId) { $TenantId } else { 'resolved via login' }
    $clientLine = if ($CompanyName) { $CompanyName } else { '(not specified)' }

    $summary = @"

╔══════════════════════════════════════════════════════════════╗
║              OFFBOARDING COMPLETE — $mode
╠══════════════════════════════════════════════════════════════╣
  Client       : $clientLine
  Tenant       : $tenantLine
  User         : $DisplayName ($UserPrincipalName)
  Manager      : $(if ($ManagerEmail) { $ManagerEmail } else { 'Not specified' })
  Mail Forward : $(if ($ForwardToManager -and $ManagerEmail) { "→ $ManagerEmail" } else { 'Disabled' })
  Duration     : $([math]::Round($elapsed.TotalSeconds, 1))s
  Log          : $script:LogFile
╚══════════════════════════════════════════════════════════════╝
"@
    Write-Host $summary -ForegroundColor Cyan
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $summary -Encoding UTF8
    }
}

Export-ModuleMember -Function *
