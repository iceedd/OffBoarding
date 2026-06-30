#Requires -Version 7.0
<#
.SYNOPSIS
    Automates the complete M365 user offboarding process across any tenant.

.DESCRIPTION
    Designed for MSP/multi-tenant use. Prompts for an admin login on each run,
    scoping all actions to whichever tenant that admin belongs to.

    Performs the following steps in order:
      1. Disables the user account in Entra ID and revokes all active sessions
      2. Converts the user's mailbox to a Shared Mailbox
      3. Configures Out-of-Office reply and (optionally) mail forwarding to manager
      4. Removes all M365 licenses, returning them to the pool
      5. Removes all Intune managed, Entra registered, and AutoPilot devices
      6. Removes the user from all static security and distribution groups
         (dynamic membership groups are automatically skipped)

.PARAMETER UserPrincipalName
    UPN of the user to offboard (e.g. jsmith@clientA.com).

.PARAMETER ForwardToManager
    When specified, incoming mail is also forwarded to the user's manager.
    The manager is resolved automatically from the user's Entra ID profile.
    If no manager is found, forwarding is skipped and a warning is logged.

.PARAMETER ManagerEmail
    Optional. Override the auto-resolved manager email. Use when the Entra
    manager attribute is missing or incorrect.

.PARAMETER TenantId
    Optional. The target tenant's ID or primary domain (e.g. clientA.onmicrosoft.com).
    When provided, the login prompt goes directly to that tenant — no tenant
    picker shown. Useful when running non-interactively or scripting across
    multiple tenants in sequence.

.PARAMETER CompanyName
    Optional. The client's company name, used in Out-of-Office messages.
    Overrides the default value in OffboardingConfig.psd1.

.PARAMETER ITContactEmail
    Optional. The IT helpdesk email shown in Out-of-Office messages.
    Overrides the default value in OffboardingConfig.psd1.

.PARAMETER WhatIf
    Runs every step in simulation mode — no changes are made.
    All actions that would have been taken are logged.

.PARAMETER SkipSteps
    Array of step numbers (1–6) to skip. Useful when re-running after a
    partial failure.

.PARAMETER ConfigPath
    Path to the configuration .psd1 file.
    Defaults to .\Config\OffboardingConfig.psd1.

.EXAMPLE
    # Dry run — see what would happen without making any changes
    .\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com -WhatIf

.EXAMPLE
    # Full offboard for a specific client tenant, forwarding mail to manager
    .\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com `
        -ForwardToManager `
        -CompanyName "Client A Ltd" `
        -ITContactEmail "helpdesk@clientA.com"

.EXAMPLE
    # Skip the login picker by specifying the tenant, skip license step
    .\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com `
        -TenantId "clientA.onmicrosoft.com" `
        -CompanyName "Client A Ltd" `
        -ITContactEmail "helpdesk@clientA.com" `
        -SkipSteps @(4)

.NOTES
    Required modules  : Microsoft.Graph.*, ExchangeOnlineManagement
    Required roles    : Global Admin or User Admin + Exchange Admin + Intune Admin
    Author            : 2X IT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, HelpMessage = 'UPN of the user to offboard')]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [switch]$ForwardToManager,

    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$ManagerEmail,

    [string]$TenantId,

    [string]$CompanyName,

    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$ITContactEmail,

    [ValidateRange(1, 6)]
    [int[]]$SkipSteps = @(),

    # Set by Start-OffboardingMenu.ps1 when auth has already been done
    [switch]$SkipConnect,

    [string]$ConfigPath = '.\Config\OffboardingConfig.psd1'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$StartTime = Get-Date

# ── Bootstrap ─────────────────────────────────────────────────────────────────

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

$modulePath = Join-Path $scriptRoot 'Modules\OffboardingHelpers.psm1'
if (-not (Test-Path $modulePath)) {
    throw "Helper module not found at: $modulePath"
}
Import-Module $modulePath -Force

$resolvedConfig = Join-Path $scriptRoot ($ConfigPath -replace '^\.\/', '')
if (-not (Test-Path $resolvedConfig)) {
    throw "Config file not found at: $resolvedConfig"
}
$Config = Import-PowerShellDataFile -Path $resolvedConfig

# ── Apply runtime overrides to config ─────────────────────────────────────────

if ($PSBoundParameters.ContainsKey('CompanyName'))   { $Config['CompanyName']    = $CompanyName   }
if ($PSBoundParameters.ContainsKey('ITContactEmail')) { $Config['ITContactEmail'] = $ITContactEmail }

# ── Logging ───────────────────────────────────────────────────────────────────

$logDir = Join-Path $scriptRoot $Config.LogDirectory.TrimStart('.').TrimStart('\').TrimStart('/')
Initialize-OffboardingLog -LogDirectory $logDir -UserPrincipalName $UserPrincipalName

Write-OffboardingLog -Message "=== User Offboarding Started ===" -Level STEP
Write-OffboardingLog -Message "Target user  : $UserPrincipalName" -Level INFO
Write-OffboardingLog -Message "Client       : $($Config.CompanyName)" -Level INFO
Write-OffboardingLog -Message "Tenant       : $(if ($TenantId) { $TenantId } else { 'resolved via login' })" -Level INFO
Write-OffboardingLog -Message "Forward mail : $ForwardToManager" -Level INFO
Write-OffboardingLog -Message "WhatIf mode  : $($PSBoundParameters.ContainsKey('WhatIf') -or $WhatIfPreference)" -Level INFO
if ($SkipSteps.Count -gt 0) {
    Write-OffboardingLog -Message "Skipping steps: $($SkipSteps -join ', ')" -Level WARN
}

$isWhatIf = $PSBoundParameters.ContainsKey('WhatIf') -or $WhatIfPreference

# ── Pre-flight checks ─────────────────────────────────────────────────────────

Write-OffboardingLog -Message "Checking required modules..." -Level STEP
Assert-RequiredModules -ModuleNames $Config.RequiredModules

# ── Connect ───────────────────────────────────────────────────────────────────

if (-not $SkipConnect) {
    Connect-OffboardingServices -GraphScopes $Config.GraphScopes -TenantId $TenantId -WhatIf:$isWhatIf
} else {
    Write-OffboardingLog -Message "Using existing connection (authenticated via menu)." -Level INFO
}

# ── Resolve user ──────────────────────────────────────────────────────────────

Write-OffboardingLog -Message "Resolving user in Entra ID..." -Level STEP
try {
    $mgUser = Get-MgUser `
        -UserId $UserPrincipalName `
        -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled,Mail,Manager' `
        -ExpandProperty 'Manager' `
        -ErrorAction Stop
}
catch {
    Write-OffboardingLog -Message "User not found in Entra ID: $UserPrincipalName — $_" -Level ERROR
    Disconnect-OffboardingServices
    exit 1
}

Write-OffboardingLog -Message "Resolved: $($mgUser.DisplayName) (ID: $($mgUser.Id))" -Level SUCCESS

if (-not $mgUser.AccountEnabled -and -not $isWhatIf) {
    Write-OffboardingLog -Message "Account is already disabled. Proceeding with remaining steps..." -Level WARN
}

# ── Resolve manager ───────────────────────────────────────────────────────────

$resolvedManagerEmail = $ManagerEmail
$resolvedManagerName  = $null

if ([string]::IsNullOrWhiteSpace($resolvedManagerEmail)) {
    try {
        $managerRef = Get-MgUserManager -UserId $mgUser.Id -ErrorAction Stop
        if ($managerRef) {
            $manager = Get-MgUser -UserId $managerRef.Id -Property 'DisplayName,Mail' -ErrorAction Stop
            $resolvedManagerEmail = $manager.Mail
            $resolvedManagerName  = $manager.DisplayName
            Write-OffboardingLog -Message "Manager resolved: $resolvedManagerName ($resolvedManagerEmail)" -Level INFO
        }
    }
    catch {
        Write-OffboardingLog -Message "Could not resolve manager from Entra ID — $_" -Level WARN
    }
}
else {
    $resolvedManagerName = $resolvedManagerEmail
    Write-OffboardingLog -Message "Using provided manager email: $resolvedManagerEmail" -Level INFO
}

# ── Step 1: Disable account + revoke sessions ─────────────────────────────────

if (1 -notin $SkipSteps) {
    try {
        Disable-EntraUser -User $mgUser -WhatIf:$isWhatIf
    }
    catch {
        Write-OffboardingLog -Message "STEP 1 FAILED: $_" -Level ERROR
        Write-OffboardingLog -Message "Stopping — account was not disabled. Fix the issue and re-run." -Level ERROR
        Disconnect-OffboardingServices
        exit 1
    }
}
else {
    Write-OffboardingLog -Message "STEP 1 skipped by request." -Level WARN
}

# ── Step 2: Convert mailbox to Shared ─────────────────────────────────────────

if (2 -notin $SkipSteps) {
    try {
        Convert-ToSharedMailbox -UserPrincipalName $UserPrincipalName -WhatIf:$isWhatIf
    }
    catch {
        Write-OffboardingLog -Message "STEP 2 FAILED: $_" -Level ERROR
        # Non-fatal — continue with remaining steps
    }
}
else {
    Write-OffboardingLog -Message "STEP 2 skipped by request." -Level WARN
}

# ── Step 3: OOO + forwarding ──────────────────────────────────────────────────

if (3 -notin $SkipSteps) {
    Set-OffboardingMailConfig `
        -UserPrincipalName $UserPrincipalName `
        -DisplayName       $mgUser.DisplayName `
        -ManagerEmail      $resolvedManagerEmail `
        -ManagerName       $resolvedManagerName `
        -Config            $Config `
        -ForwardToManager:$ForwardToManager `
        -WhatIf:$isWhatIf
}
else {
    Write-OffboardingLog -Message "STEP 3 skipped by request." -Level WARN
}

# ── Step 4: Remove licenses ───────────────────────────────────────────────────

if (4 -notin $SkipSteps) {
    try {
        Remove-UserLicenses -User $mgUser -WhatIf:$isWhatIf
    }
    catch {
        Write-OffboardingLog -Message "STEP 4 FAILED: $_" -Level ERROR
    }
}
else {
    Write-OffboardingLog -Message "STEP 4 skipped by request." -Level WARN
}

# ── Step 5: Remove devices ────────────────────────────────────────────────────

if (5 -notin $SkipSteps) {
    Remove-UserDevices -User $mgUser -WhatIf:$isWhatIf
}
else {
    Write-OffboardingLog -Message "STEP 5 skipped by request." -Level WARN
}

# ── Step 6: Remove group memberships ─────────────────────────────────────────

if (6 -notin $SkipSteps) {
    Remove-UserGroupMemberships -User $mgUser -WhatIf:$isWhatIf
}
else {
    Write-OffboardingLog -Message "STEP 6 skipped by request." -Level WARN
}

# ── Cleanup + summary ─────────────────────────────────────────────────────────

Disconnect-OffboardingServices

Write-OffboardingSummary `
    -UserPrincipalName $UserPrincipalName `
    -DisplayName       $mgUser.DisplayName `
    -ManagerEmail      $resolvedManagerEmail `
    -CompanyName       $Config.CompanyName `
    -TenantId          $TenantId `
    -ForwardToManager:$ForwardToManager `
    -WhatIf:$isWhatIf `
    -StartTime         $StartTime
