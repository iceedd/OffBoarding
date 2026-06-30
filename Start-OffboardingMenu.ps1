#Requires -Version 7.0
<#
.SYNOPSIS
    Interactive menu for the 2X M365 user offboarding tool.

.DESCRIPTION
    Authenticates to a client tenant first, then guides the admin through
    offboarding options step by step. Steps can be toggled on or off before
    executing. A dry-run preview is always available before making live changes.

.EXAMPLE
    .\Start-OffboardingMenu.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Title)
    $line = '═' * 62
    Write-Host ""
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ""
}

function Read-Input {
    param(
        [string]$Prompt,
        [string]$Default = '',
        [switch]$Required
    )
    $displayDefault = if ($Default) { " [$Default]" } else { '' }
    while ($true) {
        Write-Host "  $Prompt$displayDefault : " -NoNewline -ForegroundColor White
        $value = $Host.UI.ReadLine().Trim()
        if ([string]::IsNullOrWhiteSpace($value)) { $value = $Default }
        if ($Required -and [string]::IsNullOrWhiteSpace($value)) {
            Write-Host "  This field is required." -ForegroundColor Red
            continue
        }
        return $value
    }
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $hint = if ($Default) { 'Y/n' } else { 'y/N' }
    while ($true) {
        Write-Host "  $Prompt [$hint] : " -NoNewline -ForegroundColor White
        $value = $Host.UI.ReadLine().Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
        if ($value -in @('y', 'yes')) { return $true  }
        if ($value -in @('n', 'no'))  { return $false }
        Write-Host "  Please enter Y or N." -ForegroundColor Red
    }
}

function Read-ValidEmail {
    param([string]$Prompt, [string]$Default = '', [switch]$Required)
    $emailPattern = '^[^@]+@[^@]+\.[^@]+$'
    while ($true) {
        $value = Read-Input -Prompt $Prompt -Default $Default -Required:$Required
        if ([string]::IsNullOrWhiteSpace($value)) { return $value }
        if ($value -match $emailPattern) { return $value }
        Write-Host "  Not a valid email address." -ForegroundColor Red
    }
}

# ── Step definitions (string keys to avoid OrderedDictionary index ambiguity) ─

$StepDefinitions = [ordered]@{
    '1' = 'Disable account + revoke all active sessions'
    '2' = 'Convert mailbox to Shared Mailbox'
    '3' = 'Set Out-of-Office reply + mail forwarding'
    '4' = 'Remove all M365 licenses'
    '5' = 'Remove devices (Intune / Entra / AutoPilot)'
    '6' = 'Remove group memberships (static groups only)'
}

# ── Banner ────────────────────────────────────────────────────────────────────

Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║           2X  —  M365 USER OFFBOARDING TOOL                 ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  This tool will guide you through offboarding a user from a" -ForegroundColor Gray
Write-Host "  Microsoft 365 tenant. You choose which steps to run." -ForegroundColor Gray
Write-Host ""
Write-Host "  Press Ctrl+C at any time to cancel." -ForegroundColor DarkGray
Write-Host ""

# ── Bootstrap module + config ─────────────────────────────────────────────────

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

$modulePath = Join-Path $scriptRoot 'Modules\OffboardingHelpers.psm1'
if (-not (Test-Path $modulePath)) { throw "Helper module not found at: $modulePath" }
Import-Module $modulePath -Force

$configPath = Join-Path $scriptRoot 'Config\OffboardingConfig.psd1'
if (-not (Test-Path $configPath)) { throw "Config not found at: $configPath" }
$Config = Import-PowerShellDataFile -Path $configPath

# ── Authenticate FIRST ────────────────────────────────────────────────────────

Write-Header "AUTHENTICATION — Sign in to the client tenant"

Write-Host "  A browser login window will open." -ForegroundColor Gray
Write-Host "  Sign in with a Global Admin account for the client tenant." -ForegroundColor Gray
Write-Host ""
Write-Host "  Press Enter to open the login window..." -NoNewline -ForegroundColor White
$null = $Host.UI.ReadLine()

try {
    Connect-MgGraph -Scopes $Config.GraphScopes -NoWelcome -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "  ERROR: Could not connect to Microsoft Graph." -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    exit 1
}

try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "  ERROR: Could not connect to Exchange Online." -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    exit 1
}

# Show connected account + tenant for confirmation
$mgContext = Get-MgContext
Write-Host ""
Write-Host "  Connected successfully:" -ForegroundColor Green
Write-Host "    Account : $($mgContext.Account)" -ForegroundColor White
Write-Host "    Tenant  : $($mgContext.TenantId)" -ForegroundColor White
Write-Host ""

$correct = Read-YesNo -Prompt "Is this the correct tenant?" -Default $true
if (-not $correct) {
    Write-Host ""
    Write-Host "  Disconnecting. Please re-run the script and sign into the correct tenant." -ForegroundColor Yellow
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    exit 0
}

# ── Step menu helper (defined once, outside the loop) ─────────────────────────

function Show-StepMenu {
    param(
        [System.Collections.Specialized.OrderedDictionary]$Steps,
        [System.Collections.Specialized.OrderedDictionary]$Selected
    )
    foreach ($key in $Steps.Keys) {
        $tick  = if ($Selected[$key]) { '[X]' } else { '[ ]' }
        $color = if ($Selected[$key]) { 'Green' } else { 'DarkGray' }
        Write-Host "    $key. $tick $($Steps[$key])" -ForegroundColor $color
    }
    Write-Host ""
    Write-Host "  Enter a step number to toggle, or press Enter to continue: " -NoNewline -ForegroundColor White
}

$scriptPath = Join-Path $scriptRoot 'Invoke-UserOffboarding.ps1'
if (-not (Test-Path $scriptPath)) {
    Write-Host "  ERROR: Invoke-UserOffboarding.ps1 not found." -ForegroundColor Red
    exit 1
}

# ── Main loop — stays in the same tenant session until the user quits ──────────

while ($true) {

    # ── Section 1: User details ───────────────────────────────────────────────

    Write-Header "USER DETAILS"

    $upn         = Read-ValidEmail -Prompt "User to offboard (UPN)" -Required
    $companyName = Read-Input      -Prompt "Client / company name (used in Out-of-Office messages)"

    # ── Section 2: Manager & mail ─────────────────────────────────────────────

    Write-Header "MANAGER & MAIL SETTINGS"

    Write-Host "  The manager is looked up automatically from Entra ID." -ForegroundColor Gray
    Write-Host "  You can override this or leave blank to auto-resolve." -ForegroundColor Gray
    Write-Host ""

    $forwardToManager = Read-YesNo      -Prompt "Forward incoming mail to the user's manager?" -Default $true
    $managerEmail     = ''
    if ($forwardToManager) {
        $managerEmail = Read-ValidEmail -Prompt "Manager email override (optional — press Enter to auto-resolve)"
    }

    # ── Section 3: Step selection ─────────────────────────────────────────────

    Write-Header "SELECT STEPS TO RUN"

    Write-Host "  Type a step number to toggle it on or off." -ForegroundColor Gray
    Write-Host "  Press Enter when your selection is ready." -ForegroundColor Gray
    Write-Host ""

    $selectedSteps = [ordered]@{}
    foreach ($key in $StepDefinitions.Keys) { $selectedSteps.Add($key, $true) }

    while ($true) {
        Show-StepMenu -Steps $StepDefinitions -Selected $selectedSteps
        $toggleInput = $Host.UI.ReadLine().Trim()

        if ([string]::IsNullOrWhiteSpace($toggleInput)) { break }

        if ($toggleInput -match '^\d+$' -and $StepDefinitions.Contains($toggleInput)) {
            $selectedSteps[$toggleInput] = -not $selectedSteps[$toggleInput]
            $linesToClear = $StepDefinitions.Count + 2
            for ($i = 0; $i -lt $linesToClear; $i++) {
                [Console]::SetCursorPosition(0, [Console]::CursorTop - 1)
                Write-Host (' ' * [Console]::WindowWidth)
                [Console]::SetCursorPosition(0, [Console]::CursorTop - 1)
            }
        }
        else {
            Write-Host "  Invalid — enter a number between 1 and 6, or press Enter to continue." -ForegroundColor Red
        }
    }

    $enabledSteps  = @($selectedSteps.Keys | Where-Object { $selectedSteps[$_] })
    $disabledSteps = @($selectedSteps.Keys | Where-Object { -not $selectedSteps[$_] })

    if ($enabledSteps.Count -eq 0) {
        Write-Host ""
        Write-Host "  No steps selected — skipping this user." -ForegroundColor Yellow
        Write-Host ""
    }
    else {
        # ── Confirm ───────────────────────────────────────────────────────────

        Write-Host ""
        Write-Host "  ══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  CONFIRM — Review before executing" -ForegroundColor Cyan
        Write-Host "  ══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Tenant        : $($mgContext.TenantId)" -ForegroundColor White
        Write-Host "  Account       : $($mgContext.Account)" -ForegroundColor White
        Write-Host "  User          : $upn" -ForegroundColor White
        Write-Host "  Client        : $(if ($companyName) { $companyName } else { '(not set)' })" -ForegroundColor White
        Write-Host "  Mail Forward  : $(if ($forwardToManager) { "Yes$(if ($managerEmail) { " → $managerEmail" } else { ' (auto-resolve manager)' })" } else { 'No' })" -ForegroundColor White
        Write-Host ""
        Write-Host "  Steps to run:" -ForegroundColor White
        foreach ($n in $enabledSteps)  { Write-Host "    [X] $n. $($StepDefinitions[$n])" -ForegroundColor Green   }
        foreach ($n in $disabledSteps) { Write-Host "    [ ] $n. $($StepDefinitions[$n])" -ForegroundColor DarkGray }

        Write-Host ""
        Write-Host "  What would you like to do?" -ForegroundColor Yellow
        Write-Host "    1. Dry run  — preview all actions, make NO changes" -ForegroundColor Cyan
        Write-Host "    2. Execute  — run live now" -ForegroundColor Green
        Write-Host "    3. Cancel   — skip this user and return to menu" -ForegroundColor DarkGray
        Write-Host ""

        $choice = ''
        while ($choice -notin @('1', '2', '3')) {
            Write-Host "  Enter 1, 2, or 3 : " -NoNewline -ForegroundColor White
            $choice = $Host.UI.ReadLine().Trim()
        }

        if ($choice -ne '3') {
            $isWhatIf    = ($choice -eq '1')
            $skippedNums = @($disabledSteps | ForEach-Object { [int]$_ })

            $params = @{
                UserPrincipalName = $upn
                ForwardToManager  = $forwardToManager
                SkipConnect       = $true
            }
            if ($companyName)             { $params['CompanyName']  = $companyName  }
            if ($managerEmail)            { $params['ManagerEmail'] = $managerEmail }
            if ($skippedNums.Count -gt 0) { $params['SkipSteps']   = $skippedNums  }
            if ($isWhatIf)                { $params['WhatIf']       = $true         }

            Write-Host ""
            if ($isWhatIf) {
                Write-Host "  ── DRY RUN — no changes will be made ─────────────────────" -ForegroundColor Cyan
            } else {
                Write-Host "  ── EXECUTING ───────────────────────────────────────────────" -ForegroundColor Green
            }
            Write-Host ""

            & $scriptPath @params
        }
    }

    # ── Prompt: another user or quit ──────────────────────────────────────────

    Write-Host ""
    Write-Host "  ══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    $another = Read-YesNo -Prompt "Offboard another user in this tenant?" -Default $true
    if (-not $another) {
        Write-Host ""
        Write-Host "  Disconnecting and exiting. Goodbye." -ForegroundColor Gray
        Write-Host ""
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
        break
    }
}
