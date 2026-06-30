#Requires -Version 7.0
<#
.SYNOPSIS
    Interactive menu for the 2X M365 user offboarding tool.

.DESCRIPTION
    Guides an admin or helpdesk user through offboarding options step by step,
    allowing them to toggle individual steps on or off before executing.
    Always offers a dry-run preview before making any live changes.

.EXAMPLE
    .\Start-OffboardingMenu.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Title)
    $width = 62
    $line  = '═' * $width
    Write-Host ""
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "  ── $Title" -ForegroundColor Yellow
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

# ── Step definitions ──────────────────────────────────────────────────────────

$StepDefinitions = [ordered]@{
    1 = 'Disable account + revoke all active sessions'
    2 = 'Convert mailbox to Shared Mailbox'
    3 = 'Set Out-of-Office reply + mail forwarding'
    4 = 'Remove all M365 licenses'
    5 = 'Remove devices (Intune / Entra / AutoPilot)'
    6 = 'Remove group memberships (static groups only)'
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

# ── Section 1: User details ───────────────────────────────────────────────────

Write-Header "SECTION 1 of 4 — USER DETAILS"

$upn = Read-ValidEmail -Prompt "User to offboard (UPN)" -Required

Write-Host ""
$companyName    = Read-Input  -Prompt "Client / company name"
$itContactEmail = Read-ValidEmail -Prompt "IT helpdesk email (used in OOO messages)"
$tenantId       = Read-Input  -Prompt "Tenant ID or domain (optional — skip to use login picker)"

# ── Section 2: Manager & mail ─────────────────────────────────────────────────

Write-Header "SECTION 2 of 4 — MANAGER & MAIL SETTINGS"

Write-Host "  The manager is looked up automatically from Entra ID." -ForegroundColor Gray
Write-Host "  You can override this or leave it blank to use the auto-resolved one." -ForegroundColor Gray
Write-Host ""

$forwardToManager = Read-YesNo -Prompt "Forward incoming mail to the user's manager?" -Default $true
$managerEmail     = ''

if ($forwardToManager) {
    $managerEmail = Read-ValidEmail -Prompt "Manager email override (optional — press Enter to auto-resolve)"
}

# ── Section 3: Step selection ─────────────────────────────────────────────────

Write-Header "SECTION 3 of 4 — SELECT STEPS TO RUN"

Write-Host "  Use the number keys to toggle steps on or off." -ForegroundColor Gray
Write-Host "  Press Enter when your selection is ready." -ForegroundColor Gray

# Default: all steps enabled
$selectedSteps = [ordered]@{}
foreach ($key in $StepDefinitions.Keys) { $selectedSteps[$key] = $true }

function Show-StepMenu {
    param([ordered]$Steps, [ordered]$Selected)
    Write-Host ""
    foreach ($key in $Steps.Keys) {
        $tick  = if ($Selected[$key]) { '[X]' } else { '[ ]' }
        $color = if ($Selected[$key]) { 'Green' } else { 'DarkGray' }
        Write-Host "    $key. $tick $($Steps[$key])" -ForegroundColor $color
    }
    Write-Host ""
    Write-Host "  Enter a step number to toggle, or press Enter to continue: " -NoNewline -ForegroundColor White
}

while ($true) {
    Show-StepMenu -Steps $StepDefinitions -Selected $selectedSteps
    $input = $Host.UI.ReadLine().Trim()
    if ([string]::IsNullOrWhiteSpace($input)) { break }
    if ($input -match '^\d+$' -and [int]$input -in $StepDefinitions.Keys) {
        $n = [int]$input
        $selectedSteps[$n] = -not $selectedSteps[$n]
        # Redraw in-place
        $lines = $StepDefinitions.Count + 3
        for ($i = 0; $i -lt $lines; $i++) {
            [Console]::SetCursorPosition(0, [Console]::CursorTop - 1)
            Write-Host (' ' * [Console]::WindowWidth)
            [Console]::SetCursorPosition(0, [Console]::CursorTop - 1)
        }
    }
    else {
        Write-Host "  Invalid input — enter a step number (1–6) or press Enter." -ForegroundColor Red
    }
}

$enabledSteps  = $selectedSteps.Keys | Where-Object { $selectedSteps[$_] }
$disabledSteps = $selectedSteps.Keys | Where-Object { -not $selectedSteps[$_] }

if ($enabledSteps.Count -eq 0) {
    Write-Host ""
    Write-Host "  No steps selected — nothing to do. Exiting." -ForegroundColor Red
    exit 0
}

# ── Section 4: Confirm ────────────────────────────────────────────────────────

Write-Header "SECTION 4 of 4 — CONFIRM & EXECUTE"

Write-Host "  Review your selections before proceeding:" -ForegroundColor Gray
Write-Host ""
Write-Host "  User          : $upn" -ForegroundColor White
Write-Host "  Client        : $(if ($companyName) { $companyName } else { '(not set)' })" -ForegroundColor White
Write-Host "  Tenant        : $(if ($tenantId) { $tenantId } else { 'resolved via login' })" -ForegroundColor White
Write-Host "  IT Email      : $(if ($itContactEmail) { $itContactEmail } else { '(not set)' })" -ForegroundColor White
Write-Host "  Mail Forward  : $(if ($forwardToManager) { "Yes$(if ($managerEmail) { " → $managerEmail" } else { ' (auto-resolve manager)' })" } else { 'No' })" -ForegroundColor White
Write-Host ""
Write-Host "  Steps to run  :" -ForegroundColor White
foreach ($n in $enabledSteps) {
    Write-Host "    [X] $n. $($StepDefinitions[$n])" -ForegroundColor Green
}
if ($disabledSteps.Count -gt 0) {
    foreach ($n in $disabledSteps) {
        Write-Host "    [ ] $n. $($StepDefinitions[$n])" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "  What would you like to do?" -ForegroundColor Yellow
Write-Host "    1. Dry run  — preview all actions, make NO changes" -ForegroundColor Cyan
Write-Host "    2. Execute  — run live (login prompts will appear)" -ForegroundColor Green
Write-Host "    3. Cancel   — exit without doing anything" -ForegroundColor DarkGray
Write-Host ""

$choice = ''
while ($choice -notin @('1', '2', '3')) {
    Write-Host "  Enter 1, 2, or 3 : " -NoNewline -ForegroundColor White
    $choice = $Host.UI.ReadLine().Trim()
}

if ($choice -eq '3') {
    Write-Host ""
    Write-Host "  Cancelled. No changes made." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

$isWhatIf   = ($choice -eq '1')
$skippedArr = $disabledSteps | ForEach-Object { [int]$_ }

# ── Build parameter splat and call the main script ────────────────────────────

$scriptPath = Join-Path $PSScriptRoot 'Invoke-UserOffboarding.ps1'
if (-not (Test-Path $scriptPath)) {
    Write-Host ""
    Write-Host "  ERROR: Invoke-UserOffboarding.ps1 not found at $scriptPath" -ForegroundColor Red
    exit 1
}

$params = @{
    UserPrincipalName = $upn
    ForwardToManager  = $forwardToManager
}

if ($companyName)      { $params['CompanyName']    = $companyName    }
if ($itContactEmail)   { $params['ITContactEmail'] = $itContactEmail }
if ($tenantId)         { $params['TenantId']       = $tenantId       }
if ($managerEmail)     { $params['ManagerEmail']   = $managerEmail   }
if ($skippedArr.Count -gt 0) { $params['SkipSteps'] = $skippedArr   }
if ($isWhatIf)         { $params['WhatIf']         = $true           }

Write-Host ""
if ($isWhatIf) {
    Write-Host "  ── DRY RUN — no changes will be made ─────────────────────" -ForegroundColor Cyan
} else {
    Write-Host "  ── EXECUTING — browser login prompts will appear ──────────" -ForegroundColor Green
}
Write-Host ""

& $scriptPath @params
