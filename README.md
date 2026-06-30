# 2X User Offboarding Automation

PowerShell automation for secure, consistent, and complete M365 user offboarding.

## What it does

Runs 6 steps in sequence, logging every action:

| Step | Action |
|------|--------|
| 1 | Disables the account in Entra ID and immediately revokes all active sessions |
| 2 | Converts the user mailbox to a Shared Mailbox |
| 3 | Sets an Out-of-Office reply; optionally forwards mail to the user's manager |
| 4 | Removes all M365 licenses and returns them to the available pool |
| 5 | Removes Intune managed devices, Entra registered devices, and AutoPilot registrations |
| 6 | Removes the user from all static security and distribution groups (dynamic groups are skipped automatically) |

## Prerequisites

### PowerShell modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### Required admin roles

- **User Administrator** — disable accounts, remove licenses, group membership
- **Exchange Administrator** — shared mailbox conversion, OOO, forwarding
- **Intune Administrator** — device removal

---

## Usage

### Dry run (no changes made)

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@2x.com -WhatIf
```

### Standard offboard

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@2x.com
```

### Offboard and forward mail to manager

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@2x.com -ForwardToManager
```

The manager is resolved automatically from Entra ID. Use `-ManagerEmail` to override:

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@2x.com `
    -ForwardToManager `
    -ManagerEmail ceo@2x.com
```

### Skip specific steps

Useful when re-running after a partial failure:

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@2x.com -SkipSteps @(2, 3)
```

---

## File structure

```
OffBoarding/
├── Invoke-UserOffboarding.ps1      # Main entry point
├── Modules/
│   └── OffboardingHelpers.psm1    # Step functions and logging
├── Config/
│   └── OffboardingConfig.psd1     # OOO templates, company name, Graph scopes
└── Logs/                          # Auto-created; one log file per run
```

## Customisation

Edit `Config/OffboardingConfig.psd1` to change:

- `CompanyName` / `ITContactEmail` — used in OOO messages
- `OOOInternal` / `OOOExternal` — message templates (`{0}` = departed user, `{1}` = manager name, `{2}` = manager email, `{3}` = IT contact, `{4}` = company name)

## Logs

A timestamped log file is written to `Logs/` for every run, including WhatIf runs. Example:

```
Logs/Offboarding_jsmith@2x.com_20260630_143022.log
```
