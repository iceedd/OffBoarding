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

### Interactive menu (recommended for helpdesk)

```powershell
.\Start-OffboardingMenu.ps1
```

Walks through 4 screens:

1. **User details** — UPN, client name, tenant, IT email
2. **Manager & mail** — whether to forward, manager email override
3. **Step selection** — toggle each of the 6 steps on or off with number keys
4. **Confirm & execute** — review everything, then choose dry run or live

No parameters needed — everything is prompted interactively.

---

### Command-line (for scripting / automation)

#### Dry run (no changes made)

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com -WhatIf
```

#### Full offboard for a client tenant

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com `
    -ForwardToManager `
    -CompanyName "Client A Ltd" `
    -ITContactEmail "helpdesk@clientA.com"
```

#### Skip the tenant picker (go straight to a specific tenant's login)

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com `
    -TenantId "clientA.onmicrosoft.com" `
    -CompanyName "Client A Ltd" `
    -ITContactEmail "helpdesk@clientA.com"
```

#### Skip specific steps (e.g. when re-running after a partial failure)

```powershell
.\Invoke-UserOffboarding.ps1 -UserPrincipalName jsmith@clientA.com -SkipSteps @(2, 3)
```

---

## File structure

```
OffBoarding/
├── Start-OffboardingMenu.ps1       # Interactive menu — start here
├── Invoke-UserOffboarding.ps1      # Direct command-line entry point
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
