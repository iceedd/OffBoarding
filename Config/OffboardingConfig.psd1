@{
    # ── Company branding ──────────────────────────────────────────────────────
    CompanyName    = '2X'
    ITContactEmail = 'it@2x.com'
    ITContactName  = '2X IT Team'

    # ── Out-of-Office message templates ──────────────────────────────────────
    # Use {0} = DisplayName, {1} = ManagerName, {2} = ManagerEmail, {3} = ITContactEmail
    OOOInternal = @'
Thank you for your message.

{0} is no longer with {4}. For assistance, please contact {1} at {2}.

If your query relates to IT matters, please contact {3}.

Best regards,
{4} IT
'@

    OOOExternal = @'
Thank you for your email.

{0} has left {4} and is no longer able to respond to messages.
For further assistance, please contact {1} at {2}.

If you need immediate IT support, please reach {3}.

Best regards,
{4} IT
'@

    # OOO when no manager is specified
    OOOInternalNoManager = @'
Thank you for your message.

{0} is no longer with {4}. For assistance, please contact our IT team at {3}.

Best regards,
{4} IT
'@

    OOOExternalNoManager = @'
Thank you for your email.

{0} has left {4}. For further assistance, please contact our IT team at {3}.

Best regards,
{4} IT
'@

    # ── Logging ───────────────────────────────────────────────────────────────
    LogDirectory = '.\Logs'

    # ── Required PowerShell modules ───────────────────────────────────────────
    RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.DeviceManagement.Enrollment',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'ExchangeOnlineManagement'
    )

    # ── Graph scopes ──────────────────────────────────────────────────────────
    GraphScopes = @(
        'User.ReadWrite.All',
        'User.RevokeSessions.All',
        'Group.ReadWrite.All',
        'Directory.ReadWrite.All',
        'DeviceManagementManagedDevices.ReadWrite.All',
        'DeviceManagementServiceConfig.ReadWrite.All'
    )
}
