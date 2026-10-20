# M365 Tenant Inventory

A read-only PowerShell script that inventories a Microsoft 365 tenant and produces accurate object counts for tenant-to-tenant migration scoping and quoting.

Run it against a tenant (yours, a client's, or in a screen-share session). It signs in interactively, reads names, counts, sizes and dates, and writes an Excel workbook, a read-only HTML copy and a zip of both. Counts are grouped the way migration tools bill them, so a quote rests on real numbers instead of what a client remembers.

- **Read-only.** Only `Get-*` cmdlets and Microsoft Graph `GET` requests. Nothing in the tenant is created, changed or deleted.
- **No content access.** It never reads mail, files, chats or calendars.
- **Nothing is uploaded.** Results are saved on the computer that ran the script.

## What it counts

Billable object categories, aligned with per-object migration licensing (for example AvePoint Fly):

| Category | Counted as one object |
|---|---|
| User mailboxes | Each user mailbox |
| Shared mailboxes | Each shared mailbox, licensed or not |
| OneDrive accounts | Each OneDrive |
| SharePoint sites | Each standalone site (not attached to a Team or Group) |
| Teams | Each Team, including its group, group mailbox, site and channel sites |
| Microsoft 365 Groups | Each group that is not a Team, including its mailbox and site |

Listed for information but not counted: guests, room and equipment mailboxes, distribution lists, mail-enabled security groups, contacts, inactive mailboxes, public folders, system SharePoint sites and deleted sites. The Reconciliation tab shows every site and mailbox that sits inside a Team or Group, which explains why the SharePoint count is lower than the admin centre shows.

## Requirements

- Windows with **Windows PowerShell 5.1** (starting it from PowerShell 7 relaunches it in 5.1)
- Internet access
- An administrator account for the tenant. Global Reader covers Exchange and SharePoint reading
- A **one-time Graph consent** from a Global Administrator or Privileged Role Administrator (a Global Reader cannot grant it). Without it, run with `-NoGraph`
- Microsoft Excel is not needed to run it, only to open the workbook

Missing modules are installed for the current user automatically: `ExchangeOnlineManagement`, `Microsoft.Online.SharePoint.PowerShell`, `Microsoft.Graph.Authentication` and `ImportExcel`. If the gallery install fails, a private copy is saved under `%LOCALAPPDATA%\Infinigence\Modules`.

## Usage

Save `Get-M365TenantInventory.ps1` to a folder, right-click it, choose Properties and tick **Unblock** if offered, then:

```powershell
cd C:\Inventory
powershell.exe -ExecutionPolicy Bypass -File .\Get-M365TenantInventory.ps1
```

You are asked to sign in three times with the same account. They are separate Microsoft services:

| Order | Service | Used for |
|---|---|---|
| 1 | Microsoft Graph | Users, licences, last sign-in, Teams, channels, domains |
| 2 | Exchange Online | Mailboxes, shared mailboxes, sizes, groups |
| 3 | SharePoint Online | SharePoint sites, OneDrive accounts, sizes |

The Graph sign-in requests these read-only scopes: `User.Read.All`, `Group.Read.All`, `Domain.Read.All`, `Organization.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `AuditLog.Read.All`.

When it finishes, a RESULTS summary is printed and the output folder opens.

### Parameters

| Parameter | Description | Default |
|---|---|---|
| `-ClientName` | Override the client name read from the tenant | Read from the tenant |
| `-OutputPath` | Folder in which the results subfolder is created | The script's folder (Desktop if not writable) |
| `-SmallObjectThresholdMB` | Objects smaller than this are flagged Review | `10` |
| `-StaleDays` | Users with no sign-in for this many days are flagged Review | `180` |
| `-SharePointAdminUrl` | SharePoint admin address, only if it cannot be derived | Derived from the tenant |
| `-NoGraph` | Skip Microsoft Graph (no consent needed, Entra details limited) | Off |
| `-SkipChannelDetail` | Skip the per-Team channel count for speed | Off |
| `-SkipGroupMailboxSize` | Skip group mailbox sizes for speed | Off |
| `-SkipModuleInstall` | Do not install missing modules automatically | Off |

```powershell
.\Get-M365TenantInventory.ps1 -ClientName "Contoso" -NoGraph
```

## Output

Each run creates its own subfolder, `M365-Inventory_<tenant>_<yyyyMMdd-HHmm>`, containing three files with the same base name:

- `.xlsx` the workbook
- `.html` a read-only snapshot of the same data
- `.zip` both of the above, ready to hand over

If the ImportExcel module cannot be loaded, CSV files are produced instead of a workbook and the HTML is still created.

### Workbook tabs

| Tab | Contents |
|---|---|
| Summary | Tenant details, live counts per category and per UPN domain, size to migrate, what is not counted |
| User Mailboxes, Shared Mailboxes | One row per mailbox: sizes (MB and GB), archive, hold, last logon, licences, flags |
| OneDrive | One row per OneDrive with owner status and size |
| SharePoint Sites | Standalone sites only |
| Teams, M365 Groups | One row per Team or Group with site and mailbox sizes |
| Domains | Every domain in the tenant with usage counts |
| Users | Member users with their licences (guests excluded) |
| Licenses | Each product bought, assigned and available |
| Informational | Items listed but not counted |
| Reconciliation | Sites and mailboxes hidden inside each Team or Group |
| Warnings | Anything the script could not read, and why |
| Sign-off | Confirmation block with the final counts |

Object tabs carry an **Action** column (Migrate / Skip / Review) plus Batch, Status and Notes columns, so the same workbook can become the migration tracker. Items are set to Review by default when they are under the small-object size, belong to a disabled or unlicensed user, or have had no sign-in for the stale period. Everything else starts as Migrate. Summary counts update as Action values change, and a Check column confirms every row is accounted for.

## Notes

- Designed for small and medium tenants. Tenants with thousands of sites or users take much longer.
- Last sign-in dates from Entra need an Entra ID P1/P2 licence. Without one, the Exchange last-logon date is used and the workbook says so.
- One object that cannot be read does not stop the run: it is skipped and listed on the Warnings tab.

## Troubleshooting

| Problem | What to do |
|---|---|
| Script is blocked | Unblock the file in Properties, or start it with `-ExecutionPolicy Bypass` |
| A module cannot be installed | Check connectivity, or run `Install-Module <name> -Scope CurrentUser` manually |
| Graph sign-in or consent fails | Have a Global Administrator approve the permissions once, or use `-NoGraph` |
| Sign-in is blocked | Conditional access may block PowerShell sign-in; allow it for the account during the run |
| SharePoint connection fails | Use `-SharePointAdminUrl https://<tenant>-admin.sharepoint.com` |
| Some mailboxes show size 0 | See the Warnings tab for the reason |
| Counts differ from the admin centre | Check Reconciliation and Informational: Team and Group sites and mailboxes are counted inside their parent |

## Version history

See the `.NOTES` block at the top of the script for a short changelog. Current version: **1.1.2**.

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 INFINIGENCE CONSULTING W.L.L.

This tool is read-only by design, but you run it at your own risk and should test it on a non-production tenant first if in doubt.
