#Requires -Version 5.1
<#
.SYNOPSIS
    INFINIGENCE Microsoft 365 Tenant Inventory - read-only discovery for tenant-to-tenant migration scoping.

.DESCRIPTION
    Signs in interactively to a Microsoft 365 tenant and produces a ZIP containing
      1. an Excel workbook  (counts, per-object detail, Migrate/Skip sign-off columns, tracker columns)
      2. a read-only HTML version of the same data

    Billable object categories (aligned to AvePoint Fly object licensing):
      User mailboxes, Shared mailboxes, OneDrive accounts, SharePoint sites, Teams, Microsoft 365 Groups.
    A Team or Microsoft 365 Group is ONE object. Its group mailbox, SharePoint site and channel sites
    are NOT counted again (they are listed on the Reconciliation tab).

    THIS SCRIPT IS READ-ONLY. It only uses Get-* cmdlets and Microsoft Graph GET requests.
    It never creates, changes or deletes anything in the tenant, and it never reads mail,
    file or chat content - only object names, counts, sizes and dates.

.PARAMETER ClientName
    Optional override for the client / organisation name. If omitted, the name is read from the tenant itself.

.PARAMETER OutputPath
    Folder in which the results subfolder is created. Default: the folder the script is saved in (the Desktop if that folder is not writable).

.PARAMETER SmallObjectThresholdMB
    Objects smaller than this are flagged for review. Default 10.

.PARAMETER StaleDays
    Users with no sign-in/logon for this many days are flagged for review. Default 180.

.PARAMETER SharePointAdminUrl
    Only needed if the SharePoint admin URL cannot be derived (for example a renamed tenant),
    e.g. https://contoso-admin.sharepoint.com

.PARAMETER NoGraph
    Skip Microsoft Graph (no Entra sign-in data, no Teams channel detail). Use this if the signed-in
    account cannot be granted Graph consent. Exchange and SharePoint data are still collected.

.PARAMETER SkipChannelDetail
    Skip the per-Team channel count (faster on tenants with many Teams).

.PARAMETER SkipGroupMailboxSize
    Do not read group mailbox sizes (faster). Team/Group totals then show the SharePoint site size only.

.PARAMETER SkipModuleInstall
    Do not install missing PowerShell modules automatically.

.NOTES
    Product   : INFINIGENCE M365 Tenant Inventory
    Version   : 1.1.2
    Changes   : 1.0.0 first release
                1.0.1 module install fallback (private copy) and safe module loading
                1.0.2 sign-in notice, Graph signs in first, clearer mailbox-statistics errors
                1.0.3 faster group listing, group mailbox fast-fail, -SkipGroupMailboxSize
                1.0.4 overall progress bar with step counter (Step n of 10)
                1.0.5 Graph module loads before Exchange and SharePoint modules (fixes Graph sign-in clash)
                1.0.6 one bad object no longer stops the run (skipped and listed on Warnings); errors show the script line
                1.0.7 client name read from the tenant (no prompt); tables sorted by UPN domain with a Domain column and a by-domain summary
                1.0.8 mailbox statistics call fixed (-Identity); lists converted to plain arrays; fuller error detail
                1.0.9 licence product names per user and a Licenses tab; owner IDs resolved to names
                1.1.0 sizes shown in GB as well as MB
                1.1.1 results saved next to the script; branding kept to the console and a small footer in the files
                1.1.2 each run saves into its own subfolder (xlsx, html and zip together)
    Author    : Farrukh Bashir, Founder & Managing Director
    Company   : INFINIGENCE CONSULTING W.L.L.   |   https://infinigence.me
    Contact   : farrukh@infinigence.me
    Copyright : Copyright (c) 2026 INFINIGENCE CONSULTING W.L.L. Released under the MIT License.

    This script is proprietary to INFINIGENCE CONSULTING W.L.L. and is provided for use in
    engagements with INFINIGENCE. It may not be copied, modified, redistributed or reused for any
    other purpose without prior written permission. Provided "AS IS", without warranty of any kind.
    Review the script before running it; it makes no changes to your tenant.

    RUN WITH   : Windows PowerShell 5.1 (the blue console). PowerShell 7 will relaunch itself in 5.1.
                 Start with:  powershell.exe -ExecutionPolicy Bypass -File .\Get-M365TenantInventory.ps1
                 (or right-click the file > Properties > Unblock).
    SIGN-IN    : Three interactive prompts (Exchange Online, Microsoft Graph, SharePoint Online) - use the same account.
    ROLES      : Global Reader is the target role. SharePoint and Exchange read access is included in it.
                 Microsoft Graph needs a one-time admin CONSENT for the read-only scopes below; a Global Reader
                 cannot grant consent. Either have a Global/Privileged Role Administrator approve the prompt once,
                 or run with -NoGraph.
    GRAPH SCOPES (all read-only): User.Read.All, Group.Read.All, Domain.Read.All, Organization.Read.All,
                 Team.ReadBasic.All, Channel.ReadBasic.All, AuditLog.Read.All (last sign-in; needs Entra ID P1/P2)
    MODULES    : ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell, Microsoft.Graph.Authentication,
                 ImportExcel - missing ones are installed for the current user only.
#>
[CmdletBinding()]
param(
    [string]$ClientName,
    [string]$OutputPath,
    [int]$SmallObjectThresholdMB = 10,
    [int]$StaleDays = 180,
    [string]$SharePointAdminUrl,
    [switch]$NoGraph,
    [switch]$SkipChannelDetail,
    [switch]$SkipGroupMailboxSize,
    [switch]$SkipModuleInstall
)

$ErrorActionPreference = 'Stop'
$script:BoundParams = $PSBoundParameters
$script:ScriptPath  = $PSCommandPath

$script:Version   = '1.1.2'
$script:Company   = 'INFINIGENCE CONSULTING W.L.L.'
$script:Web       = 'infinigence.me'
$script:Author    = 'Farrukh Bashir, Founder & Managing Director'
$script:Copyright = 'Copyright (c) 2026 INFINIGENCE CONSULTING W.L.L. Released under the MIT License.'
$script:Warnings  = New-Object System.Collections.Generic.List[object]
$script:GraphOk   = $false
$script:SignInOk  = $false
$script:SkuNameById = @{}
$script:ToolName  = 'INFINIGENCE M365 Tenant Inventory'

# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Show-Banner {
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host "   $($script:Company)" -ForegroundColor White
    Write-Host "   Microsoft 365 Tenant Inventory  v$($script:Version)   |   $($script:Web)" -ForegroundColor White
    Write-Host "   Author: $($script:Author)" -ForegroundColor Gray
    Write-Host "   $($script:Copyright)" -ForegroundColor Gray
    Write-Host '  ============================================================' -ForegroundColor Cyan
    Write-Host '   READ-ONLY: this tool makes no changes to your tenant and does' -ForegroundColor Green
    Write-Host '   not read mail, file or chat content - only names, sizes, dates.' -ForegroundColor Green
    Write-Host ''
}
function Write-Step { param([string]$Text) Write-Host "  > $Text" -ForegroundColor Cyan }
function Write-Note { param([string]$Text) Write-Host "    $Text" -ForegroundColor Gray }
$script:StageTotal = 10
function Start-Stage {
    param([int]$Number, [string]$Text)
    Write-Step "[$Number/$($script:StageTotal)] $Text"
    Write-Progress -Id 0 -Activity 'INFINIGENCE M365 Tenant Inventory' -Status "Step $Number of $($script:StageTotal): $Text" -PercentComplete ([int](100 * ($Number - 1) / $script:StageTotal))
}
function Add-Warn {
    param([string]$Area, [string]$Message, [string]$Level = 'Warning')
    $script:Warnings.Add([pscustomobject][ordered]@{ Area = $Area; Level = $Level; Message = $Message }) | Out-Null
    if ($Level -eq 'Warning') { Write-Host "    ! [$Area] $Message" -ForegroundColor Yellow }
}
function Invoke-WithRetry {
    param([scriptblock]$Script, [int]$Max = 4)
    for ($i = 1; $i -le $Max; $i++) {
        try { return (& $Script) }
        catch {
            $msg = $_.Exception.Message
            if ($i -ge $Max -or $msg -notmatch '429|throttl|timed out|503|502|504|Too Many') { throw }
            Start-Sleep -Seconds ([int]([math]::Pow(2, $i) * 2))
        }
    }
}
function Get-Prop {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}
function Test-RealGuid { param($G) $s = "$G"; return ($s -ne '' -and $s -ne '00000000-0000-0000-0000-000000000000') }
function ConvertTo-Bytes {
    param($V)
    if ($null -eq $V) { return [double]0 }
    try {
        $inner = Get-Prop $V 'Value'
        if ($inner -and $inner.PSObject.Methods['ToBytes']) { return [double]$inner.ToBytes() }
    } catch { }
    $s = [string]$V
    if ($s -match '\(([\d,]+) bytes\)') { return [double]($Matches[1] -replace ',', '') }
    return [double]0
}
function ConvertTo-DateTimeSafe {
    param($V)
    if ($null -eq $V -or "$V" -eq '') { return $null }
    try {
        if ($V -is [datetime]) { return $V.ToUniversalTime() }
        return ([datetimeoffset]::Parse([string]$V, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
    } catch { return $null }
}
function Format-DateStr { param($D) if ($D) { return ([datetime]$D).ToString('yyyy-MM-dd') } else { return '' } }
function ConvertTo-MB { param([double]$Bytes) return [math]::Round($Bytes / 1MB, 1) }
function ConvertTo-YesNo { param($B) if ($B -eq $true) { 'Yes' } elseif ($B -eq $false) { 'No' } else { 'Unknown' } }
function Get-LatestDate { param($A, $B) if ($A -and $B) { if ($A -gt $B) { return $A } else { return $B } } elseif ($A) { return $A } else { return $B } }

function Ensure-Module {
    param([string]$Name)
    # A private module folder (used if the normal per-user install does not end up on the module path)
    $localRoot = Join-Path $env:LOCALAPPDATA 'Infinigence\Modules'
    if ((Test-Path $localRoot) -and ($env:PSModulePath -notlike "*$localRoot*")) { $env:PSModulePath = "$localRoot;$env:PSModulePath" }
    if (Get-Module -ListAvailable -Name $Name) { return $true }
    if ($SkipModuleInstall) { return $false }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    try {
        Write-Step "Installing PowerShell module $Name (current user only)..."
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "    Install-Module failed for $Name : $($_.Exception.Message)" -ForegroundColor Yellow
    }
    if (Get-Module -ListAvailable -Name $Name) { return $true }

    # Fallback: download a private copy and add it to this session's module path
    try {
        Write-Step "Trying a private copy of $Name in $localRoot ..."
        if (-not (Test-Path $localRoot)) { New-Item -ItemType Directory -Path $localRoot -Force | Out-Null }
        Save-Module -Name $Name -Path $localRoot -Repository PSGallery -Force -ErrorAction Stop | Out-Null
        if ($env:PSModulePath -notlike "*$localRoot*") { $env:PSModulePath = "$localRoot;$env:PSModulePath" }
        if (Get-Module -ListAvailable -Name $Name) { return $true }
    } catch {
        Write-Host "    Could not download $Name : $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return $false
}
function Import-ModuleSafe {
    param([string]$Name, [switch]$DisableNameChecking)
    try {
        if ($DisableNameChecking) { Import-Module $Name -DisableNameChecking -ErrorAction Stop }
        else { Import-Module $Name -ErrorAction Stop }
        return $true
    } catch {
        Write-Host "    Could not load module $Name : $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

function Get-GraphPaged {
    param([string]$Uri)
    $out = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $u = $next
        $r = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri $u -OutputType PSObject -ErrorAction Stop }
        $vals = Get-Prop $r 'value'
        if ($null -ne $vals) { foreach ($v in $vals) { $out.Add($v) } }
        $next = Get-Prop $r '@odata.nextLink'
    }
    return $out.ToArray()
}

function Get-MbxStat {
    param([string]$Id, [switch]$Archive, [switch]$Fast)
    $res =[pscustomobject]@{ Bytes = [double]0; Items = 0; LastLogon = $null; Ok = $false }
    if (-not $Id) { return $res }
    try {
        $p = @{ Identity = $Id; Properties = @('TotalItemSize', 'ItemCount', 'LastLogonTime'); ErrorAction = 'Stop' }
        if ($Archive) { $p['Archive'] = $true }
        $mx = 4; if ($Fast) { $mx = 1 }
        $s = Invoke-WithRetry -Max $mx -Script { Get-EXOMailboxStatistics @p }
        if ($s) {
            $res.Bytes = ConvertTo-Bytes (Get-Prop $s 'TotalItemSize')
            $res.Items = [int](Nz (Get-Prop $s 'ItemCount') 0)
            $res.LastLogon = ConvertTo-DateTimeSafe (Get-Prop $s 'LastLogonTime')
            $res.Ok = $true
        }
    } catch {
        if ($Fast) { if (-not $script:GroupStatError) { $script:GroupStatError = $_.Exception.Message }; return $res }
        if (-not $script:FirstStatError) { $script:FirstStatError = $_.Exception.Message }
        try {
            $s2 = Invoke-WithRetry { if ($Archive) { Get-MailboxStatistics -Identity $Id -Archive -ErrorAction Stop } else { Get-MailboxStatistics -Identity $Id -ErrorAction Stop } }
            if ($s2) {
                $res.Bytes = ConvertTo-Bytes (Get-Prop $s2 'TotalItemSize')
                $res.Items = [int](Nz (Get-Prop $s2 'ItemCount') 0)
                $res.LastLogon = ConvertTo-DateTimeSafe (Get-Prop $s2 'LastLogonTime')
                $res.Ok = $true
            }
        } catch {
            if (-not $script:FirstStatError2) { $script:FirstStatError2 = $_.Exception.Message }
        }
    }
    return $res
}
function Nz { param($V, $Default = '') if ($null -eq $V -or "$V" -eq '') { return $Default } return $V }

function Get-UpnDomain {
    param($Upn)
    $u = [string]$Upn
    if ($u -match '@(.+)$') { return $Matches[1].ToLower() }
    return ''
}

$script:SkuFriendly = @{
    'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'; 'O365_BUSINESS_PREMIUM' = 'Microsoft 365 Business Standard'
    'SPB' = 'Microsoft 365 Business Premium'; 'O365_BUSINESS' = 'Microsoft 365 Apps for business'; 'SMB_BUSINESS' = 'Microsoft 365 Apps for business'
    'OFFICESUBSCRIPTION' = 'Microsoft 365 Apps for enterprise'; 'STANDARDPACK' = 'Office 365 E1'; 'ENTERPRISEPACK' = 'Office 365 E3'
    'ENTERPRISEPREMIUM' = 'Office 365 E5'; 'DESKLESSPACK' = 'Office 365 F3'; 'SPE_E3' = 'Microsoft 365 E3'; 'SPE_E5' = 'Microsoft 365 E5'
    'SPE_F1' = 'Microsoft 365 F3'; 'M365_F1' = 'Microsoft 365 F1'; 'EXCHANGESTANDARD' = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE' = 'Exchange Online (Plan 2)'; 'EXCHANGEDESKLESS' = 'Exchange Online Kiosk'
    'EXCHANGE_S_ESSENTIALS' = 'Exchange Online Essentials'; 'EXCHANGEARCHIVE_ADDON' = 'Exchange Online Archiving for Exchange Online'
    'EXCHANGEARCHIVE' = 'Exchange Online Archiving for Exchange Server'; 'EXCHANGE_S_FOUNDATION' = 'Exchange Foundation'
    'TEAMS_EXPLORATORY' = 'Microsoft Teams Exploratory'; 'MICROSOFT_TEAMS_EXPLORATORY_DEPT' = 'Microsoft Teams Exploratory Dept'
    'TEAMS_ESSENTIALS' = 'Microsoft Teams Essentials'; 'MCOEV' = 'Microsoft Teams Phone Standard'; 'MCOMEETADV' = 'Microsoft 365 Audio Conferencing'
    'PHONESYSTEM_VIRTUALUSER' = 'Microsoft Teams Phone Resource Account'; 'MEETING_ROOM' = 'Microsoft Teams Rooms Standard'
    'POWER_BI_STANDARD' = 'Microsoft Fabric (Free)'; 'POWER_BI_PRO' = 'Power BI Pro'; 'FLOW_FREE' = 'Microsoft Power Automate Free'
    'AAD_PREMIUM' = 'Microsoft Entra ID P1'; 'AAD_PREMIUM_P2' = 'Microsoft Entra ID P2'; 'EMS' = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM' = 'Enterprise Mobility + Security E5'; 'INTUNE_A' = 'Microsoft Intune Plan 1'; 'WINDOWS_STORE' = 'Windows Store for Business'
    'WIN10_PRO_ENT_SUB' = 'Windows 10/11 Enterprise E3'; 'PROJECT_P1' = 'Project Plan 1'; 'PROJECTPREMIUM' = 'Project Plan 5'
    'PROJECTPROFESSIONAL' = 'Project Plan 3'; 'VISIOCLIENT' = 'Visio Plan 2'; 'VISIO_PLAN1_DEPT' = 'Visio Plan 1'
    'ATP_ENTERPRISE' = 'Microsoft Defender for Office 365 (Plan 1)'; 'THREAT_INTELLIGENCE' = 'Microsoft Defender for Office 365 (Plan 2)'
    'DEFENDER_ENDPOINT_P1' = 'Microsoft Defender for Endpoint P1'; 'SHAREPOINTSTORAGE' = 'Office 365 Extra File Storage'
    'MICROSOFT_365_COPILOT' = 'Microsoft 365 Copilot'; 'RIGHTSMANAGEMENT' = 'Azure Information Protection Plan 1'
}
function Get-SkuFriendly {
    param([string]$Part)
    $k = $Part.ToUpper()
    if ($script:SkuFriendly.ContainsKey($k)) { return $script:SkuFriendly[$k] }
    return $Part
}
function Get-LicenseText {
    param($U)
    if (-not $U) { return '' }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($l in @(Get-Prop $U 'assignedLicenses')) {
        if (-not $l) { continue }
        $id = ([string](Get-Prop $l 'skuId')).ToLower()
        if ($script:SkuNameById.ContainsKey($id)) { $names.Add($script:SkuNameById[$id]) } else { $names.Add($id) }
    }
    return ($names -join '; ')
}

function New-Row {
    param([string[]]$Columns, [hashtable]$Values)
    $o = [ordered]@{}
    foreach ($c in $Columns) {
        if ($Values.ContainsKey($c)) { $o[$c] = $Values[$c] }
        elseif ($c -match '^(.*Size)GB$' -and $Values.ContainsKey($Matches[1] + 'MB')) {
            $mbv = "$($Values[$Matches[1] + 'MB'])"
            if ($mbv -ne '' -and $mbv -match '^-?[\d\.]+$') { $o[$c] = [math]::Round([double]$mbv / 1024, 2) } else { $o[$c] = '' }
        }
        elseif ($c -eq 'Status') { $o[$c] = 'Not started' }
        else { $o[$c] = '' }
    }
    return [pscustomobject]$o
}
function Get-DefaultAction { param($Flags) if ($Flags -and $Flags.Count -gt 0) { return 'Review' } else { return 'Migrate' } }

function Get-ColLetter {
    param([string[]]$Columns, [string]$Name)
    $idx = [array]::IndexOf($Columns, $Name) + 1
    if ($idx -lt 1) { return 'A' }
    $s = ''
    while ($idx -gt 0) {
        $m = ($idx - 1) % 26
        $s = ([string][char](65 + $m)) + $s
        $idx = [int][math]::Floor(($idx - 1) / 26)
    }
    return $s
}

# ----------------------------------------------------------------------------------------------
# Workbook helpers (ImportExcel)
# ----------------------------------------------------------------------------------------------
function Write-Sheet {
    param($Pkg, [string]$Name, $Rows, [string[]]$Columns)
    $rowsArr = New-Object System.Collections.ArrayList
    foreach ($rw in $Rows) { [void]$rowsArr.Add($rw) }
    $empty = ($rowsArr.Count -eq 0)
    if ($empty) {
        $d = [ordered]@{}
        foreach ($c in $Columns) { $d[$c] = '' }
        [void]$rowsArr.Add([pscustomobject]$d)
    }
    $null = $rowsArr | Select-Object -Property $Columns |
        Export-Excel -ExcelPackage $Pkg -WorksheetName $Name -AutoFilter -FreezeTopRow -BoldTopRow -AutoSize -PassThru
    $ws = $Pkg.Workbook.Worksheets[$Name]
    if ($empty) { $ws.DeleteRow(2) }
    $last = [math]::Max($rowsArr.Count, 1) + 1
    if ($Columns -contains 'Action') {
        try {
            $aCol = Get-ColLetter $Columns 'Action'
            $dv = $ws.DataValidations.AddListValidation("$($aCol)2:$($aCol)$($last + 200)")
            foreach ($v in @('Migrate', 'Skip', 'Review')) { $dv.Formula.Values.Add($v) | Out-Null }
            $dv.ShowErrorMessage = $true
            $dv.ErrorTitle = 'Action'
            $dv.Error = 'Choose Migrate, Skip or Review.'
            $sCol = Get-ColLetter $Columns 'Status'
            $dv2 = $ws.DataValidations.AddListValidation("$($sCol)2:$($sCol)$($last + 200)")
            foreach ($v in @('Not started', 'Pre-sync', 'Cutover scheduled', 'Completed', 'On hold', 'Skipped')) { $dv2.Formula.Values.Add($v) | Out-Null }
            $dv2.ShowErrorMessage = $false
            $rng = "$($aCol)2:$($aCol)$($last + 200)"
            $cf1 = $ws.ConditionalFormatting.AddEqual($ws.Cells[$rng]); $cf1.Formula = '"Review"'
            $cf1.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cf1.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 235, 156)
            $cf2 = $ws.ConditionalFormatting.AddEqual($ws.Cells[$rng]); $cf2.Formula = '"Skip"'
            $cf2.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cf2.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(217, 217, 217)
        } catch { }
    }
    try {
        foreach ($wide in @('Flags', 'Notes')) {
            if ($Columns -contains $wide) { $ws.Column([array]::IndexOf($Columns, $wide) + 1).Width = 45 }
        }
    } catch { }
}

# ----------------------------------------------------------------------------------------------
# HTML builder (self-contained, read-only, no external resources)
# ----------------------------------------------------------------------------------------------
function Build-Html {
    param($Model)
    $json = $Model | ConvertTo-Json -Depth 8 -Compress
    $json = $json.Replace('<', '<').Replace('>', '>').Replace('&', '&')
    $tpl = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Microsoft 365 Tenant Inventory</title>
<style>
:root{--bg:#f4f6f9;--card:#fff;--ink:#17212b;--mute:#5b6a7a;--line:#dbe1e8;--accent:#0b5cab;--chip:#e8f0fa;--rev:#fff3cd;--skip:#e6e8eb}
@media (prefers-color-scheme:dark){:root{--bg:#0f141b;--card:#161d26;--ink:#e5eaf0;--mute:#93a1b2;--line:#283341;--accent:#6aa9ee;--chip:#1d2a3a;--rev:#4a3d12;--skip:#2a313a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif}
header{padding:20px 24px 8px}header h1{margin:0;font-size:22px}header .sub{color:var(--mute);margin-top:2px}
.badge{display:inline-block;background:var(--chip);color:var(--accent);border-radius:999px;padding:2px 10px;font-size:12px;margin-top:8px}
nav{display:flex;flex-wrap:wrap;gap:6px;padding:8px 24px}nav button{border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:8px;padding:6px 12px;cursor:pointer;font:inherit}
nav button.on{background:var(--accent);color:#fff;border-color:var(--accent)}nav button span{opacity:.7;margin-left:6px;font-size:12px}
main{padding:8px 24px 32px}.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin-bottom:14px;overflow:auto}
.card h2{margin:0 0 8px;font-size:15px}table{border-collapse:collapse;width:100%}th,td{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;white-space:nowrap}
td.wrapc{white-space:normal}th{position:sticky;top:0;background:var(--card);cursor:pointer;font-weight:600;color:var(--mute)}
.kv td:first-child{color:var(--mute);width:240px}.bar{display:flex;gap:12px;align-items:center;margin-bottom:8px}
input[type=search]{padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);min-width:260px}
.cnt{color:var(--mute)}.wrap{max-height:70vh;overflow:auto;border:1px solid var(--line);border-radius:8px}
td.act.review{background:var(--rev)}td.act.skip{background:var(--skip)}.num{text-align:right}
ul.notes{margin:0;padding-left:20px}ul.notes li{margin:3px 0}footer{padding:14px 24px 28px;color:var(--mute);font-size:12px}
</style></head><body>
<header><h1 id="ttl"></h1><div class="sub" id="sub"></div><span class="badge">Read-only snapshot as of the run date. Migrate / Skip decisions are made in the Excel workbook.</span></header>
<nav id="nav"></nav><main id="main"></main>
<footer id="foot"></footer>
<script id="d" type="application/json">__DATA__</script>
<script>
(function(){
var D=JSON.parse(document.getElementById('d').textContent);
function el(t,c,x){var e=document.createElement(t);if(c)e.className=c;if(x!==undefined&&x!==null)e.textContent=x;return e;}
document.getElementById('ttl').textContent='Microsoft 365 Tenant Inventory - '+D.client;
document.getElementById('sub').textContent='Run date: '+D.generated;
document.getElementById('foot').textContent='Generated by '+D.tool+' v'+D.version;
var nav=document.getElementById('nav'),main=document.getElementById('main'),btns=[];
function card(title){var c=el('div','card');if(title)c.append(el('h2',null,title));return c;}
function simpleTable(cols,rows,cls){var t=el('table',cls),h=el('tr');cols.forEach(function(c){h.append(el('th',null,c));});var th=el('thead');th.append(h);t.append(th);var b=el('tbody');rows.forEach(function(r){var tr=el('tr');r.forEach(function(v,i){var td=el('td',i>0&&typeof v==='number'?'num':null,v===null?'':String(v));tr.append(td);});b.append(tr);});t.append(b);return t;}
function renderSummary(){
 main.textContent='';
 var c1=card('Tenant');c1.append(simpleTable(['Item','Value'],D.meta,'kv'));main.append(c1);
 var c2=card('Billable objects (as run)');c2.append(simpleTable(D.summaryCols,D.summary));main.append(c2);
 if(D.byDomain&&D.byDomain.length){var c5=card('Objects by UPN domain (as run)');c5.append(simpleTable(D.byDomainCols,D.byDomain));main.append(c5);}
 var c3=card('Not counted (for information)');c3.append(simpleTable(['Category','Count'],D.notCounted));main.append(c3);
 var c4=card('How objects are counted');var ul=el('ul','notes');D.notes.forEach(function(n){ul.append(el('li',null,n));});c4.append(ul);main.append(c4);
}
function renderTab(tab){
 main.textContent='';
 var c=card(tab.name);var bar=el('div','bar');var q=el('input');q.type='search';q.placeholder='Search this tab...';var cnt=el('span','cnt');bar.append(q,cnt);c.append(bar);
 var wrap=el('div','wrap'),tbl=el('table'),thead=el('thead'),trh=el('tr'),tbody=el('tbody');
 var sc=-1,asc=true;
 tab.columns.forEach(function(cn,i){var th=el('th',null,cn);th.addEventListener('click',function(){if(sc===i)asc=!asc;else{sc=i;asc=true;}draw();});trh.append(th);});
 thead.append(trh);tbl.append(thead,tbody);wrap.append(tbl);c.append(wrap);main.append(c);
 function draw(){
  var term=q.value.toLowerCase();
  var r=tab.rows.filter(function(row){return !term||row.join(' ').toLowerCase().indexOf(term)>=0;});
  if(sc>=0){r=r.slice().sort(function(a,b){var x=a[sc],y=b[sc];var nx=parseFloat(x),ny=parseFloat(y);var res;if(!isNaN(nx)&&!isNaN(ny)&&String(x).trim()!==''&&String(y).trim()!==''){res=nx-ny;}else{res=String(x===null?'':x).localeCompare(String(y===null?'':y));}return asc?res:-res;});}
  tbody.textContent='';
  r.forEach(function(row){var tr=el('tr');row.forEach(function(v,i){var cn=tab.columns[i];var td=el('td',null,v===null?'':String(v));if(cn==='Action'){td.className='act '+String(v).toLowerCase();}if(cn==='Flags'||cn==='Message'||cn==='Notes'){td.style.whiteSpace='normal';td.style.minWidth='220px';}if(typeof v==='number'){td.className=(td.className+' num').trim();}tr.append(td);});tbody.append(tr);});
  cnt.textContent=r.length+' of '+tab.rows.length+' rows';
 }
 q.addEventListener('input',draw);draw();
}
function show(i){btns.forEach(function(b,j){b.className=(i===j)?'on':'';});if(i===0)renderSummary();else renderTab(D.tabs[i-1]);}
var all=[{name:'Summary',n:null}].concat(D.tabs.map(function(t){return{name:t.name,n:t.rows.length};}));
all.forEach(function(t,i){var b=el('button',null,t.name);if(t.n!==null){b.append(el('span',null,String(t.n)));}b.addEventListener('click',function(){show(i);});nav.append(b);btns.push(b);});
show(0);
})();
</script></body></html>
'@
    return $tpl.Replace('__DATA__', $json)
}

# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
function Invoke-Inventory {
    Show-Banner
    if ($env:OS -ne 'Windows_NT') { throw 'This script must be run on Windows (Windows PowerShell 5.1).' }

    # PowerShell 7 -> relaunch in Windows PowerShell 5.1 (the SharePoint module requires it)
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Write-Step 'PowerShell 7 detected - relaunching in Windows PowerShell 5.1...'
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$($script:ScriptPath)`"")
        foreach ($k in $script:BoundParams.Keys) {
            $v = $script:BoundParams[$k]
            if ($v -is [switch]) { if ($v.IsPresent) { $argList += "-$k" } }
            else { $argList += "-$k"; $argList += "`"$v`"" }
        }
        & powershell.exe @argList
        return
    }

    $script:ClientName = $ClientName   # blank = taken from the tenant after sign-in
    if ($OutputPath) { $script:OutDir = $OutputPath }
    elseif ($script:ScriptPath) { $script:OutDir = Split-Path -Parent $script:ScriptPath }
    else { $script:OutDir = (Get-Location).Path }
    if (-not (Test-Path $script:OutDir)) { New-Item -ItemType Directory -Path $script:OutDir -Force | Out-Null }
    try {
        $probe = Join-Path $script:OutDir ('.w' + [guid]::NewGuid().ToString('N'))
        Set-Content -Path $probe -Value '' -ErrorAction Stop
        Remove-Item $probe -Force -ErrorAction SilentlyContinue
    } catch {
        $script:OutDir = [Environment]::GetFolderPath('Desktop')
        Write-Note "The script folder is not writable - saving to the Desktop instead."
    }

    # ---- Modules ----------------------------------------------------------------------------
    Start-Stage 1 'Checking PowerShell modules'
    $haveEXO = Ensure-Module 'ExchangeOnlineManagement'
    if (-not $haveEXO) { throw 'ExchangeOnlineManagement is required but is not available. Install it and run again.' }
    $haveSPO = Ensure-Module 'Microsoft.Online.SharePoint.PowerShell'
    $haveGraph = $false
    if (-not $NoGraph) { $haveGraph = Ensure-Module 'Microsoft.Graph.Authentication' }
    $haveXlsx = Ensure-Module 'ImportExcel'
    # Import order matters: Graph loads its sign-in libraries first; Exchange and SharePoint load theirs later,
    # just before their own sign-in, so the three modules do not clash over shared libraries.
    if (-not $haveSPO) { Add-Warn 'SharePoint' 'SharePoint module not available - OneDrive and SharePoint site counts are MISSING.' }
    if ($haveGraph) { $haveGraph = Import-ModuleSafe 'Microsoft.Graph.Authentication' }
    if ($haveGraph -ne $true -and -not $NoGraph) { Add-Warn 'Graph' 'Microsoft Graph module not available - Entra details will be limited.' }
    if ($haveXlsx) { $haveXlsx = Import-ModuleSafe 'ImportExcel' }
    if (-not $haveXlsx) { Add-Warn 'Output' 'ImportExcel module not available - CSV files will be produced instead of an Excel workbook.' }

    try {
        # ---- Sign-ins -----------------------------------------------------------------------
        Start-Stage 2 'Signing in to Microsoft 365 (three sign-ins)'
        Write-Host ''
        Write-Host '  YOU WILL BE ASKED TO SIGN IN THREE TIMES (use the same account each time):' -ForegroundColor White
        Write-Host '    1. Microsoft Graph      - users, licences, groups, Teams, domains (read-only)' -ForegroundColor Gray
        Write-Host '    2. Exchange Online      - mailboxes, shared mailboxes, sizes' -ForegroundColor Gray
        Write-Host '    3. SharePoint Online    - SharePoint sites and OneDrive accounts, sizes' -ForegroundColor Gray
        Write-Host '  These are separate Microsoft services. After the first sign-in the others are usually a single click.' -ForegroundColor Gray
        Write-Host ''

        if ($haveGraph) {
            Write-Step 'Sign in 1 of 3: Microsoft Graph (read-only scopes)'
            try {
                $scopes = @('User.Read.All', 'Group.Read.All', 'Domain.Read.All', 'Organization.Read.All', 'Team.ReadBasic.All', 'Channel.ReadBasic.All', 'AuditLog.Read.All')
                Connect-MgGraph -Scopes $scopes -NoWelcome -ContextScope Process -ErrorAction Stop
                $script:GraphOk = $true
            } catch {
                Add-Warn 'Graph' "Microsoft Graph sign-in/consent failed ($($_.Exception.Message)). Continuing without Entra data: last sign-in, licence status, guests and Team channel detail will be limited. Ask a Global Administrator to approve the consent once, or run with -NoGraph."
            }
        }

        Write-Step 'Sign in 2 of 3: Exchange Online'
        if (-not (Import-ModuleSafe 'ExchangeOnlineManagement')) { throw 'ExchangeOnlineManagement could not be loaded. Install it manually (Install-Module ExchangeOnlineManagement -Scope CurrentUser) and run again.' }
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

        # ---- Tenant identity ----------------------------------------------------------------
        Start-Stage 3 'Reading tenant details and domains'
        $conn = $null
        try { $conn = Get-ConnectionInformation | Select-Object -First 1 } catch { }
        $runBy = [string](Nz (Get-Prop $conn 'UserPrincipalName') '')
        $tenantId = [string](Nz (Get-Prop $conn 'TenantID') '')
        $accepted = @(Invoke-WithRetry { Get-AcceptedDomain -ErrorAction Stop })
        $initialDomain = [string](@($accepted | Where-Object { $_.DomainName -like '*.onmicrosoft.com' -and $_.DomainName -notlike '*.mail.onmicrosoft.com' } | Select-Object -First 1 -ExpandProperty DomainName))
        $defaultDomain = [string](@($accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1 -ExpandProperty DomainName))
        $tenantName = $initialDomain

        $graphDomains = @()
        if ($script:GraphOk) {
            try {
                $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -OutputType PSObject
                $o1 = @($org.value)[0]
                if ($o1) { $tenantId = [string](Nz $o1.id $tenantId); $tenantName = [string](Nz $o1.displayName $tenantName) }
            } catch { Add-Warn 'Graph' "Could not read organization details: $($_.Exception.Message)" }
            try { $graphDomains = @(Get-GraphPaged 'https://graph.microsoft.com/v1.0/domains') }
            catch { Add-Warn 'Domains' "Could not read domains from Graph (Exchange accepted domains used instead): $($_.Exception.Message)" }
        }

        if (-not $script:ClientName) {
            $cn = ''
            if ($script:GraphOk -and $tenantName -and $tenantName -ne $initialDomain) { $cn = [string]$tenantName }
            if (-not $cn) { try { $oc = Get-OrganizationConfig -ErrorAction Stop; $cn = [string](Get-Prop $oc 'DisplayName') } catch { } }
            if (-not $cn -and $initialDomain) { $cn = $initialDomain.Split('.')[0] }
            if (-not $cn) { $cn = 'Client' }
            $script:ClientName = $cn
            Write-Note "Client name taken from the tenant: $cn"
        }

        # ---- Entra users --------------------------------------------------------------------
        $usersById = @{}
        $usersByUpn = @{}
        $guestCount = 0
        $memberUsers = @()
        if ($script:GraphOk) {
            Start-Stage 4 'Reading Entra users'
            $sel = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,createdDateTime,assignedLicenses,onPremisesSyncEnabled'
            $allUsers = @()
            try {
                $allUsers = @(Get-GraphPaged ('https://graph.microsoft.com/v1.0/users?$select=' + $sel + ',signInActivity&$top=120'))
                $script:SignInOk = $true
            } catch {
                Add-Warn 'Entra' 'Last sign-in (Entra) is unavailable - the tenant needs Entra ID P1/P2 and AuditLog.Read.All. Exchange last-logon time is used instead.'
                try { $allUsers = @(Get-GraphPaged ('https://graph.microsoft.com/v1.0/users?$select=' + $sel + '&$top=999')) }
                catch { Add-Warn 'Entra' "Could not read users from Graph: $($_.Exception.Message)" ; $script:GraphOk = $false }
            }
            foreach ($u in $allUsers) {
                if ([string]$u.userType -eq 'Guest') { $guestCount++; continue }
                $memberUsers += $u
                $usersById[([string]$u.id).ToLower()] = $u
                $usersByUpn[([string]$u.userPrincipalName).ToLower()] = $u
            }
        } else {
            try { $guestCount = @(Get-User -RecipientTypeDetails GuestMailUser -ResultSize Unlimited -ErrorAction Stop).Count } catch { }
        }

        # ---- Licences (subscribed products in the tenant) ----------------------------------------
        $cLic = @('Product', 'SkuPartNumber', 'Purchased', 'Assigned', 'Available', 'Status')
        $licRows = New-Object System.Collections.Generic.List[object]
        if ($script:GraphOk) {
            try {
                foreach ($k in @(Get-GraphPaged 'https://graph.microsoft.com/v1.0/subscribedSkus')) {
                    $pn = [string]$k.skuPartNumber
                    $fn = Get-SkuFriendly $pn
                    $script:SkuNameById[([string]$k.skuId).ToLower()] = $fn
                    $pre = Get-Prop $k 'prepaidUnits'
                    $en = [int](Nz (Get-Prop $pre 'enabled') 0)
                    $used = [int](Nz (Get-Prop $k 'consumedUnits') 0)
                    $licRows.Add((New-Row $cLic @{ Product = $fn; SkuPartNumber = $pn; Purchased = $en; Assigned = $used; Available = ($en - $used); Status = [string](Get-Prop $k 'capabilityStatus') }))
                }
            } catch { Add-Warn 'Licences' "Licence products could not be read; licences show Yes/No only: $($_.Exception.Message)" }
        }

        function Get-UserSignIn {
            param($U)
            if (-not $U) { return $null }
            $sa = Get-Prop $U 'signInActivity'
            if (-not $sa) { return $null }
            $a = ConvertTo-DateTimeSafe (Get-Prop $sa 'lastSuccessfulSignInDateTime')
            $b = ConvertTo-DateTimeSafe (Get-Prop $sa 'lastSignInDateTime')
            return (Get-LatestDate $a $b)
        }
        $now = (Get-Date).ToUniversalTime()
        $staleCut = $now.AddDays(-1 * $StaleDays)

        # ---- Exchange mailboxes -------------------------------------------------------------
        Start-Stage 5 'Reading Exchange Online mailboxes and sizes'
        $allMbx = @(Invoke-WithRetry { Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox -Properties ArchiveStatus, LitigationHoldEnabled, AccountDisabled, WhenCreated, ExternalDirectoryObjectId -ErrorAction Stop })

        $cTrk = @('Action', 'Batch', 'Status', 'Notes')
        $cMbx = @('Name', 'UserPrincipalName', 'Domain', 'PrimarySmtp', 'AccountEnabled', 'Licensed', 'Licenses', 'MailboxSizeMB', 'MailboxSizeGB', 'ItemCount', 'Archive', 'ArchiveSizeMB', 'ArchiveSizeGB', 'TotalSizeMB', 'TotalSizeGB', 'LitigationHold', 'LastLogonExchange', 'LastSignInEntra', 'Created', 'Flags') + $cTrk
        $userRows = New-Object System.Collections.Generic.List[object]
        $sharedRows = New-Object System.Collections.Generic.List[object]
        $infoRows = New-Object System.Collections.Generic.List[object]
        $smtpDomainCount = @{}
        $statFail = 0
        $i = 0
        $n = $allMbx.Count
        foreach ($m in $allMbx) {
            $i++
            Write-Progress -Id 1 -ParentId 0 -Activity 'Reading mailbox statistics' -Status "$i of $n" -PercentComplete ([int](100 * $i / [math]::Max($n, 1)))
            $type = [string]$m.RecipientTypeDetails
            $id = [string](Get-Prop $m 'ExternalDirectoryObjectId')
            $smtp = [string]$m.PrimarySmtpAddress
            if ($smtp -match '@(.+)$') { $dn = $Matches[1].ToLower(); if ($smtpDomainCount.ContainsKey($dn)) { $smtpDomainCount[$dn]++ } else { $smtpDomainCount[$dn] = 1 } }
            $st = Get-MbxStat $id
            if (-not $st.Ok) { $statFail++ }
            $hasArch = ([string](Get-Prop $m 'ArchiveStatus') -eq 'Active')
            $ast = $null
            if ($hasArch -and ($type -eq 'UserMailbox' -or $type -eq 'SharedMailbox')) { $ast = Get-MbxStat $id -Archive }
            $archMB = 0.0; if ($ast) { $archMB = ConvertTo-MB $ast.Bytes }
            $mbMB = ConvertTo-MB $st.Bytes

            if ($type -eq 'RoomMailbox' -or $type -eq 'EquipmentMailbox') {
                $lbl = 'Room mailbox'; if ($type -eq 'EquipmentMailbox') { $lbl = 'Equipment mailbox' }
                $infoRows.Add([pscustomobject][ordered]@{ Category = $lbl; Name = [string]$m.DisplayName; Identifier = $smtp; Detail = "Items: $($st.Items)"; SizeMB = $mbMB })
                continue
            }

            $u = $null
            if ($id -and $usersById.ContainsKey($id.ToLower())) { $u = $usersById[$id.ToLower()] }
            elseif ($m.UserPrincipalName -and $usersByUpn.ContainsKey(([string]$m.UserPrincipalName).ToLower())) { $u = $usersByUpn[([string]$m.UserPrincipalName).ToLower()] }
            $enabled = 'Unknown'; $licensed = 'Unknown'; $licText = ''; $entraSign = $null
            if ($u) {
                $enabled = ConvertTo-YesNo ([bool]$u.accountEnabled)
                $licensed = ConvertTo-YesNo (@($u.assignedLicenses).Count -gt 0)
                $licText = Get-LicenseText $u
                $entraSign = Get-UserSignIn $u
            } else {
                $enabled = ConvertTo-YesNo (-not [bool](Get-Prop $m 'AccountDisabled'))
            }
            $lastAct = Get-LatestDate $st.LastLogon $entraSign
            $total = [math]::Round($mbMB + $archMB, 1)

            $flags = New-Object System.Collections.Generic.List[string]
            if ($total -lt $SmallObjectThresholdMB) { $flags.Add("Empty/small (under $SmallObjectThresholdMB MB)") }
            if ($type -eq 'UserMailbox') {
                if ($enabled -eq 'No') { $flags.Add('Disabled / sign-in blocked') }
                if ($licensed -eq 'No') { $flags.Add('Unlicensed user mailbox') }
                if (-not $lastAct) { $flags.Add('No sign-in or logon recorded') }
                elseif ($lastAct -lt $staleCut) { $flags.Add("No sign-in/logon in $StaleDays+ days") }
            }
            $row = New-Row $cMbx @{
                Name = [string]$m.DisplayName; UserPrincipalName = [string]$m.UserPrincipalName; Domain = (Get-UpnDomain $m.UserPrincipalName); PrimarySmtp = $smtp
                AccountEnabled = $enabled; Licensed = $licensed; Licenses = $licText; MailboxSizeMB = $mbMB; ItemCount = $st.Items
                Archive = $(if ($hasArch) { 'Yes' } else { 'No' }); ArchiveSizeMB = $archMB; TotalSizeMB = $total
                LitigationHold = ConvertTo-YesNo ([bool](Get-Prop $m 'LitigationHoldEnabled'))
                LastLogonExchange = Format-DateStr $st.LastLogon; LastSignInEntra = Format-DateStr $entraSign
                Created = Format-DateStr (ConvertTo-DateTimeSafe (Get-Prop $m 'WhenCreated'))
                Flags = ($flags -join '; '); Action = (Get-DefaultAction $flags)
            }
            if ($type -eq 'SharedMailbox') { $sharedRows.Add($row) } else { $userRows.Add($row) }
        }
        Write-Progress -Id 1 -ParentId 0 -Activity 'Reading mailbox statistics' -Completed
        if ($statFail -gt 0) { Add-Warn 'Exchange' "$statFail mailbox(es) returned no statistics (size shows 0). First error: $($script:FirstStatError) | Fallback error: $($script:FirstStatError2)" }

        # Informational recipients
        Start-Stage 6 'Reading other Exchange recipients (information only)'
        try {
            foreach ($x in @(Invoke-WithRetry { Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop })) {
                $infoRows.Add([pscustomobject][ordered]@{ Category = 'Inactive mailbox'; Name = [string]$x.DisplayName; Identifier = [string]$x.PrimarySmtpAddress; Detail = 'Retained mailbox of a deleted user'; SizeMB = '' })
            }
        } catch { Add-Warn 'Exchange' "Inactive mailboxes could not be read: $($_.Exception.Message)" }
        try {
            foreach ($x in @(Invoke-WithRetry { Get-MailContact -ResultSize Unlimited -ErrorAction Stop })) {
                $infoRows.Add([pscustomobject][ordered]@{ Category = 'Mail contact'; Name = [string]$x.DisplayName; Identifier = [string]$x.PrimarySmtpAddress; Detail = ''; SizeMB = '' })
            }
        } catch { Add-Warn 'Exchange' "Mail contacts could not be read: $($_.Exception.Message)" }
        try {
            foreach ($x in @(Invoke-WithRetry { Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop })) {
                $cat = 'Distribution list'
                if ([string]$x.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') { $cat = 'Mail-enabled security group' }
                elseif ([string]$x.RecipientTypeDetails -eq 'RoomList') { $cat = 'Room list' }
                $infoRows.Add([pscustomobject][ordered]@{ Category = $cat; Name = [string]$x.DisplayName; Identifier = [string]$x.PrimarySmtpAddress; Detail = [string]$x.RecipientTypeDetails; SizeMB = '' })
            }
        } catch { Add-Warn 'Exchange' "Distribution groups could not be read: $($_.Exception.Message)" }
        try {
            foreach ($x in @(Invoke-WithRetry { Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop })) {
                $infoRows.Add([pscustomobject][ordered]@{ Category = 'Dynamic distribution list'; Name = [string]$x.DisplayName; Identifier = [string]$x.PrimarySmtpAddress; Detail = ''; SizeMB = '' })
            }
        } catch { Add-Warn 'Exchange' "Dynamic distribution groups could not be read: $($_.Exception.Message)" }
        try {
            $pf = @(Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction Stop | Where-Object { [string]$_.Identity -ne '\' })
            if ($pf.Count -gt 0) {
                $infoRows.Add([pscustomobject][ordered]@{ Category = 'Public folder'; Name = "$($pf.Count) public folder(s)"; Identifier = ''; Detail = 'Public folders are migrated separately; count only'; SizeMB = '' })
            }
        } catch { }

        # ---- Microsoft 365 Groups / Teams (Exchange view) -----------------------------------
        Start-Stage 7 'Reading Microsoft 365 Groups and Teams'
        Write-Note 'Listing groups (a few minutes on larger tenants - the window may look idle)...'
        $ugroups = @(Invoke-WithRetry { Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop })
        Write-Note "$($ugroups.Count) Microsoft 365 group(s) found."
        $groupById = @{}
        $urlToGid = @{}
        foreach ($g in $ugroups) {
            $gid = ([string](Get-Prop $g 'ExternalDirectoryObjectId')).ToLower()
            if (-not $gid) { continue }
            $groupById[$gid] = $g
            $su = [string](Get-Prop $g 'SharePointSiteUrl')
            if ($su) { $urlToGid[$su.TrimEnd('/').ToLower()] = $gid }
        }

        $graphTeams = @{}
        if ($script:GraphOk) {
            try {
                $tl = @(Get-GraphPaged 'https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x%20eq%20''Team'')&$select=id,displayName&$top=100')
                foreach ($t in $tl) { $graphTeams[([string]$t.id).ToLower()] = @{ Channels = ''; Private = ''; Shared = ''; Archived = '' } }
            } catch { Add-Warn 'Teams' "Could not list Teams from Graph (Teams detected from Exchange/SharePoint hints instead): $($_.Exception.Message)" }
            if (-not $SkipChannelDetail -and $graphTeams.Count -gt 0) {
                $ti = 0; $tFail = 0
                foreach ($tid in @($graphTeams.Keys)) {
                    $ti++
                    Write-Progress -Id 1 -ParentId 0 -Activity 'Reading Team details' -Status "$ti of $($graphTeams.Count)" -PercentComplete ([int](100 * $ti / $graphTeams.Count))
                    try {
                        $ch = @(Get-GraphPaged ('https://graph.microsoft.com/v1.0/teams/{0}/channels?$select=id,membershipType' -f $tid))
                        $graphTeams[$tid].Channels = $ch.Count
                        $graphTeams[$tid].Private = @($ch | Where-Object { [string]$_.membershipType -eq 'private' }).Count
                        $graphTeams[$tid].Shared = @($ch | Where-Object { [string]$_.membershipType -eq 'shared' }).Count
                        $tm = Invoke-MgGraphRequest -Method GET -Uri ('https://graph.microsoft.com/v1.0/teams/{0}?$select=id,isArchived' -f $tid) -OutputType PSObject
                        $graphTeams[$tid].Archived = ConvertTo-YesNo ([bool](Get-Prop $tm 'isArchived'))
                    } catch { $tFail++ }
                }
                Write-Progress -Id 1 -ParentId 0 -Activity 'Reading Team details' -Completed
                if ($tFail -gt 0) { Add-Warn 'Teams' "Channel/archive detail was unavailable for $tFail Team(s)." }
            }
        }

        # ---- SharePoint and OneDrive --------------------------------------------------------
        $spoAll = @(); $spoPersonal = @()
        $spoOk = $false
        if ($haveSPO) {
            Write-Step 'Sign in 3 of 3: SharePoint Online'
            $adminUrl = $SharePointAdminUrl
            if (-not $adminUrl -and $initialDomain) { $adminUrl = 'https://{0}-admin.sharepoint.com' -f $initialDomain.Split('.')[0] }
            if (-not $adminUrl) { $adminUrl = Read-Host 'SharePoint admin URL (for example https://contoso-admin.sharepoint.com)' }
            try {
                Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                Connect-SPOService -Url $adminUrl -ErrorAction Stop
                $spoOk = $true
            } catch {
                Add-Warn 'SharePoint' "Could not connect to $adminUrl : $($_.Exception.Message). If the tenant was renamed, run again with -SharePointAdminUrl. OneDrive and site counts are MISSING."
            }
        }
        if ($spoOk) {
            Start-Stage 8 'Reading SharePoint sites and OneDrive accounts (can take a few minutes)'
            try { $spoAll = @(Invoke-WithRetry { Get-SPOSite -Limit All -Detailed -ErrorAction Stop }) }
            catch { Add-Warn 'SharePoint' "Sites could not be read: $($_.Exception.Message)" }
            try { $spoPersonal = @(Invoke-WithRetry { Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop }) }
            catch { Add-Warn 'OneDrive' "OneDrive accounts could not be read: $($_.Exception.Message)" }
        }

        # ---- OneDrive rows ------------------------------------------------------------------
        Start-Stage 9 'Analysing OneDrive, sites, Teams and groups'
        $cOD = @('Name', 'OwnerUPN', 'Domain', 'Url', 'SizeMB', 'SizeGB', 'LastContentModified', 'LockState', 'OwnerEnabled', 'OwnerLicensed', 'OwnerLicenses', 'OwnerLastSignIn', 'Flags') + $cTrk
        $odRows = New-Object System.Collections.Generic.List[object]
        $odByOwner = @{}
        $odErr = 0
        foreach ($s in $spoPersonal) {
            try {
            $owner = [string](Get-Prop $s 'Owner')
            $u = $null; if ($owner -and $usersByUpn.ContainsKey($owner.ToLower())) { $u = $usersByUpn[$owner.ToLower()] }
            $sizeMB = [math]::Round([double](Nz (Get-Prop $s 'StorageUsageCurrent') 0), 1)
            $lm = ConvertTo-DateTimeSafe (Get-Prop $s 'LastContentModifiedDate')
            $enabled = 'Unknown'; $licensed = 'Unknown'; $licText = ''; $sign = $null
            $flags = New-Object System.Collections.Generic.List[string]
            if ($u) {
                $enabled = ConvertTo-YesNo ([bool]$u.accountEnabled); $licensed = ConvertTo-YesNo (@($u.assignedLicenses).Count -gt 0); $licText = Get-LicenseText $u; $sign = Get-UserSignIn $u
                if ($enabled -eq 'No') { $flags.Add('Owner account disabled') }
                if ($licensed -eq 'No') { $flags.Add('Owner unlicensed') }
            } elseif ($script:GraphOk -and $owner) { $flags.Add('Owner not found in Entra (deleted user?)') }
            if ($sizeMB -lt $SmallObjectThresholdMB) { $flags.Add("Empty/small (under $SmallObjectThresholdMB MB)") }
            $ls = [string](Get-Prop $s 'LockState'); if ($ls -and $ls -ne 'Unlock') { $flags.Add("Site locked ($ls)") }
            $nm = [string](Nz $(if ($u) { $u.displayName } else { '' }) (Nz (Get-Prop $s 'Title') $owner))
            $odRows.Add((New-Row $cOD @{
                Name = $nm; OwnerUPN = $owner; Domain = (Get-UpnDomain $owner); Url = [string]$s.Url; SizeMB = $sizeMB; LastContentModified = Format-DateStr $lm
                LockState = $ls; OwnerEnabled = $enabled; OwnerLicensed = $licensed; OwnerLicenses = $licText; OwnerLastSignIn = Format-DateStr $sign
                Flags = ($flags -join '; '); Action = (Get-DefaultAction $flags) }))
            if ($owner) { $odByOwner[$owner.ToLower()] = $sizeMB }
            } catch {
                $odErr++
                if ($odErr -le 3) { Add-Warn 'OneDrive' "Skipped one account ($($s.Url)) because of an error: $($_.Exception.Message) (script line $($_.InvocationInfo.ScriptLineNumber); $($_.Exception.GetType().Name))" }
            }
        }

        # ---- SharePoint site classification -------------------------------------------------
        $systemTemplates = @('SPSMSITEHOST#0', 'APPCATALOG#0', 'SRCHCEN#0', 'SRCHCENTERLITE#0', 'POINTPUBLISHINGHUB#0', 'POINTPUBLISHINGPERSONAL#0', 'POINTPUBLISHINGTOPIC#0', 'TENANTADMIN#0', 'EDISC#0', 'REDIRECTSITE#0', 'SPSPERS#0', 'SPSPERS#2', 'SPSPERS#3', 'SPSPERS#4', 'SPSPERS#5', 'SPSPERS#6', 'SPSPERS#7', 'SPSPERS#8', 'SPSPERS#9', 'SPSPERS#10', 'SPSPERS#11')
        $cSite = @('Name', 'Url', 'Template', 'SiteType', 'SizeMB', 'SizeGB', 'Owner', 'LastContentModified', 'LockState', 'HubSite', 'Flags') + $cTrk
        $siteRows = New-Object System.Collections.Generic.List[object]
        $groupSites = @{}
        $teamHints = @{}
        $reconRows = New-Object System.Collections.Generic.List[object]
        $sysCount = 0
        $rootUrl = ''
        if ($initialDomain) { $rootUrl = ('https://{0}.sharepoint.com' -f $initialDomain.Split('.')[0]).ToLower() }
        $siteErr = 0
        foreach ($s in $spoAll) {
            try {
            $url = ([string]$s.Url).TrimEnd('/')
            $tpl = [string]$s.Template
            $sizeMB = [math]::Round([double](Nz (Get-Prop $s 'StorageUsageCurrent') 0), 1)
            $gid = [string](Get-Prop $s 'GroupId'); $rgid = [string](Get-Prop $s 'RelatedGroupId')
            $isChan = ($tpl -like 'TEAMCHANNEL*') -or ((Get-Prop $s 'IsTeamsChannelConnected') -eq $true)
            $isTeamSite = ((Get-Prop $s 'IsTeamsConnected') -eq $true)
            $parent = ''
            if ($isChan) { if (Test-RealGuid $rgid) { $parent = $rgid } elseif (Test-RealGuid $gid) { $parent = $gid } }
            elseif (Test-RealGuid $gid) { $parent = $gid }
            elseif ($urlToGid.ContainsKey($url.ToLower())) { $parent = $urlToGid[$url.ToLower()] }
            $parent = $parent.ToLower()

            $isSystem = ($systemTemplates -contains $tpl.ToUpper()) -or ($url -match '/sites/contentTypeHub$') -or ($url -match '/portals/hub$') -or ($url -match '/search$' -and $tpl -like 'SRCH*') -or ($url -match '-my\.sharepoint\.com$')
            if ($isSystem) {
                $sysCount++
                $infoRows.Add([pscustomobject][ordered]@{ Category = 'System / hidden site'; Name = [string](Nz $s.Title $url); Identifier = $url; Detail = $tpl; SizeMB = $sizeMB })
                continue
            }
            if ($parent -and $groupById.ContainsKey($parent)) {
                if (-not $groupSites.ContainsKey($parent)) { $groupSites[$parent] = New-Object System.Collections.Generic.List[object] }
                $groupSites[$parent].Add([pscustomobject]@{ Url = $url; SizeMB = $sizeMB; IsChannel = $isChan; Title = [string]$s.Title })
                if ($isTeamSite -or $isChan) { $teamHints[$parent] = $true }
                $kind = 'Group / Team site'; if ($isChan) { $kind = 'Channel site (private/shared)' }
                $reconRows.Add([pscustomobject][ordered]@{ Object = $kind; Parent = [string]$groupById[$parent].DisplayName; ParentType = 'Group/Team'; Identifier = $url; SizeMB = $sizeMB; CountedAs = 'Included in parent Team/Group (not counted separately)' })
                continue
            }
            # Standalone (or orphaned group/channel site)
            $flags = New-Object System.Collections.Generic.List[string]
            if ($parent -or $isChan) { $flags.Add('Group/Team-connected site but the group was not found (deleted or unreadable)') }
            if ($sizeMB -lt $SmallObjectThresholdMB) { $flags.Add("Empty/small (under $SmallObjectThresholdMB MB)") }
            $ls = [string](Get-Prop $s 'LockState'); if ($ls -and $ls -ne 'Unlock') { $flags.Add("Site locked ($ls)") }
            $lm = ConvertTo-DateTimeSafe (Get-Prop $s 'LastContentModifiedDate')
            if ($lm -and $lm -lt $staleCut) { $flags.Add("No content change in $StaleDays+ days") }
            if ($rootUrl -and $url.ToLower() -eq $rootUrl) { $flags.Add('Tenant root site') }
            $stype = switch -Wildcard ($tpl) {
                'SITEPAGEPUBLISHING*' { 'Communication site' }
                'STS#3' { 'Team site (no Microsoft 365 group)' }
                'STS#0' { 'Classic team site' }
                'GROUP*' { 'Group site' }
                default { 'Other' }
            }
            $siteRows.Add((New-Row $cSite @{
                Name = [string](Nz $s.Title $url); Url = $url; Template = $tpl; SiteType = $stype; SizeMB = $sizeMB; Owner = [string](Get-Prop $s 'Owner')
                LastContentModified = Format-DateStr $lm; LockState = $ls; HubSite = ConvertTo-YesNo ([bool](Get-Prop $s 'IsHubSite'))
                Flags = ($flags -join '; '); Action = (Get-DefaultAction $flags) }))
            } catch {
                $siteErr++
                if ($siteErr -le 3) { Add-Warn 'SharePoint' "Skipped one site ($($s.Url)) because of an error: $($_.Exception.Message) (script line $($_.InvocationInfo.ScriptLineNumber); $($_.Exception.GetType().Name))" }
            }
        }
        if ($spoOk) {
            try {
                foreach ($ds in @(Get-SPODeletedSite -Limit All -ErrorAction Stop)) {
                    $infoRows.Add([pscustomobject][ordered]@{ Category = 'Deleted site (recycle bin)'; Name = [string]$ds.Url; Identifier = [string]$ds.Url; Detail = "Deleted: $(Format-DateStr (ConvertTo-DateTimeSafe (Get-Prop $ds 'DeletionTime')))"; SizeMB = '' })
                }
            } catch { }
        }

        # ---- Teams and Groups rows ----------------------------------------------------------
        $cTeam = @('Name', 'Email', 'Privacy', 'Owners', 'Members', 'Channels', 'PrivateChannels', 'SharedChannels', 'Archived', 'Created', 'SiteUrl', 'SiteSizeMB', 'SiteSizeGB', 'MailboxSizeMB', 'MailboxSizeGB', 'TotalSizeMB', 'TotalSizeGB', 'GroupId', 'Flags') + $cTrk
        $cGrp = @('Name', 'Email', 'Privacy', 'Owners', 'Members', 'Created', 'SiteUrl', 'SiteSizeMB', 'SiteSizeGB', 'MailboxSizeMB', 'MailboxSizeGB', 'TotalSizeMB', 'TotalSizeGB', 'GroupId', 'Flags') + $cTrk
        $teamRows = New-Object System.Collections.Generic.List[object]
        $grpRows = New-Object System.Collections.Generic.List[object]
        $gmFail = 0
        $gmStreak = 0
        $gi = 0
        $grpErr = 0
        foreach ($g in $ugroups) {
            try {
            $gi++
            Write-Progress -Id 1 -ParentId 0 -Activity 'Reading group mailboxes' -Status "$gi of $($ugroups.Count)" -PercentComplete ([int](100 * $gi / [math]::Max($ugroups.Count, 1)))
            $gid = ([string](Get-Prop $g 'ExternalDirectoryObjectId')).ToLower()
            $rpo = @(Get-Prop $g 'ResourceProvisioningOptions')
            $isTeam = ($graphTeams.ContainsKey($gid)) -or ($rpo -contains 'Team') -or ($teamHints.ContainsKey($gid))
            $sites = New-Object System.Collections.ArrayList
            if ($groupSites.ContainsKey($gid)) { foreach ($gs in $groupSites[$gid]) { [void]$sites.Add($gs) } }
            $siteSum = 0.0
            foreach ($stt in $sites) { $siteSum += [double]$stt.SizeMB }
            $siteMB = [math]::Round([double]$siteSum, 1)
            $mainUrl = [string](Get-Prop $g 'SharePointSiteUrl'); if (-not $mainUrl -and $sites.Count -gt 0) { $mainUrl = ($sites | Where-Object { -not $_.IsChannel } | Select-Object -First 1).Url }

            $gms = [pscustomobject]@{ Bytes = [double]0; Items = 0; LastLogon = $null; Ok = $false }
            if (-not $SkipGroupMailboxSize -and $gmStreak -lt 3) {
                $gms = Get-MbxStat $gid -Fast
                if ($gms.Ok) { $gmStreak = 0 } else { $gmFail++; $gmStreak++ }
            } else { $gmFail++ }
            $gmMB = ConvertTo-MB $gms.Bytes
            $total = [math]::Round($siteMB + $gmMB, 1)
            $owners = @(@((Get-Prop $g 'ManagedBy')) | Where-Object { $_ } | ForEach-Object { $ov = [string]$_; if ($usersById.ContainsKey($ov.ToLower())) { [string]$usersById[$ov.ToLower()].displayName } else { $ov } })
            $flags = New-Object System.Collections.Generic.List[string]
            if ($g.PSObject.Properties['ManagedBy'] -and $owners.Count -eq 0) { $flags.Add('No owners') }
            if ($total -lt $SmallObjectThresholdMB) { $flags.Add("Empty/small (under $SmallObjectThresholdMB MB)") }
            if (-not $mainUrl) { $flags.Add('No SharePoint site found') }
            $created = Format-DateStr (ConvertTo-DateTimeSafe (Get-Prop $g 'WhenCreated'))
            $ownerTxt = (($owners | Select-Object -First 3) -join ', '); if ($owners.Count -gt 3) { $ownerTxt += " (+$($owners.Count - 3))" }
            $common = @{
                Name = [string]$g.DisplayName; Email = [string]$g.PrimarySmtpAddress; Privacy = [string](Get-Prop $g 'AccessType')
                Owners = $ownerTxt; Members = (Nz (Get-Prop $g 'GroupMemberCount') ''); Created = $created; SiteUrl = $mainUrl
                SiteSizeMB = $siteMB; MailboxSizeMB = $gmMB; TotalSizeMB = $total; GroupId = $gid
            }
            if ($isTeam) {
                $ti = @{ Channels = ''; Private = ''; Shared = ''; Archived = '' }
                if ($graphTeams.ContainsKey($gid)) { $ti = $graphTeams[$gid] }
                if ($ti.Archived -eq 'Yes') { $flags.Add('Archived Team') }
                $vals = $common.Clone()
                $vals['Channels'] = $ti.Channels; $vals['PrivateChannels'] = $ti.Private; $vals['SharedChannels'] = $ti.Shared; $vals['Archived'] = $ti.Archived
                $vals['Flags'] = ($flags -join '; '); $vals['Action'] = (Get-DefaultAction $flags)
                $teamRows.Add((New-Row $cTeam $vals))
            } else {
                $vals = $common.Clone()
                $vals['Flags'] = ($flags -join '; '); $vals['Action'] = (Get-DefaultAction $flags)
                $grpRows.Add((New-Row $cGrp $vals))
            }
            $reconRows.Add([pscustomobject][ordered]@{ Object = 'Group mailbox'; Parent = [string]$g.DisplayName; ParentType = $(if ($isTeam) { 'Team' } else { 'Group' }); Identifier = [string]$g.PrimarySmtpAddress; SizeMB = $gmMB; CountedAs = 'Included in parent Team/Group (not counted separately)' })
            } catch {
                $grpErr++
                if ($grpErr -le 3) { Add-Warn 'Groups' "Skipped one group ($($g.DisplayName)) because of an error: $($_.Exception.Message) (script line $($_.InvocationInfo.ScriptLineNumber); $($_.Exception.GetType().Name))" }
            }
        }
        Write-Progress -Id 1 -ParentId 0 -Activity 'Reading group mailboxes' -Completed
        if ($odErr + $siteErr + $grpErr -gt 0) { Add-Warn 'General' "Objects skipped because of errors: OneDrive $odErr, sites $siteErr, groups $grpErr. Only the first 3 of each are listed." }
        if ($gmFail -gt 0) { Add-Warn 'Groups' "Group mailbox size was not read for $gmFail group(s) (skipped with -SkipGroupMailboxSize, or after 3 failures in a row to save time); their total shows the SharePoint site size only. Error seen: $($script:GroupStatError)" }
        if ($NoGraph -or -not $script:GraphOk) { Add-Warn 'Teams' 'Teams were identified from Exchange/SharePoint hints (no Graph). A Team with no Teams-connected site marker may appear under M365 Groups - review the two tabs.' 'Info' }

        # ---- Users reference + domains ------------------------------------------------------
        $cUsr = @('Name', 'UserPrincipalName', 'Domain', 'Enabled', 'Licensed', 'Licenses', 'SyncedFromOnPrem', 'Created', 'LastSignInEntra', 'MailboxType', 'MailboxSizeMB', 'MailboxSizeGB', 'HasOneDrive', 'OneDriveSizeMB', 'OneDriveSizeGB')
        $usrRows = New-Object System.Collections.Generic.List[object]
        $mbxTypeById = @{}
        foreach ($m in $allMbx) { $mbxTypeById[([string](Get-Prop $m 'ExternalDirectoryObjectId')).ToLower()] = [string]$m.RecipientTypeDetails }
        $upnDomainCount = @{}
        foreach ($u in $memberUsers) {
            $upn = [string]$u.userPrincipalName
            if ($upn -match '@(.+)$') { $dn = $Matches[1].ToLower(); if ($upnDomainCount.ContainsKey($dn)) { $upnDomainCount[$dn]++ } else { $upnDomainCount[$dn] = 1 } }
            $uid = ([string]$u.id).ToLower()
            $mt = ''; if ($mbxTypeById.ContainsKey($uid)) { $mt = $mbxTypeById[$uid] }
            $mrow = $null
            if ($mt -eq 'UserMailbox') { $mrow = @($userRows | Where-Object { $_.UserPrincipalName -eq $upn }) | Select-Object -First 1 }
            $hasOd = 'No'; $odMB = ''
            if ($odByOwner.ContainsKey($upn.ToLower())) { $hasOd = 'Yes'; $odMB = $odByOwner[$upn.ToLower()] }
            $usrRows.Add((New-Row $cUsr @{
                Name = [string]$u.displayName; UserPrincipalName = $upn; Domain = (Get-UpnDomain $upn); Enabled = ConvertTo-YesNo ([bool]$u.accountEnabled)
                Licensed = ConvertTo-YesNo (@($u.assignedLicenses).Count -gt 0); Licenses = (Get-LicenseText $u); SyncedFromOnPrem = ConvertTo-YesNo ([bool](Get-Prop $u 'onPremisesSyncEnabled'))
                Created = Format-DateStr (ConvertTo-DateTimeSafe $u.createdDateTime); LastSignInEntra = Format-DateStr (Get-UserSignIn $u)
                MailboxType = $mt; MailboxSizeMB = $(if ($mrow) { $mrow.TotalSizeMB } else { '' }); HasOneDrive = $hasOd; OneDriveSizeMB = $odMB }))
        }
        if (-not $script:GraphOk) { foreach ($r in $userRows) { $dn = ''; if ($r.UserPrincipalName -match '@(.+)$') { $dn = $Matches[1].ToLower() }; if ($dn) { if ($upnDomainCount.ContainsKey($dn)) { $upnDomainCount[$dn]++ } else { $upnDomainCount[$dn] = 1 } } } }

        $cDom = @('Domain', 'Default', 'Initial', 'Verified', 'Authentication', 'Services', 'ExchangeAccepted', 'ExchangeDomainType', 'UsersWithThisUPN', 'MailboxesWithThisPrimarySMTP')
        $domMap = @{}
        foreach ($d in $graphDomains) { $domMap[([string]$d.id).ToLower()] = $d }
        foreach ($a in $accepted) { if (-not $domMap.ContainsKey(([string]$a.DomainName).ToLower())) { $domMap[([string]$a.DomainName).ToLower()] = $null } }
        $domRows = New-Object System.Collections.Generic.List[object]
        foreach ($k in ($domMap.Keys | Sort-Object)) {
            $gd = $domMap[$k]
            $ad = $accepted | Where-Object { ([string]$_.DomainName).ToLower() -eq $k } | Select-Object -First 1
            $isDef = if ($gd) { [bool]$gd.isDefault } elseif ($ad) { [bool]$ad.Default } else { $false }
            $cu = 0; if ($upnDomainCount.ContainsKey($k)) { $cu = $upnDomainCount[$k] }
            $cm = 0; if ($smtpDomainCount.ContainsKey($k)) { $cm = $smtpDomainCount[$k] }
            $domRows.Add((New-Row $cDom @{
                Domain = $k; Default = ConvertTo-YesNo $isDef; Initial = $(if ($gd) { ConvertTo-YesNo ([bool]$gd.isInitial) } elseif ($k -eq $initialDomain.ToLower()) { 'Yes' } else { 'No' })
                Verified = $(if ($gd) { ConvertTo-YesNo ([bool]$gd.isVerified) } else { 'Unknown' })
                Authentication = $(if ($gd) { [string]$gd.authenticationType } else { '' })
                Services = $(if ($gd) { (@($gd.supportedServices) -join ', ') } else { '' })
                ExchangeAccepted = $(if ($ad) { 'Yes' } else { 'No' }); ExchangeDomainType = $(if ($ad) { [string]$ad.DomainType } else { '' })
                UsersWithThisUPN = $cu; MailboxesWithThisPrimarySMTP = $cm }))
        }

        # ---- Group the lists by UPN domain (the sign-in name domain, not email aliases) ------
        $userRows = @($userRows | Sort-Object -Property Domain, Name)
        $sharedRows = @($sharedRows | Sort-Object -Property Domain, Name)
        $odRows = @($odRows | Sort-Object -Property Domain, Name)
        $usrRows = @($usrRows | Sort-Object -Property Domain, Name)
        $siteRows = $siteRows.ToArray()
        $teamRows = $teamRows.ToArray()
        $grpRows = $grpRows.ToArray()
        $infoRows = $infoRows.ToArray()
        foreach ($ir in $infoRows) { $mbv = "$($ir.SizeMB)"; $gbv = ''; if ($mbv -ne '' -and $mbv -match '^-?[\d\.]+$') { $gbv = [math]::Round([double]$mbv / 1024, 2) }; Add-Member -InputObject $ir -NotePropertyName SizeGB -NotePropertyValue $gbv -Force }
        $reconRows = $reconRows.ToArray()
        foreach ($rr2 in $reconRows) { $mbv = "$($rr2.SizeMB)"; $gbv = ''; if ($mbv -ne '' -and $mbv -match '^-?[\d\.]+$') { $gbv = [math]::Round([double]$mbv / 1024, 2) }; Add-Member -InputObject $rr2 -NotePropertyName SizeGB -NotePropertyValue $gbv -Force }
        $domRows = $domRows.ToArray()
        $licRows = $licRows.ToArray()
        $domSet = @{}
        foreach ($r0 in (@($userRows) + @($sharedRows) + @($odRows))) { $dd0 = [string]$r0.Domain; if ($dd0) { $domSet[$dd0] = $true } }
        $domList = @($domSet.Keys | Sort-Object)
        $byDomain = @()
        foreach ($dd in $domList) {
            $um = @($userRows | Where-Object { $_.Domain -eq $dd }); $sm = @($sharedRows | Where-Object { $_.Domain -eq $dd }); $om = @($odRows | Where-Object { $_.Domain -eq $dd })
            $byDomain += , @($dd, $um.Count, @($um | Where-Object { $_.Action -eq 'Migrate' }).Count, $sm.Count, @($sm | Where-Object { $_.Action -eq 'Migrate' }).Count, $om.Count, @($om | Where-Object { $_.Action -eq 'Migrate' }).Count)
        }
        if ($domList.Count -gt 1) { Write-Note "Multiple UPN domains found: $($domList -join ', ') - tables are grouped by domain." }

        # ---- Summary data -------------------------------------------------------------------
        $cats = @(
            @{ Label = 'User mailboxes'; Sheet = 'User Mailboxes'; Cols = $cMbx; Size = 'TotalSizeMB'; Rows = $userRows },
            @{ Label = 'Shared mailboxes'; Sheet = 'Shared Mailboxes'; Cols = $cMbx; Size = 'TotalSizeMB'; Rows = $sharedRows },
            @{ Label = 'OneDrive accounts'; Sheet = 'OneDrive'; Cols = $cOD; Size = 'SizeMB'; Rows = $odRows },
            @{ Label = 'SharePoint sites (standalone)'; Sheet = 'SharePoint Sites'; Cols = $cSite; Size = 'SizeMB'; Rows = $siteRows },
            @{ Label = 'Teams'; Sheet = 'Teams'; Cols = $cTeam; Size = 'TotalSizeMB'; Rows = $teamRows },
            @{ Label = 'Microsoft 365 Groups (non-Team)'; Sheet = 'M365 Groups'; Cols = $cGrp; Size = 'TotalSizeMB'; Rows = $grpRows }
        )
        function Count-Info { param([string]$Cat) return @($infoRows | Where-Object { $_.Category -eq $Cat }).Count }
        $notCounted = @(
            @('Room mailboxes', (Count-Info 'Room mailbox')),
            @('Equipment mailboxes', (Count-Info 'Equipment mailbox')),
            @('Inactive mailboxes', (Count-Info 'Inactive mailbox')),
            @('Mail contacts', (Count-Info 'Mail contact')),
            @('Distribution lists', (Count-Info 'Distribution list')),
            @('Dynamic distribution lists', (Count-Info 'Dynamic distribution list')),
            @('Mail-enabled security groups', (Count-Info 'Mail-enabled security group')),
            @('Guest accounts (excluded)', $guestCount),
            @('Team/Group sites inside a Team or Group', @($reconRows | Where-Object { $_.Object -eq 'Group / Team site' }).Count),
            @('Channel sites inside a Team', @($reconRows | Where-Object { $_.Object -like 'Channel site*' }).Count),
            @('Group mailboxes inside a Team or Group', @($reconRows | Where-Object { $_.Object -eq 'Group mailbox' }).Count),
            @('System / hidden SharePoint sites', $sysCount),
            @('Deleted sites (recycle bin)', (Count-Info 'Deleted site (recycle bin)'))
        )
        $notes = @(
            'Each Team or Microsoft 365 Group counts as ONE object. Its group mailbox, SharePoint site and channel sites are not counted again (see the Reconciliation tab).',
            'Shared mailboxes are counted whether licensed or not. Whether they are migrated is decided in the Action column.',
            'SharePoint sites counts standalone sites only: OneDrive, Team/Group sites and system sites are excluded.',
            'Guests, room and equipment mailboxes, distribution lists, contacts and public folders are listed for information and are not counted.',
            'Action column: Migrate = in scope, Skip = out of scope, Review = flagged and needs a decision. In the Excel workbook the Summary counts update when Action values change.',
            "Flags: objects under $SmallObjectThresholdMB MB, disabled or unlicensed owners, and users with no sign-in for $StaleDays+ days are set to Review by default.",
            'Sizes are shown in MB (as reported by Exchange Online and SharePoint Online) and in GB (MB divided by 1024). Mailbox Total includes the archive.'
        )
        $runStamp = (Get-Date).ToString('yyyy-MM-dd HH:mm') + ' (local)  /  ' + $now.ToString('yyyy-MM-dd HH:mm') + ' UTC'
        $metaRows = @(
            @('Client', $script:ClientName), @('Tenant', $tenantName), @('Tenant ID', $tenantId),
            @('Initial domain', $initialDomain), @('Default domain', $defaultDomain), @('Run date', $runStamp), @('Run by', $runBy),
            @('Script version', $script:Version),
            @('Microsoft Graph data', $(if ($script:GraphOk) { 'Yes' } else { 'No - Entra details limited' })),
            @('Entra last sign-in', $(if ($script:SignInOk) { 'Available (shown with Exchange last logon)' } else { 'Not available - Exchange last logon used' })),
            @('Small object threshold (MB)', $SmallObjectThresholdMB), @('Stale days', $StaleDays),
            @('Warnings', "$(@($script:Warnings | Where-Object { $_.Level -eq 'Warning' }).Count) (see Warnings tab)")
        )
        if ($script:Warnings.Count -eq 0) { Add-Warn 'General' 'No warnings.' 'Info' }
        $cWarn = @('Area', 'Level', 'Message')
        $cInfo = @('Category', 'Name', 'Identifier', 'Detail', 'SizeMB', 'SizeGB')
        $cRec = @('Object', 'Parent', 'ParentType', 'Identifier', 'SizeMB', 'SizeGB', 'CountedAs')

        # ---- Output files -------------------------------------------------------------------
        Start-Stage 10 'Building the workbook and HTML'
        $slug = ($initialDomain.Split('.')[0] -replace '[^A-Za-z0-9\-]', '')
        if (-not $slug) { $slug = ($script:ClientName -replace '[^A-Za-z0-9\-]', '') }
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmm')
        $baseName = "M365-Inventory_$($slug)_$stamp"
        $runDir = Join-Path $script:OutDir $baseName
        New-Item -ItemType Directory -Path $runDir -Force | Out-Null
        $xlsx = Join-Path $runDir "$baseName.xlsx"
        $html = Join-Path $runDir "$baseName.html"
        $zip = Join-Path $runDir "$baseName.zip"
        $csvDir = Join-Path $runDir "$baseName-csv"

        $tabDefs = @(
            @{ Name = 'User Mailboxes'; Cols = $cMbx; Rows = $userRows }, @{ Name = 'Shared Mailboxes'; Cols = $cMbx; Rows = $sharedRows },
            @{ Name = 'OneDrive'; Cols = $cOD; Rows = $odRows }, @{ Name = 'SharePoint Sites'; Cols = $cSite; Rows = $siteRows },
            @{ Name = 'Teams'; Cols = $cTeam; Rows = $teamRows }, @{ Name = 'M365 Groups'; Cols = $cGrp; Rows = $grpRows },
            @{ Name = 'Domains'; Cols = $cDom; Rows = $domRows }, @{ Name = 'Users'; Cols = $cUsr; Rows = $usrRows }, @{ Name = 'Licenses'; Cols = $cLic; Rows = $licRows },
            @{ Name = 'Informational'; Cols = $cInfo; Rows = $infoRows }, @{ Name = 'Reconciliation'; Cols = $cRec; Rows = $reconRows },
            @{ Name = 'Warnings'; Cols = $cWarn; Rows = $script:Warnings.ToArray() }
        )

        # HTML first (independent of Excel)
        $summaryHtml = @()
        foreach ($c in $cats) {
            $rws = New-Object System.Collections.ArrayList
            foreach ($rr in $c.Rows) { [void]$rws.Add($rr) }
            $mig = @($rws | Where-Object { $_.Action -eq 'Migrate' })
            $rev = @($rws | Where-Object { $_.Action -eq 'Review' })
            $szSum = 0.0
            foreach ($mr2 in $mig) { $szSum += [double](Nz $mr2.($c.Size) 0) }
            $szMig = [math]::Round([double]($szSum / 1024), 2)
            $summaryHtml += , @($c.Label, $rws.Count, $mig.Count, 0, $rev.Count, $szMig)
        }
        $model = [ordered]@{
            client = $script:ClientName; company = $script:Company; web = $script:Web; version = $script:Version
            tool = $script:ToolName; generated = $runStamp
            meta = @($metaRows | ForEach-Object { , @([string]$_[0], [string]$_[1]) })
            summaryCols = @('Category', 'Found', 'Migrate (as run)', 'Skip (as run)', 'Review (as run)', 'Size to migrate (GB)')
            summary = @($summaryHtml)
            notCounted = @($notCounted | ForEach-Object { , @([string]$_[0], $_[1]) })
            byDomainCols = @('UPN domain', 'User mailboxes', 'of which Migrate', 'Shared mailboxes', 'of which Migrate', 'OneDrive accounts', 'of which Migrate')
            byDomain = @($byDomain)
            notes = @($notes)
            tabs = @(foreach ($t in $tabDefs) {
                $rows2 = @(foreach ($r in $t.Rows) { , @(foreach ($c in $t.Cols) { $v = $r.$c; if ($null -eq $v) { '' } else { $v } }) })
                [ordered]@{ name = $t.Name; columns = @($t.Cols); rows = $rows2 }
            })
        }
        Build-Html $model | Set-Content -Path $html -Encoding UTF8

        $filesForZip = @($html)
        if ($haveXlsx) {
            if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
            $pkg = Open-ExcelPackage -Path $xlsx -Create
            $wsS = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Summary'
            foreach ($t in $tabDefs) { Write-Sheet -Pkg $pkg -Name $t.Name -Rows $t.Rows -Columns $t.Cols }
            $wsO = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Sign-off'

            # Summary sheet
            $wsS.Column(1).Width = 40; foreach ($cc in 2..7) { $wsS.Column($cc).Width = 18 }
            $wsS.Cells['A1'].Value = 'Microsoft 365 Tenant Inventory'; $wsS.Cells['A1'].Style.Font.Size = 18; $wsS.Cells['A1'].Style.Font.Bold = $true
            $r = 4
            foreach ($mr in $metaRows) {
                $wsS.Cells["A$r"].Value = [string]$mr[0]; $wsS.Cells["A$r"].Style.Font.Bold = $true
                $wsS.Cells["B$r"].Value = [string]$mr[1]; $r++
            }
            $r++
            $wsS.Cells["A$r"].Value = 'How to use: review each tab, set Action to Migrate / Skip for every Review row, then sign the Sign-off tab. Counts below update automatically.'
            $wsS.Cells["A$r"].Style.Font.Italic = $true
            $r += 2
            $hdr = @('Billable object category', 'Found', 'Migrate', 'Skip', 'Review (pending)', 'Size to migrate (GB)', 'Check')
            for ($k = 0; $k -lt $hdr.Count; $k++) { $cc2 = $k + 1; $wsS.Cells[$r, $cc2].Value = $hdr[$k]; $wsS.Cells[$r, $cc2].Style.Font.Bold = $true }
            $hdrRow = $r; $r++
            $firstCat = $r
            foreach ($c in $cats) {
                $sh = "'" + $c.Sheet + "'"
                $aL = Get-ColLetter $c.Cols 'Action'; $sL = Get-ColLetter $c.Cols $c.Size
                $wsS.Cells["A$r"].Value = $c.Label
                $wsS.Cells["B$r"].Formula = "COUNTA($sh!`$A:`$A)-1"
                $wsS.Cells["C$r"].Formula = "COUNTIF($sh!`$$($aL):`$$($aL),""Migrate"")"
                $wsS.Cells["D$r"].Formula = "COUNTIF($sh!`$$($aL):`$$($aL),""Skip"")"
                $wsS.Cells["E$r"].Formula = "COUNTIF($sh!`$$($aL):`$$($aL),""Review"")"
                $wsS.Cells["F$r"].Formula = "ROUND(SUMIF($sh!`$$($aL):`$$($aL),""Migrate"",$sh!`$$($sL):`$$($sL))/1024,2)"
                $wsS.Cells["G$r"].Formula = "IF(B$r=C$r+D$r+E$r,""OK"",""Check Action column"")"
                $r++
            }
            $lastCat = $r - 1
            $wsS.Cells["A$r"].Value = 'TOTAL OBJECTS'
            foreach ($col in @('B', 'C', 'D', 'E', 'F')) { $wsS.Cells["$col$r"].Formula = "SUM($col$($firstCat):$col$($lastCat))" }
            $wsS.Cells["A$($r):G$($r)"].Style.Font.Bold = $true
            $totalRow = $r
            $r += 2
            if ($domList.Count -gt 0) {
                $wsS.Cells["A$r"].Value = 'Objects by UPN domain (domain of the sign-in name, not email aliases)'; $wsS.Cells["A$r"].Style.Font.Bold = $true; $r++
                $dh = @('UPN domain', 'User mailboxes', 'of which Migrate', 'Shared mailboxes', 'of which Migrate', 'OneDrive accounts', 'of which Migrate')
                for ($k = 0; $k -lt $dh.Count; $k++) { $cc3 = $k + 1; $wsS.Cells[$r, $cc3].Value = $dh[$k]; $wsS.Cells[$r, $cc3].Style.Font.Bold = $true }
                $r++
                $uD = Get-ColLetter $cMbx 'Domain'; $uA = Get-ColLetter $cMbx 'Action'
                $oD = Get-ColLetter $cOD 'Domain'; $oA = Get-ColLetter $cOD 'Action'
                foreach ($dd in $domList) {
                    $wsS.Cells["A$r"].Value = $dd
                    $wsS.Cells["B$r"].Formula = "COUNTIF('User Mailboxes'!`$$($uD):`$$($uD),`$A$r)"
                    $wsS.Cells["C$r"].Formula = "COUNTIFS('User Mailboxes'!`$$($uD):`$$($uD),`$A$r,'User Mailboxes'!`$$($uA):`$$($uA),""Migrate"")"
                    $wsS.Cells["D$r"].Formula = "COUNTIF('Shared Mailboxes'!`$$($uD):`$$($uD),`$A$r)"
                    $wsS.Cells["E$r"].Formula = "COUNTIFS('Shared Mailboxes'!`$$($uD):`$$($uD),`$A$r,'Shared Mailboxes'!`$$($uA):`$$($uA),""Migrate"")"
                    $wsS.Cells["F$r"].Formula = "COUNTIF(OneDrive!`$$($oD):`$$($oD),`$A$r)"
                    $wsS.Cells["G$r"].Formula = "COUNTIFS(OneDrive!`$$($oD):`$$($oD),`$A$r,OneDrive!`$$($oA):`$$($oA),""Migrate"")"
                    $r++
                }
                $r++
            }
            $wsS.Cells["A$r"].Value = 'Not counted (for information)'; $wsS.Cells["A$r"].Style.Font.Bold = $true; $r++
            foreach ($nc in $notCounted) { $wsS.Cells["A$r"].Value = [string]$nc[0]; $wsS.Cells["B$r"].Value = $nc[1]; $r++ }
            $r++
            $wsS.Cells["A$r"].Value = 'How objects are counted'; $wsS.Cells["A$r"].Style.Font.Bold = $true; $r++
            foreach ($nt in $notes) { $wsS.Cells["A$r"].Value = $nt; $r++ }
            $r++
            $wsS.Cells["A$r"].Value = "Generated by $($script:ToolName) v$($script:Version)"
            $wsS.Cells["A$r"].Style.Font.Size = 8
            try { $wsS.Cells["A$r"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray) } catch { }
            try { $wsS.Cells["A$($hdrRow):G$($hdrRow)"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsS.Cells["A$($hdrRow):G$($hdrRow)"].Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(220, 230, 241) } catch { }

            # Sign-off sheet
            $wsO.Column(1).Width = 34; $wsO.Column(2).Width = 60
            $wsO.Cells['A1'].Value = 'Client sign-off'; $wsO.Cells['A1'].Style.Font.Size = 16; $wsO.Cells['A1'].Style.Font.Bold = $true
            $wsO.Cells['A3'].Value = "I confirm that the object counts and the Migrate / Skip selections in this workbook are accurate and complete for the Microsoft 365 tenant named below, as at the run date shown on the Summary tab."
            $wsO.Cells['A5'].Value = 'Client'; $wsO.Cells['B5'].Value = $script:ClientName
            $wsO.Cells['A6'].Value = 'Tenant'; $wsO.Cells['B6'].Value = $tenantName
            $wsO.Cells['A7'].Value = 'Total objects to migrate'; $wsO.Cells['B7'].Formula = "Summary!C$totalRow"
            $wsO.Cells['A8'].Value = 'Objects still pending review'; $wsO.Cells['B8'].Formula = "Summary!E$totalRow"
            $wsO.Cells['A9'].Value = 'Size to migrate (GB)'; $wsO.Cells['B9'].Formula = "Summary!F$totalRow"
            $wsO.Cells['A11'].Value = 'Client representative name'; $wsO.Cells['A12'].Value = 'Title'; $wsO.Cells['A13'].Value = 'Signature'; $wsO.Cells['A14'].Value = 'Date'
            $wsO.Cells['A16'].Value = 'Consultant representative'; $wsO.Cells['A17'].Value = 'Signature'; $wsO.Cells['A18'].Value = 'Date'
            $wsO.Cells['A5:A18'].Style.Font.Bold = $true

            try { $pkg.Workbook.Worksheets.MoveToStart('Summary') } catch { }
            Close-ExcelPackage -ExcelPackage $pkg
            $filesForZip += $xlsx
        } else {
            New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
            foreach ($t in $tabDefs) { $t.Rows | Select-Object -Property $t.Cols | Export-Csv -Path (Join-Path $csvDir ("{0}.csv" -f $t.Name)) -NoTypeInformation -Encoding UTF8 }
            $filesForZip += $csvDir
        }

        if (Test-Path $zip) { Remove-Item $zip -Force }
        Compress-Archive -Path $filesForZip -DestinationPath $zip -Force

        Write-Progress -Id 1 -Activity 'done' -Completed
        Write-Progress -Id 0 -Activity 'INFINIGENCE M365 Tenant Inventory' -Completed
        # ---- Console summary ----------------------------------------------------------------
        Write-Host ''
        Write-Host '  RESULTS' -ForegroundColor Cyan
        foreach ($sr in $summaryHtml) { Write-Host ("    {0,-34} found {1,5}   (migrate {2}, review {3})" -f $sr[0], $sr[1], $sr[2], $sr[4]) }
        Write-Host ''
        $wc = @($script:Warnings | Where-Object { $_.Level -eq 'Warning' }).Count
        if ($wc -gt 0) { Write-Host "  $wc warning(s) - see the Warnings tab." -ForegroundColor Yellow }
        Write-Host "  Results saved to:" -ForegroundColor Green
        Write-Host "    $zip" -ForegroundColor White
        Write-Host ''
        try { Invoke-Item $runDir } catch { }
    }
    catch {
        Write-Host ''
        Write-Host "  ERROR (script line $($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  $($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
        Write-Host "  $($_.ScriptStackTrace)" -ForegroundColor DarkRed
        Write-Host "  DETAIL: $($_.Exception.ToString())" -ForegroundColor DarkRed
        throw
    }
    finally {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
        try { if ($script:GraphOk) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } } catch { }
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }
    }
}

# Run unless the file is being dot-sourced (for testing)
if ($MyInvocation.InvocationName -ne '.') { Invoke-Inventory }
