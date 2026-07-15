# ============================================================================
# Script    : CDP-BillingUsageReport.v07.ps1
# Synopsis  : Pull N-able billing usage and export flat device-per-row CSV + HTML report
# Version   : 07.0
# Author    : Eric Harless, Head Nerd — N-able
# Twitter   : @Backup_Nerd  |  Email: eric.harless@n-able.com
# GitHub    : https://github.com/BackupNerd
# ============================================================================
#
# Description:
#   Calls the N-able PSAUsageByTimeFrame billing API for Direct and Indirect
#   customers. No ContractID or InvoiceID required — just SFDCID + timeframe.
#
#   Output: ONE ROW PER DEVICE with billing quantities as dynamic columns
#   (Server, VirtualServer, Workstation, M365Users, Per-GB, Continuity, EDR, etc.)
#   Layout mirrors the Max Value Plus report style. Optional EFile CSV merge adds
#   month-over-month delta, churn detection, reconciliation, and unit price tabs.
#
# Usage:
#   .\CDP-BillingUsageReport.v07.ps1                                     # picker UI
#   .\CDP-BillingUsageReport.v07.ps1 -creds prod4 -Timeframe 202605      # XML cred by ID
#   .\CDP-BillingUsageReport.v07.ps1 -SFDCID <id> -APIKey <key>          # direct creds
#
# Authentication modes:
#   1. BillingCreds.xml  — DPAPI-encrypted Export-Clixml store (recommended)
#   2. -SFDCID + -APIKey — direct command-line override (see SECURITY NOTE below)
#   3. -creds ''         — EFile-only mode, no API call
#
# SECURITY NOTE:
#   When using -APIKey on the command line, the value is visible in PowerShell
#   history (Get-History) and may appear in process listings. For production
#   automation, use BillingCreds.xml instead. Rotate any key exposed on a
#   command line or shared screen immediately.
#
# Outputs:
#   Reports\<YYYYMM>\<Account>\  ← all output files land here
#     <timestamp>_<account>_Usage_<YYYYMM>.csv      device-per-row flat CSV
#     <timestamp>_<account>_Summary_<YYYYMM>.csv    hierarchy summary CSV
#     <timestamp>_<account>_Summary_<YYYYMM>.html   interactive HTML report
#
# Prerequisites:
#   • PowerShell 5.1+  (Out-GridView, JSON, DPAPI)
#   • .NET 4.5+        (System.Web for HTML encoding, TLS 1.2)
#   • BillingCreds.xml in script folder  OR  -SFDCID + -APIKey at runtime
#   • Network: HTTPS to prod-nableboomi.n-able.com
#
# Version History:
#   v07.0 (2026-06-26): Added -SFDCID/-APIKey/-AccountName direct-credential params
#   v06.x            : EFile merge, MoM deltas, churn detection, reconciliation, unit prices
#   v05.x            : Dynamic service column discovery, cross-reference tab
#   v04.x            : Collapsible HTML report, privacy blur toggle
#   v03.x            : Summary hierarchy (Reseller/ServiceOrg/EndCustomer/Site)
#   v01-02           : Initial flat CSV export, PSAUsageByTimeFrame integration
#
# References:
#   N-able Billing API:    https://developer.n-able.com/n-able-billing-api/reference
#   Script repository:     https://github.com/BackupNerd
#
# Legal / Disclaimer:
#   This script is provided AS-IS without warranty of any kind, express or
#   implied, including but not limited to the warranties of merchantability,
#   fitness for a particular purpose, and non-infringement. In no event shall
#   the author or N-able Technologies be liable for any claim, damages, or
#   other liability arising from the use of this script.
#
#   Use of this script is at your own risk and implies acceptance of these terms.
#   Always test in a non-production environment before production deployment.
#   The author is not responsible for unintended data exposure, billing
#   discrepancies, or any other outcome resulting from script use or modification.
#
#   This script accesses the N-able Billing API using credentials you supply.
#   You are solely responsible for the security of those credentials and for
#   compliance with your organization's data handling and security policies.
#
# ============================================================================

<#
.SYNOPSIS
    Pull N-able billing usage and export to a device-per-row flat CSV + HTML report.

.DESCRIPTION
    Calls PSAUsageByTimeFrame (works for both Direct and Indirect customers).
    No ContractID or InvoiceID required — just SFDCID + timeframe.
    Output: ONE ROW PER DEVICE with billing quantities as separate columns
    (Server, VirtualServer, Workstation, M365Users, SelectedSizeGB, UsedSizeGB,
    CoveContinuity, EDR_Complete, EDR_Vigilance, EDR_Control, and any future products).
    Layout mirrors the Max Value Plus report style.
    Optional EFile CSV merge adds month-over-month delta, churn detection,
    reconciliation block, and unit price tab.

.PARAMETER creds
    Credential set ID from BillingCreds.xml (e.g. prod1-prod16).
    Leave blank to show the account picker GridView.
    Default: '' (show picker)

.PARAMETER Timeframe
    Billing period in YYYYMM format.
    Default: 202605

.PARAMETER EFileFolder
    Path to folder containing EFile_*.csv files.
    Default: script folder.

.PARAMETER NoGridView
    When set, suppresses the device GridView after export.
    Default: $true

.PARAMETER SFDCID
    Salesforce Account ID — bypasses BillingCreds.xml when combined with -APIKey.
    SECURITY: value is visible in shell history. Prefer BillingCreds.xml for automation.

.PARAMETER APIKey
    N-able Billing API x-api-key token — bypasses BillingCreds.xml when combined with -SFDCID.
    SECURITY: value is visible in shell history. Rotate any key exposed on the command line.

.PARAMETER AccountName
    Optional display name used in console output and output filenames when using direct creds.
    Default: SFDCID-<value>

.EXAMPLE
    # Interactive: show account picker from BillingCreds.xml
    .\CDP-UsageFlatCSV.ps1

.EXAMPLE
    # Specific account by ID, specific timeframe
    .\CDP-UsageFlatCSV.ps1 -creds prod4 -Timeframe 202605

.EXAMPLE
    # Direct credentials (no BillingCreds.xml required)
    .\CDP-UsageFlatCSV.ps1 -SFDCID "0015000000uJNaBAAW" -APIKey "your-api-key" -Timeframe 202605

.EXAMPLE
    # Direct credentials with a friendly account name in output files
    .\CDP-UsageFlatCSV.ps1 -SFDCID "0015000000uJNaBAAW" -APIKey "your-api-key" -AccountName "Acme Corp"

.EXAMPLE
    # EFile-only mode (no API call, just process local EFile CSV files)
    .\CDP-UsageFlatCSV.ps1 -creds ''

.NOTES
    Author  : Eric Harless, Head Nerd — N-able
    Version : 07.0 (2026-06-26)
    GitHub  : https://github.com/BackupNerd

    ⚠  API keys passed via -APIKey are visible in PowerShell history.
       Use BillingCreds.xml for scheduled/automated runs.
    ⚠  Always test in a non-production environment first.
    ⚠  This script is provided AS-IS. See the legal disclaimer in the script header.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)] [string]$creds       = '',   # prod1-prod16, or '' to show picker
    [Parameter(Mandatory=$False)] [string]$Timeframe   = "202607",
    [Parameter(Mandatory=$False)] [string]$EFileFolder = '',   ## Folder containing EFile_*.csv files (default: script folder)
    [switch]$NoGridView = $true,
    # ── Direct credential override (bypasses BillingCreds.xml when both are supplied) ──
    [Parameter(Mandatory=$False)] [string]$SFDCID      = '',   # Salesforce Account ID
    [Parameter(Mandatory=$False)] [string]$APIKey      = '',   # x-api-key token
    [Parameter(Mandatory=$False)] [string]$AccountName = ''    # Display name (optional, used in output/filenames)
)

Clear-Host
Set-Location -Path (Split-Path -Path $MyInvocation.MyCommand.Path)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ── Load credentials from BillingCreds.xml (Export-Clixml / DPAPI-encrypted) ─
function Convert-SecureToPlainText {
    param([securestring]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.NetworkCredential]::new('', $Value).Password
}

function Read-RequiredInput {
    param(
        [string]$Prompt,
        [switch]$Secret
    )

    while ($true) {
        if ($Secret) {
            $secure = Read-Host -Prompt $Prompt -AsSecureString
            $value = Convert-SecureToPlainText -Value $secure
        } else {
            $value = Read-Host -Prompt $Prompt
        }

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }

        Write-Warning 'Value is required.'
    }
}

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$DefaultNo = $true
    )

    while ($true) {
        $suffix = if ($DefaultNo) { '[y/N]' } else { '[Y/n]' }
        $answer = (Read-Host -Prompt "$Prompt $suffix").Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($answer)) { return (-not $DefaultNo) }
        if ($answer -in @('y', 'yes')) { return $true }
        if ($answer -in @('n', 'no')) { return $false }
        Write-Warning "Please answer y or n."
    }
}

function Save-BillingCredential {
    param(
        [string]$CredsFile,
        [string]$Id,
        [string]$Account,
        [string]$SfdcId,
        [string]$ApiKey
    )

    $existing = @()
    if (Test-Path $CredsFile) {
        $existing = @(Import-Clixml $CredsFile)
    }

    $newEntry = [PSCustomObject]@{
        Id      = $Id
        Account = $Account
        SFDCID  = $SfdcId
        Token   = (ConvertTo-SecureString -String $ApiKey -AsPlainText -Force)
    }

    $updated = @()
    $matched = $false
    foreach ($item in $existing) {
        if ($item.Id -eq $Id) {
            $updated += $newEntry
            $matched = $true
        } else {
            $updated += $item
        }
    }

    if (-not $matched) {
        $updated += $newEntry
    }

    $updated | Export-Clixml -Path $CredsFile
    return $matched
}

$Account = ''
$Token = ''
$credsFile = Join-Path $PSScriptRoot 'BillingCreds.xml'
$allCreds = @()
$runtimeCredentialInput = $false

if ($SFDCID -or $APIKey) {
    # Runtime parameters supplied (full or partial) — prompt for missing pieces.
    $runtimeCredentialInput = $true
    if (-not $SFDCID) {
        $SFDCID = Read-RequiredInput -Prompt 'Enter Salesforce Account ID (SFDCID)'
    }
    if (-not $APIKey) {
        $APIKey = Read-RequiredInput -Prompt 'Enter N-able Billing API key' -Secret
    }
    if (-not $AccountName) {
        $AccountName = "SFDCID-$SFDCID"
    }
    $Account = $AccountName
    $Token = $APIKey
    Write-Output "Using runtime credentials (SFDCID: $SFDCID)"
} else {
    # XML credential store mode (default) with interactive fallback.
    if (Test-Path $credsFile) {
        $allCreds = @(Import-Clixml $credsFile)
    } else {
        Write-Warning "BillingCreds.xml not found at: $credsFile"
    }

    if (-not $creds) {
        if ($allCreds.Count -gt 0) {
            $menuItems = @([PSCustomObject]@{ Id=''; Account='(EFile only — no API call)' }) +
                         ($allCreds | Select-Object Id, Account)
            $picked = $menuItems | Out-GridView -Title 'Select billing account' -OutputMode Single
            if ($null -eq $picked) {
                Write-Output 'No selection made in GridView; switching to runtime credential prompt.'
                $runtimeCredentialInput = $true
            } else {
                $creds = $picked.Id
            }
        } else {
            $runtimeCredentialInput = $true
        }
    }

    if (-not $runtimeCredentialInput -and $creds -ne '') {
        $entry = $allCreds | Where-Object { $_.Id -eq $creds }
        if (-not $entry) { Write-Warning "Unknown account ID: '$creds'"; exit 1 }
        $Account = $entry.Account
        $SFDCID  = $entry.SFDCID
        $Token   = $entry.Token | ConvertFrom-SecureString -AsPlainText
    }

    if ($runtimeCredentialInput) {
        $SFDCID = Read-RequiredInput -Prompt 'Enter Salesforce Account ID (SFDCID)'
        $APIKey = Read-RequiredInput -Prompt 'Enter N-able Billing API key' -Secret
        $Account = if ($AccountName) { $AccountName.Trim() } else { "SFDCID-$SFDCID" }
        $Token = $APIKey
        Write-Output "Using prompted credentials (SFDCID: $SFDCID)"
    }
    # else: EFile-only mode — $Account/$SFDCID/$Token stay empty; derived from EFile later
}

if ($runtimeCredentialInput -and $SFDCID -and $Token) {
    $storeChoice = Read-YesNo -Prompt 'Store these credentials securely in BillingCreds.xml?' -DefaultNo $true
    if ($storeChoice) {
        if ($allCreds.Count -eq 0 -and (Test-Path $credsFile)) {
            $allCreds = @(Import-Clixml $credsFile)
        }

        # Auto-generate next credential ID by incrementing existing prod# IDs
        $prodIds = @($allCreds | Where-Object { $_.Id -match '^prod(\d+)$' } | ForEach-Object { [int]($_.Id -replace '^prod(\d+)$', '$1') } | Sort-Object -Descending)
        $nextNum = if ($prodIds.Count -gt 0) { $prodIds[0] + 1 } else { 1 }
        $storeId = "prod$nextNum"
        Write-Output "Auto-generated credential ID: $storeId"

        # Prompt for account name if not provided via parameter
        if ($AccountName) {
            $storeAccount = $AccountName.Trim()
        } else {
            $storeAccount = Read-RequiredInput -Prompt 'Account display name to store'
        }
        $Account = $storeAccount

        $existingEntry = $allCreds | Where-Object { $_.Id -eq $storeId }
        if ($existingEntry) {
            $overwrite = Read-YesNo -Prompt "Credential ID '$storeId' exists. Overwrite it?" -DefaultNo $true
            if (-not $overwrite) {
                Write-Output 'Credential save skipped (existing entry preserved).'
            } else {
                $wasUpdate = Save-BillingCredential -CredsFile $credsFile -Id $storeId -Account $storeAccount -SfdcId $SFDCID -ApiKey $Token
                if ($wasUpdate) {
                    Write-Output "Updated secure credential '$storeId' in $credsFile"
                } else {
                    Write-Output "Saved secure credential '$storeId' in $credsFile"
                }
            }
        } else {
            $wasUpdate = Save-BillingCredential -CredsFile $credsFile -Id $storeId -Account $storeAccount -SfdcId $SFDCID -ApiKey $Token
            if ($wasUpdate) {
                Write-Output "Updated secure credential '$storeId' in $credsFile"
            } else {
                Write-Output "Saved secure credential '$storeId' in $credsFile"
            }
        }
    }
}

$Period = "$($Timeframe.Substring(0,4))-$($Timeframe.Substring(4))"
Write-Output ""
Write-Output "Account   : $Account"
Write-Output "Period    : $Period"

$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptStart = Get-Date
$ServiceMap  = [ordered]@{}
$Summary     = [System.Collections.Generic.List[PSCustomObject]]::new()
$Rows        = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── API Call ──────────────────────────────────────────────────────────────────

if ($Token) {
$Headers = @{ "x-api-key" = $Token }
$Uri     = "https://prod-nableboomi.n-able.com:443/ws/rest/V1/PSA_API/PSAUsageByTimeFrame/$SFDCID/$Timeframe"

Write-Output "Calling: $Uri"

$startTime    = Get-Date
try {
    $Response = Invoke-RestMethod -Uri $Uri -Method POST -Headers $Headers
} catch {
    Write-Warning "API call failed: $_"
    exit 1
}
$elapsed    = ((Get-Date) - $startTime).TotalSeconds
$ApiSeconds = [math]::Round($elapsed, 2)
Write-Output "$ApiSeconds seconds elapsed`n"

if (-not $Response.UsageDetailDevices) {
    Write-Warning "No UsageDetailDevices returned for $Account / $Period"
    exit
}

Write-Output "Contracts returned: $($Response.UsageDetailDevices.Count)"
} # end API call

# ── Discover all BillingServiceNames in this response ────────────────────────
# Build an ordered map: BillingServiceName → safe PS column name + short display label
# This makes the script self-adapting — new products appear automatically.

Function Get-ServiceColName {
    param([string]$ServiceName)
    # Strip product prefix (everything up to and including the first ' | ')
    $short = if ($ServiceName -match '\|\s*(.+)$') { $Matches[1].Trim() } else { $ServiceName }
    # Sanitize to a safe PS property name
    ($short -replace '[^a-zA-Z0-9]', '_' -replace '_+', '_').Trim('_')
}

Function Get-ServiceDisplayName {
    param([string]$ServiceName)
    # Short label for HTML column header — strip product prefix
    $d = if ($ServiceName -match '\|\s*(.+)$') { $Matches[1].Trim() } else { $ServiceName }
    # Shorten common verbose phrases
    $d = $d -replace '(?i)endpoint\s+detection\s+and\s+response\s*[-–]?\s*', 'EDR '
    $d = $d -replace '(?i)virtual\s+machine\s+server', 'Virtual Server'
    $d.Trim()
}

# Split a display label into 2 lines by inserting <br> near the word boundary closest to the midpoint
Function Split-DisplayLabel {
    param([string]$Label)
    if ($Label.Length -le 10) { return [System.Web.HttpUtility]::HtmlEncode($Label) }
    $words = $Label -split ' '
    if ($words.Count -lt 2) { return [System.Web.HttpUtility]::HtmlEncode($Label) }
    $mid = $Label.Length / 2
    $best = 0; $bestDist = [int]::MaxValue; $pos = 0
    for ($i = 0; $i -lt $words.Count - 1; $i++) {
        $pos += $words[$i].Length + 1   # +1 for the space
        $dist = [math]::Abs($pos - $mid)
        if ($dist -lt $bestDist) { $bestDist = $dist; $best = $i }
    }
    $line1 = [System.Web.HttpUtility]::HtmlEncode(($words[0..$best] -join ' '))
    $line2 = [System.Web.HttpUtility]::HtmlEncode(($words[($best+1)..($words.Count-1)] -join ' '))
    "$line1<br>$line2"
}

# Priority sort: Server → VirtualServer → Workstation → Documents → Continuity →
#   M365 → Per GB → other Cove/CDP → EDR → everything else
Function Get-ServiceSortKey {
    param([string]$ServiceName)
    $d = (Get-ServiceDisplayName $ServiceName).ToLower()
    $p = if ($ServiceName -match '^([^|]+)\|') { $Matches[1].Trim().ToLower() } else { '' }
    if     ($d -match 'virtual.*server')    { '2_virtual_server' }
    elseif ($d -match '^server$')           { '1_server' }
    elseif ($d -match 'workstation')        { '3_workstation' }
    elseif ($d -match 'document')           { '4_documents' }
    elseif ($d -match 'continuit')          { '5_continuity' }
    elseif ($d -match '365|microsoft.*365') { '6_m365' }
    elseif ($d -match 'per.{0,2}gb')        { '7_pergb' }
    elseif ($p -match 'edr' -or $d -match '^edr\b') {
        if     ($d -match 'complete')        { '9_edr_1_complete' }
        elseif ($d -match 'control')         { '9_edr_2_control' }
        elseif ($d -match 'vigilance')       { '9_edr_3_vigilance' }
        elseif ($d -match 'deep.{0,3}vis')   { '9_edr_4_deepvis' }
        else                                 { '9_edr_5_other' }
    }
    elseif ($p -match 'cdp|cove|backup')    { '8_cove_other' }
    elseif ($p -match 'n.sight|nsight') {
        if     ($d -match '^node$')           { 'C_ns_1_node' }
        elseif ($d -match 'antivirus|av')     { 'C_ns_2_av' }
        elseif ($d -match 'web.protect')      { 'C_ns_3_web' }
        elseif ($d -match 'team.?view')       { 'C_ns_4_tv' }
        elseif ($d -match 'platform')         { 'C_ns_5_plat' }
        else                                  { 'C_ns_9_other' }
    }
    else                                    { 'A_other' }
}

# Map EFile product prefix to a family tab name
Function Get-EFileFamily {
    param([string]$Product)
    $prefix = if ($Product -match '^([^|]+)\|') { $Matches[1].Trim() } else { $Product.Trim() }
    if     ($prefix -match 'Cove Data Protection') { 'Cove (EFile)' }
    elseif ($prefix -match 'N-able Security' -and $Product -match '\|\s*N-able Advanced MDR') { 'Adlumin' }
    elseif ($prefix -match 'N-able Security')       { 'Security (EFile)' }
    elseif ($prefix -match '^N-central')            { 'N-central' }
    elseif ($prefix -match '^Adlumin')              { 'Adlumin' }
    elseif ($prefix -match 'N-able Take Control')   { 'Take Control' }
    elseif ($prefix -match 'N-able N-sight')        { 'N-sight' }
    elseif ($prefix -match 'Cloud Commander')       { 'Cloud Commander' }
    else                                            { $prefix }
}

# Build one EFile summary row by summing Quantity per product across $Rows
Function New-EFSummaryRow {
    param($Family, $Level, $Indent, $Name, $Rows, $SvcMap)
    $rowData = [ordered]@{ Family = $Family; Level = $Level; Indent = $Indent; Name = $Name }
    foreach ($svc in $SvcMap.Keys) {
        $col    = $SvcMap[$svc].Col
        $isGB   = $svc -like '*Per GB*'
        $rawSum = ($Rows | Where-Object { $_.Product -eq $svc } | ForEach-Object { [decimal]$_.'Quantity' } | Measure-Object -Sum).Sum
        if (-not $rawSum) { $rawSum = 0 }
        $rowData[$col] = if ($isGB) { [math]::Round([decimal]$rawSum, 1) } else { [decimal]$rawSum }
    }
    [PSCustomObject]$rowData
}

# Collect unique service names across ALL contracts, sort by display priority then alphabetically within each group
if ($Token) {
$AllServiceNames = $Response.UsageDetailDevices |
    ForEach-Object { $_.Clients } |
    ForEach-Object { $_.Devices } |
    ForEach-Object { $_.TotalByService } |
    Select-Object -ExpandProperty BillingServiceName -Unique |
    Sort-Object -Property @{ Expression = { Get-ServiceSortKey $_ }; Ascending = $true },
                          @{ Expression = { $_ };                    Ascending = $true }

# Build ordered map with unique safe column names
$ServiceMap = [ordered]@{}   # BillingServiceName -> @{ Col; Display }
$usedCols   = [System.Collections.Generic.HashSet[string]]::new()
foreach ($svc in $AllServiceNames) {
    $col  = Get-ServiceColName $svc
    $base = $col; $n = 2
    while (-not $usedCols.Add($col)) { $col = "${base}_$n"; $n++ }
    $ServiceMap[$svc] = @{ Col = $col; Display = Get-ServiceDisplayName $svc }
}

Write-Output "Billing services discovered: $($ServiceMap.Count)"
$ServiceMap.GetEnumerator() | ForEach-Object { Write-Output "  [$($_.Value.Col)] $($_.Key)" }

# ── Flatten (one row per device, billing services as dynamic columns) ─────────

$Rows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Contract in $Response.UsageDetailDevices) {
    foreach ($Client in $Contract.Clients) {
        foreach ($Device in $Client.Devices) {

            $deviceRow = [ordered]@{
                # ── Contract ─────────────────────────────────────────────────
                Account           = $Account
                Timeframe         = $Timeframe
                ContractNumber    = $Contract.ContractNumber
                ContractStatus    = $Contract.ContractStatus
                ProductName       = $Contract.ProductName
                # ── Hierarchy (L3–L8) ─────────────────────────────────────────
                Distributor       = $Device.UsageLevel_3           # L3
                DistributorRef    = $Device.Customer_Reference_L3
                SubDistributor    = $Device.UsageLevel_4           # L4
                SubDistributorRef = $Device.Customer_Reference_L4
                Reseller          = $Device.UsageLevel_5           # L5
                ResellerRef       = $Device.Customer_Reference_L5
                ServiceOrg        = $Device.UsageLevel_6           # L6
                ServiceOrgRef     = $Device.Customer_Reference_L6
                EndCustomer       = $Device.UsageCompanyName       # L7
                EndCustomerRef    = $Device.Customer_Reference_L7
                Site              = $Device.UsageSiteName          # L8
                SiteRef           = $Device.Customer_Reference_L8
                # ── Device ────────────────────────────────────────────────────
                DeviceName        = $Device.Detail
                DeviceId          = $Device.DetailId
                ClientId          = $Client.ClientId
            }

            # Append one column per discovered service (exact name match, no wildcards)
            foreach ($svc in $ServiceMap.Keys) {
                $col = $ServiceMap[$svc].Col
                $val = ($Device.TotalByService |
                        Where-Object { $_.BillingServiceName -eq $svc } |
                        Measure-Object -Property Quantity -Sum).Sum
                $deviceRow[$col] = if ($val) { $val } else { $null }
            }

            $Rows.Add([PSCustomObject]$deviceRow)
        }
    }
}

Write-Output "Total rows: $($Rows.Count)"

# ── Export ────────────────────────────────────────────────────────────────────

$SafeName    = $Account -replace '[^a-zA-Z0-9_\-]', '_'
$OutputDir   = Join-Path $PSScriptRoot "Reports\$Timeframe\$SafeName"
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$OutputFile  = Join-Path $OutputDir "${Timestamp}_${SafeName}_Usage_${Timeframe}.csv"

$Rows | Export-Csv -Path $OutputFile -NoTypeInformation
Write-Output "Saved: $OutputFile"

if (-not $NoGridView) { $Rows | Out-GridView -Title "$Account | $Period | Device Usage ($($Rows.Count) rows)" }

# ── Summary Report (hierarchy per ProductName, mirrors Cove UI table) ────────
#
#  Columns: Partner/Customer/Site | PhysServers | VirtServers | Workstations
#           | M365Users | CoveContinuity | SelectedGB | UsedGB | Efficiency
#
#  Hierarchy levels (nested by Indent depth):
#    0 = Product total
#    1 = Reseller
#    2 = EndCustomer   (direct — no ServiceOrg assigned)
#    3 = Site           (under direct EndCustomer)
#    2 = ServiceOrg     (when ServiceOrg is assigned)
#    3 = EndCustomer    (under ServiceOrg)
#    4 = Site           (under EndCustomer under ServiceOrg)
#
#  Named ServiceOrgs appear first under each Reseller (collapsible),
#  followed by direct EndCustomers (no ServiceOrg assigned).
# ─────────────────────────────────────────────────────────────────────────────

Function Get-GroupSum {
    param($Group, $Prop)
    ($Group | ForEach-Object {
        $v = $_.$Prop
        if ($null -ne $v -and $v -ne '') { [decimal]$v } else { [decimal]0 }
    } | Measure-Object -Sum).Sum
}

Function Get-Efficiency {
    param([decimal]$SelGB, [decimal]$UsedGB)
    if ($SelGB -gt 0) { "$([math]::Round($UsedGB / $SelGB * 100, 1))%" } else { 'N/A' }
}

Function New-SummaryRow {
    param($ProductName, $Level, $Indent, $Name, $Ref = '', $Group)
    # Fixed fields
    $rowData = [ordered]@{
        ProductName = $ProductName
        Level       = $Level
        Indent      = $Indent
        Name        = $Name
        Ref         = $Ref
    }
    # Dynamic service columns — sum each discovered service column across the group
    foreach ($svc in $ServiceMap.Keys) {
        $col    = $ServiceMap[$svc].Col
        $isGB   = $svc -like '*Per GB*'
        $rawSum = ($Group | ForEach-Object {
            $v = $_.$col
            if ($null -ne $v -and $v -ne '') { [decimal]$v } else { [decimal]0 }
        } | Measure-Object -Sum).Sum
        $rowData[$col] = if ($isGB) { [math]::Round($rawSum, 3) } else { [int]$rawSum }
    }
    [PSCustomObject]$rowData
}

$Summary = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($prodGroup in ($Rows | Group-Object ProductName | Sort-Object Name)) {
    $prodName = $prodGroup.Name
    $prodRows = $prodGroup.Group
    # EDR products: roll up to EndCustomer only — device-level site rows aren't meaningful
    $isEDRProduct = $prodName -match '(?i)edr|endpoint.detection'

    # Product total row
    $Summary.Add((New-SummaryRow $prodName 'Total' 0 $prodName '' $prodRows))

    foreach ($resGroup in ($prodRows | Group-Object Reseller | Sort-Object Name)) {
        $resName = if ($resGroup.Name) { $resGroup.Name } else { '(No Reseller)' }
        $resRef  = $resGroup.Group[0].ResellerRef
        $Summary.Add((New-SummaryRow $prodName 'Reseller' 1 $resName $resRef $resGroup.Group))

        # Split devices under this Reseller: direct (no ServiceOrg) vs. under a named ServiceOrg
        $directECs = $resGroup.Group | Where-Object { -not $_.ServiceOrg -or $_.ServiceOrg.Trim() -eq '' }
        $soRows    = $resGroup.Group | Where-Object {  $_.ServiceOrg  -and $_.ServiceOrg.Trim() -ne '' }

        # ServiceOrgs first — depth 2; their EndCustomers at depth 3; Sites at depth 4
        foreach ($soGroup in ($soRows | Group-Object ServiceOrg | Sort-Object Name)) {
            $soName = $soGroup.Name
            $soRef  = $soGroup.Group[0].ServiceOrgRef
            $Summary.Add((New-SummaryRow $prodName 'ServiceOrg' 2 $soName $soRef $soGroup.Group))

            foreach ($ecGroup in ($soGroup.Group | Group-Object EndCustomer | Sort-Object Name)) {
                $ecName = if ($ecGroup.Name) { $ecGroup.Name } else { '(No End Customer)' }
                $ecRef  = $ecGroup.Group[0].EndCustomerRef
                $Summary.Add((New-SummaryRow $prodName 'EndCustomer' 3 $ecName $ecRef $ecGroup.Group))

                if (-not $isEDRProduct) {
                    foreach ($siteGroup in ($ecGroup.Group | Group-Object Site | Where-Object { $_.Name } | Sort-Object Name)) {
                        $siteRef = $siteGroup.Group[0].SiteRef
                        $Summary.Add((New-SummaryRow $prodName 'Site' 4 $siteGroup.Name $siteRef $siteGroup.Group))
                    }
                }
            }
        }

        # Direct EndCustomers (no ServiceOrg) — depth 2, Sites at depth 3
        foreach ($ecGroup in ($directECs | Group-Object EndCustomer | Sort-Object Name)) {
            $ecName = if ($ecGroup.Name) { $ecGroup.Name } else { '(No End Customer)' }
            $ecRef  = $ecGroup.Group[0].EndCustomerRef
            $Summary.Add((New-SummaryRow $prodName 'EndCustomer' 2 $ecName $ecRef $ecGroup.Group))

            if (-not $isEDRProduct) {
                foreach ($siteGroup in ($ecGroup.Group | Group-Object Site | Where-Object { $_.Name } | Sort-Object Name)) {
                    $siteRef = $siteGroup.Group[0].SiteRef
                    $Summary.Add((New-SummaryRow $prodName 'Site' 3 $siteGroup.Name $siteRef $siteGroup.Group))
                }
            }
        }
    }
}
} # end if ($Token) $AllServiceNames/$Rows/$Summary block

Write-Output "Summary rows: $($Summary.Count)"

if ($Token) {
    $SummaryFile = Join-Path $OutputDir "${Timestamp}_${SafeName}_Summary_${Timeframe}.csv"
    $Summary | Export-Csv -Path $SummaryFile -NoTypeInformation
    Write-Output "Summary saved: $SummaryFile"
}

# ── Collapsible HTML Report (one tab per product) ────────────────────────────

Add-Type -AssemblyName System.Web

Function Format-Num  { param($v) if ($v -eq 0) { '<span class="zero">0</span>' } else { $v } }
Function Format-GB   { param($v) if (!$v -or $v -eq 0) { '<span class="zero">0.0</span>' } else { '{0:N1}' -f [decimal]$v } }
Function Format-Delta {
    param([decimal]$current, [decimal]$prior, [bool]$isGB)
    $diff = $current - $prior
    if ($diff -eq 0) { return '' }
    $abs  = [math]::Abs($diff)
    $fmt  = if ($isGB) { '{0:N1}' -f $abs } else { '{0:N0}' -f $abs }
    if ($diff -gt 0) { " <span class='delta-up' title='Prior: $(if($isGB){"{0:N1}" -f $prior}else{$prior})'>&#9650;$fmt</span>" }
    else             { " <span class='delta-down' title='Prior: $(if($isGB){"{0:N1}" -f $prior}else{$prior})'>&#9660;$fmt</span>" }
}

# ── Cross-reference helpers ──────────────────────────────────────────────────
Function Norm-CustName {
    param([string]$s)
    $s = $s.ToLower()
    $s = $s -replace '\[.*?\]', ''
    $s = $s -replace '[^a-z0-9 ]', ' '
    $s = ($s -replace ' +', ' ').Trim()
    $s = $s -replace '\b(inc|llc|ltd|corp|co|company|the|and|of|a|pllc|pc|pa|llp|lp|dba)\b', ''
    ($s -replace ' +', ' ').Trim()
}
Function Get-MatchKeys {
    param([string]$s)
    $keys = [System.Collections.Generic.HashSet[string]]::new()
    # Key 1: basic normalize
    $n = Norm-CustName $s
    if ($n) { [void]$keys.Add($n) }
    # Key 2: CamelCase split + normalize
    $c = Norm-CustName ($s -replace '([a-z]{2,})([A-Z])', '$1 $2')
    if ($c) { [void]$keys.Add($c) }
    # Key 3: strip all non-alnum (after normalize removes stop words)
    $st = $n -replace '[^a-z0-9]', ''
    if ($st) { [void]$keys.Add($st) }
    # Key 4: CamelCase + strip all non-alnum
    $cs = $c -replace '[^a-z0-9]', ''
    if ($cs) { [void]$keys.Add($cs) }
    # Key 5: depluralize (strip trailing s from each word)
    $dp = (($n -split ' ') | ForEach-Object { $_ -replace 's$', '' } | Where-Object { $_ }) -join ' '
    if ($dp) { [void]$keys.Add($dp) }
    # Key 6: CamelCase + depluralize
    $cdp = (($c -split ' ') | ForEach-Object { $_ -replace 's$', '' } | Where-Object { $_ }) -join ' '
    if ($cdp) { [void]$keys.Add($cdp) }
    ,$keys
}
Function Get-TokenOverlap {
    param([string]$a, [string]$b)
    $ta = @($a -split ' ' | Where-Object { $_.Length -gt 2 })
    $tb = @($b -split ' ' | Where-Object { $_.Length -gt 2 })
    if ($ta.Count -eq 0 -or $tb.Count -eq 0) { return 0.0 }
    $shared = ($ta | Where-Object { $tb -contains $_ }).Count
    # Use max(|A|,|B|) so a short name that is a subset of a longer name does NOT score 100%
    [double]$shared / [math]::Max($ta.Count, $tb.Count)
}

# Group summary rows by product — each becomes its own tab
$prodGroups = $Summary | Group-Object ProductName | Sort-Object Name

$tabNavHtml    = [System.Text.StringBuilder]::new()
$tabPanelHtml  = [System.Text.StringBuilder]::new()
$tabIdx        = 0

foreach ($pg in $prodGroups) {
    $prodName    = $pg.Name
    $prodSummary = $pg.Group

    # Only show service columns that have at least one non-zero value in this product's Total row
    $totalRow   = $prodSummary | Where-Object { $_.Level -eq 'Total' } | Select-Object -First 1
    $activeSvcs = $ServiceMap.Keys | Where-Object {
        $col = $ServiceMap[$_].Col
        $val = if ($totalRow) { $totalRow.$col } else { $null }
        $val -and [decimal]$val -ne 0
    }
    if (-not $activeSvcs) { $activeSvcs = $ServiceMap.Keys }   # fallback: show all

    # Hide Reference ID column if no rows have a Ref value
    $hasRefs = ($prodSummary | Where-Object { $_.Ref -and $_.Ref.Trim() -ne '' }).Count -gt 0

    # ── Build table rows for this product ──
    $rowId       = 0
    $depthParent = @{}
    $htmlRows    = [System.Text.StringBuilder]::new()
    $pfx         = "t$tabIdx"

    for ($i = 0; $i -lt $prodSummary.Count; $i++) {
        $row   = $prodSummary[$i]
        $depth = [int]$row.Indent
        $rid   = "${pfx}_$(if ($depth -eq 0){'p'}else{'r'})$rowId"
        $rowId++
        $depthParent[$depth] = $rid

        $isParent     = ($i + 1 -lt $prodSummary.Count) -and ([int]$prodSummary[$i+1].Indent -gt $depth)
        $parentAttr   = if ($depth -eq 0) { '' }            else { "data-parent='$($depthParent[$depth-1])'" }
        $displayStyle = if ($depth -eq 0) { '' }            else { "style='display:none'" }
        $collClass    = if ($isParent)    { 'collapsible' } else { '' }
        $clickAttr    = if ($isParent)    { "onclick='toggle(this)'" } else { '' }
        $togSpan      = if ($isParent)    { "<span class='tog'>&#9658;</span> " } else { "<span class='tog-leaf'></span> " }

        $cls = switch ($row.Level) {
            'Total'       { 'row-total' }
            'Reseller'    { 'row-reseller' }
            'ServiceOrg'  { 'row-serviceorg' }
            'EndCustomer' { 'row-endcustomer' }
            'Site'        { 'row-site' }
            default       { 'row-other' }
        }

        $null = $htmlRows.AppendLine("<tr class='$cls $collClass' data-id='$rid' $parentAttr data-depth='$depth' data-open='0' $displayStyle $clickAttr>")
        $levelBadge = switch ($row.Level) {
            'Reseller'    { '<span class="badge badge-reseller">Reseller</span>' }
            'ServiceOrg'  { '<span class="badge badge-serviceorg">Service Organization</span>' }
            'EndCustomer' { '<span class="badge badge-endcustomer">End Customer</span>' }
            'Site'        { '<span class="badge badge-site">Site</span>' }
            default       { '' }
        }
        $nameWords = $row.Name -split '\s+', 2
        $nameHtml  = if ($nameWords.Count -gt 1) {
            "$([System.Web.HttpUtility]::HtmlEncode($nameWords[0])) <span class='priv-rest'>$([System.Web.HttpUtility]::HtmlEncode($nameWords[1]))</span>"
        } else { [System.Web.HttpUtility]::HtmlEncode($row.Name) }
        $null = $htmlRows.AppendLine("  <td class='name-col depth$depth'>$togSpan$nameHtml $levelBadge</td>")
        if ($hasRefs) {
            $refCell = if ($row.Ref) { "<span class='ref-id'>$([System.Web.HttpUtility]::HtmlEncode($row.Ref))</span>" } else { "<span class='zero'>&mdash;</span>" }
            $null = $htmlRows.AppendLine("  <td class='ref-col'>$refCell</td>")
        }
        foreach ($svc in $activeSvcs) {
            $col  = $ServiceMap[$svc].Col
            $val  = $row.$col
            $isGB = $svc -like '*Per GB*'
            $null = $htmlRows.AppendLine("  <td>$(if ($isGB) { Format-GB $val } else { Format-Num ([int]$val) })</td>")
        }
        $null = $htmlRows.AppendLine('</tr>')
    }

    # ── Tab nav button ──
    $activeClass = if ($tabIdx -eq 0) { 'active' } else { '' }
    # Split label roughly in half at the nearest | to the midpoint
    $tabParts = $prodName -split '\s*\|\s*'
    if ($tabParts.Count -ge 2) {
        $mid   = [int][math]::Round($tabParts.Count / 2)
        $line1 = [System.Web.HttpUtility]::HtmlEncode(($tabParts[0..($mid-1)] -join ' | '))
        $line2 = [System.Web.HttpUtility]::HtmlEncode(($tabParts[$mid..($tabParts.Count-1)] -join ' | '))
        $tabLabel = "$line1<br>$line2"
    } else {
        $tabLabel = [System.Web.HttpUtility]::HtmlEncode($prodName)
    }
    $null = $tabNavHtml.AppendLine("<button class='tab-btn $activeClass' onclick='showTab(""tab$tabIdx"",this)'>$tabLabel</button>")

        # Column headers for this tab — GB cols wider, commit cols annotated
        $thHtml = ($activeSvcs | ForEach-Object {
            $d    = $ServiceMap[$_].Display
            $isGB = $_ -like '*Per GB*'
            $lbl  = Split-DisplayLabel $d
            if ($isGB) { "      <th style='min-width:100px'>$lbl</th>" } else { "      <th>$lbl</th>" }
        }) -join "`n"

    # ── Tab panel ──
    $panelStyle = if ($tabIdx -eq 0) { '' } else { "style='display:none'" }
    $null = $tabPanelHtml.AppendLine("<div id='tab$tabIdx' class='tab-panel' $panelStyle>")
    $null = $tabPanelHtml.AppendLine("<div class='tab-actions'><button class='btn' onclick='expandAll(""tab$tabIdx"")'>Expand All</button> <button class='btn' onclick='collapseAll(""tab$tabIdx"")'>Collapse All</button> <input type='text' class='search-box' id='search-tab$tabIdx' placeholder='Filter rows...' oninput='filterRows(""tab$tabIdx"",this)'> <button class='btn btn-clear' onclick='clearFilter(""tab$tabIdx"")' title='Clear'>&#x2715;</button> <label class='preserve-lbl'><input type='checkbox' class='preserve-cb' onchange='syncPreserve(this)'> Preserve filter</label></div>")
    $null = $tabPanelHtml.AppendLine("<div class='tab-scroll'><table><thead><tr>")
    $null = $tabPanelHtml.AppendLine("  <th>Partner / Customer / Site</th>")
    if ($hasRefs) { $null = $tabPanelHtml.AppendLine("  <th style='text-align:left'>Reference ID</th>") }
    $null = $tabPanelHtml.AppendLine($thHtml)
    $null = $tabPanelHtml.AppendLine("</tr></thead><tbody>")
    $null = $tabPanelHtml.AppendLine($htmlRows.ToString())
    $null = $tabPanelHtml.AppendLine("</tbody></table></div></div>")

    $tabIdx++
}

# ── EFile folder scan + interactive picklist ─────────────────────────────────
$EFilePath  = ''; $EFilePath2 = ''; $EFilePrior = ''; $EFilePrior2 = ''
$scanFolder = if ($EFileFolder -and (Test-Path $EFileFolder)) { $EFileFolder } else { $PSScriptRoot }
$allEFiles  = Get-ChildItem $scanFolder -Filter 'EFile_*.csv' -ErrorAction SilentlyContinue

if ($allEFiles.Count -gt 0) {
    # Build index: one entry per file with Account + YYYYMM period
    $efileIndex = $allEFiles | ForEach-Object {
        $row = Import-Csv $_.FullName | Select-Object -First 1
        $pf  = $row.'Period From'   # MM-DD-YYYY
        $ym  = if ($pf -match '^(\d{2})-\d{2}-(\d{4})$') { "$($Matches[2])$($Matches[1])" } else { '000000' }
        [PSCustomObject]@{ File = $_.FullName; Account = $row.'Account Name'; Period = $ym }
    } | Where-Object { $_.Account -and $_.Period -ne '000000' }

    if ($efileIndex.Count -gt 0) {
        # Step 1 – pick account
        # If an API account is selected, try to auto-match against EFile account names
        $uniqueAccounts = @($efileIndex | Select-Object -ExpandProperty Account -Unique | Sort-Object)
        $pickedAccount  = $null

        if ($Account) {
            # Score every EFile account against the API account name.
            # Get-MatchKeys returns ,$hashset (array wrapper) — enumerate [0] to get the HashSet.
            $apiKeySet = @(Get-MatchKeys $Account)[0]   # HashSet[string]
            $apiNorm   = Norm-CustName $Account
            $scored  = $uniqueAccounts | ForEach-Object {
                $efKeySet = @(Get-MatchKeys $_)[0]      # HashSet[string]
                $efNorm   = Norm-CustName $_
                # Key match: any shared normalised key between the two sets
                $keyHit   = ($apiKeySet | Where-Object { $efKeySet.Contains($_) }).Count -gt 0
                # Exact normalised name match (catches AIRIAM == Airiam)
                $exactHit = $apiNorm -eq $efNorm
                # Token overlap fallback
                $overlap  = Get-TokenOverlap $apiNorm $efNorm
                [PSCustomObject]@{ Name = $_; KeyHit = ($keyHit -or $exactHit); Overlap = $overlap }
            }
            $goodMatches = @($scored | Where-Object { $_.KeyHit -or $_.Overlap -ge 0.5 } | Sort-Object -Property KeyHit,Overlap -Descending)

            if ($goodMatches.Count -eq 1) {
                # Single clear match — use it automatically
                $pickedAccount = $goodMatches[0].Name
                Write-Output "EFile auto-matched: '$pickedAccount' (matched API account '$Account')"
            } elseif ($goodMatches.Count -gt 1) {
                # Multiple matches — show only the plausible ones
                $pickedAccount = ($goodMatches.Name) | Out-GridView -Title "Select EFile account for '$Account'" -OutputMode Single
            } else {
                # No match found — fall back to full list
                Write-Output "No EFile match found for '$Account' — showing all accounts"
                if ($uniqueAccounts.Count -eq 1) {
                    $pickedAccount = $uniqueAccounts[0]
                } else {
                    $pickedAccount = $uniqueAccounts | Out-GridView -Title 'Select EFile Account' -OutputMode Single
                }
            }
        } elseif ($uniqueAccounts.Count -eq 1) {
            $pickedAccount = $uniqueAccounts[0]
        } else {
            $pickedAccount = $uniqueAccounts | Out-GridView -Title 'Select EFile Account' -OutputMode Single
        }

        if ($pickedAccount) {
            $acctIndex = @($efileIndex | Where-Object { $_.Account -eq $pickedAccount })

            # Step 2 – pick billing period (skip picker if only one)
            $uniquePeriods = @($acctIndex | Select-Object -ExpandProperty Period -Unique | Sort-Object -Descending)
            if ($uniquePeriods.Count -eq 1) {
                $pickedPeriod = $uniquePeriods[0]
            } else {
                $pickedPeriod = $uniquePeriods | Out-GridView -Title "Select Billing Period for $pickedAccount" -OutputMode Single
            }

            if ($pickedPeriod) {
                # Current-period files (up to 2)
                $curFiles   = @($acctIndex | Where-Object { $_.Period -eq $pickedPeriod })
                $EFilePath  = $curFiles[0].File
                if ($curFiles.Count -gt 1) { $EFilePath2 = $curFiles[1].File }

                # Prior-period files — auto-detect, no prompt needed
                $priorMoNum = [int]$pickedPeriod.Substring(4, 2) - 1
                $priorYrNum = [int]$pickedPeriod.Substring(0, 4)
                if ($priorMoNum -lt 1) { $priorMoNum = 12; $priorYrNum-- }
                $priorPeriodYM = "$priorYrNum$($priorMoNum.ToString('00'))"
                $priorFiles    = @($acctIndex | Where-Object { $_.Period -eq $priorPeriodYM })
                if ($priorFiles.Count -gt 0) { $EFilePrior  = $priorFiles[0].File }
                if ($priorFiles.Count -gt 1) { $EFilePrior2 = $priorFiles[1].File }
            }
        }
    }
}

# ── EFile Tabs (per product family) ──────────────────────────────────────────
$EFileCustomers = [ordered]@{}   # keyed by family name: list of EndCustomer names for cross-ref

if ($EFilePath -and (Test-Path $EFilePath)) {
    Write-Output ""
    Write-Output "EFile: $EFilePath"
    $EFRaw     = Import-Csv $EFilePath
    $EFAccount = ($EFRaw | Select-Object -First 1).'Account Name'
    $EFInvoice = ($EFRaw | Select-Object -First 1).'Invoice Number'
    $EFPeriod  = ($EFRaw | Select-Object -First 1).'Period From'  # e.g. '03-01-2026'
    # Merge second current-month EFile if provided
    if ($EFilePath2 -and (Test-Path $EFilePath2)) {
        $EFRaw2     = Import-Csv $EFilePath2
        $EFInvoice2 = ($EFRaw2 | Select-Object -First 1).'Invoice Number'
        $EFRaw      = @($EFRaw) + @($EFRaw2)
        $EFInvoice  = "$EFInvoice + $EFInvoice2"
        Write-Output "EFile2: $EFilePath2 (Invoice $EFInvoice2)"
    }
    $EFRows    = $EFRaw  # all rows including blank-customer global SKUs (N-central, Take Control etc.)
    Write-Output "EFile Account : $EFAccount"
    Write-Output "EFile Invoice : $EFInvoice"
    # EFile-only mode: derive $Account, $SafeName, $OutputDir from EFile when no API creds
    if (-not $Token) {
        $Account   = $EFAccount
        $SafeName  = $Account -replace '[^a-zA-Z0-9_\-]', '_'
        $OutputDir = Join-Path $PSScriptRoot "Reports\$Timeframe\$SafeName"
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    # ── Load prior-period EFile for month-over-month deltas ──────────
    $hasPrior    = $false
    $priorLookup     = @{}   # priorLookup[currentCustName][product] = qty
    $priorTotals     = @{}   # priorTotals[product] = total qty (raw file, incl. global SKUs)
    $priorGlobalSku  = @{}   # product → qty for blank-customer (contract/global) rows in prior file
    $churnedByFamily = @{}   # churnedByFamily[family] = @( @{Name; Rows} )
    if ($EFilePrior -and (Test-Path $EFilePrior)) {
        $hasPrior    = $true
        $EFPriorRaw  = Import-Csv $EFilePrior
        $priorInv    = ($EFPriorRaw | Select-Object -First 1).'Invoice Number'
        $priorPeriod = ($EFPriorRaw | Select-Object -First 1).'Period From'  # e.g. '02-01-2026'
        # Merge second prior-month EFile if provided
        if ($EFilePrior2 -and (Test-Path $EFilePrior2)) {
            $EFPriorRaw2 = Import-Csv $EFilePrior2
            $priorInv2   = ($EFPriorRaw2 | Select-Object -First 1).'Invoice Number'
            $EFPriorRaw  = @($EFPriorRaw) + @($EFPriorRaw2)
            $priorInv    = "$priorInv + $priorInv2"
            Write-Output "EFile Prior2  : $EFilePrior2 (Invoice $priorInv2)"
        }
        $EFPriorRows = $EFPriorRaw  # all rows including blank-customer global SKUs
        Write-Output "EFile Prior   : $EFilePrior (Invoice $priorInv)"

        # Build prior customer list keyed by match-keys for fast lookup
        $priorCusts     = @($EFPriorRows | Select-Object -ExpandProperty 'Usage Customer' -Unique | Where-Object { $_ })
        $priorKeyMap    = @{}   # matchKey → list of priorCustNames
        foreach ($pc in $priorCusts) {
            $pKeys = Get-MatchKeys $pc
            foreach ($pk in $pKeys) {
                if (-not $priorKeyMap[$pk]) { $priorKeyMap[$pk] = [System.Collections.Generic.List[string]]::new() }
                if (-not $priorKeyMap[$pk].Contains($pc)) { $priorKeyMap[$pk].Add($pc) }
            }
        }

        # Map current customer names → prior customer names using multi-key matching
        # Key-based: allow N:1 (same cust different casing across families can share prior)
        # Fuzzy: 1:1 only (prevent false positives)
        $currentCusts    = @($EFRows | Select-Object -ExpandProperty 'Usage Customer' -Unique | Where-Object { $_ })
        $cur2prior       = @{}   # currentCustName → priorCustName
        $matchedPrior    = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($cc in $currentCusts) {
            $cKeys = Get-MatchKeys $cc
            foreach ($ck in $cKeys) {
                if ($priorKeyMap[$ck]) {
                    $cur2prior[$cc] = $priorKeyMap[$ck][0]
                    [void]$matchedPrior.Add($priorKeyMap[$ck][0])
                    break
                }
            }
            # Fallback: token overlap on Norm-CustName
            if (-not $cur2prior[$cc]) {
                $cNorm  = Norm-CustName $cc
                $cCamel = Norm-CustName ($cc -replace '([a-z]{2,})([A-Z])', '$1 $2')
                foreach ($pc in $priorCusts) {
                    if ($matchedPrior.Contains($pc)) { continue }
                    $pNorm  = Norm-CustName $pc
                    $pCamel = Norm-CustName ($pc -replace '([a-z]{2,})([A-Z])', '$1 $2')
                    $sc = [math]::Max((Get-TokenOverlap $cNorm $pNorm), [math]::Max((Get-TokenOverlap $cCamel $pNorm), (Get-TokenOverlap $cCamel $pCamel)))
                    if ($sc -ge 0.80) {
                        $cur2prior[$cc] = $pc
                        [void]$matchedPrior.Add($pc)
                        break
                    }
                }
            }
        }
        Write-Output "  Prior customers matched: $($cur2prior.Count) / $($currentCusts.Count)"

        # Build priorLookup: for each current customer, sum prior quantities by product
        # Build priorLookup: for each current customer, sum prior quantities across ALL name variants
        # (prior EFile may have same customer under multiple spellings — collect them all via key overlap)
        foreach ($cc in $cur2prior.Keys) {
            $pc       = $cur2prior[$cc]
            $pcKeys   = Get-MatchKeys $pc
            # Collect all prior name variants that share any key with the matched prior name
            $pcVariants = $priorCusts | Where-Object {
                $vKeys = Get-MatchKeys $_
                $vKeys.Overlaps($pcKeys)
            }
            $priorLookup[$cc] = @{}
            # Deduplicate by Device ID per product so the same physical device billed under
            # multiple customer name spellings is only counted once.
            $seenDeviceIds = @{}   # key = "$prod|$deviceId" → already counted
            foreach ($variant in $pcVariants) {
                foreach ($r in ($EFPriorRows | Where-Object { $_.'Usage Customer' -ceq $variant })) {
                    $prod     = $r.Product
                    $qty      = [decimal]$r.Quantity
                    $deviceId = $r.'Device ID'
                    # If a Device ID exists, deduplicate across name variants
                    if ($deviceId) {
                        $dedupeKey = "$prod|$deviceId"
                        if ($seenDeviceIds[$dedupeKey]) { continue }
                        $seenDeviceIds[$dedupeKey] = $true
                    }
                    if ($priorLookup[$cc][$prod]) { $priorLookup[$cc][$prod] += $qty } else { $priorLookup[$cc][$prod] = $qty }
                }
            }
        }

        # Build priorTotals from raw file (all rows by product, including global/contract SKUs)
        foreach ($r in $EFPriorRaw) {
            $prod = $r.Product; $qty = [decimal]$r.Quantity
            if ($priorTotals[$prod]) { $priorTotals[$prod] += $qty } else { $priorTotals[$prod] = $qty }
        }
        # Build priorGlobalSku (blank-customer contract rows) for reconciliation auto-match
        foreach ($r in ($EFPriorRaw | Where-Object { -not $_.'Usage Customer' })) {
            $prod = $r.Product; $qty = [decimal]$r.Quantity
            if ($priorGlobalSku[$prod]) { $priorGlobalSku[$prod] += $qty } else { $priorGlobalSku[$prod] = $qty }
        }

        # Per-family churn detection is done inside the tab loop below
    }

    # Determine preferred family tab order; put extras at the end alphabetically
    $familyOrderPref = @('Cove (EFile)', 'Adlumin', 'Security (EFile)', 'N-central', 'Take Control', 'Cloud Commander', 'N-sight')
    $foundFamilies   = $EFRows | Select-Object -ExpandProperty Product -Unique |
                           ForEach-Object { Get-EFileFamily $_ } | Sort-Object -Unique
    $ordFamilies     = $familyOrderPref | Where-Object { $foundFamilies -contains $_ }
    $extraFamilies   = $foundFamilies   | Where-Object { $ordFamilies   -notcontains $_ } | Sort-Object
    $allFamilies     = @($ordFamilies) + @($extraFamilies)
    Write-Output "EFile families: $($allFamilies -join ', ')"

    foreach ($family in $allFamilies) {
        $famRows = $EFRows | Where-Object { (Get-EFileFamily $_.'Product') -eq $family }

        # Service map for this family only
        $famSvcNames = $famRows | Select-Object -ExpandProperty Product -Unique |
            Sort-Object -Property @{ Expression = { Get-ServiceSortKey $_ }; Ascending = $true },
                                  @{ Expression = { $_ };                    Ascending = $true }
        $famSvcMap   = [ordered]@{}
        $famColsUsed = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($svc in $famSvcNames) {
            $col  = Get-ServiceColName $svc
            $base = $col; $n = 2
            while (-not $famColsUsed.Add($col)) { $col = "${base}_$n"; $n++ }
            $famSvcMap[$svc] = @{ Col = $col; Display = Get-ServiceDisplayName $svc }
        }
        Write-Output "  [$family] $($famSvcMap.Count) products, $($famRows.Count) rows"

        # Pair each product with its '- Commitment' variant → single combined column
        # Uses normalized display matching to handle cross-prefix pairs (e.g. Adlumin base + N-able Security commit)
        $famPairs = [System.Collections.Generic.List[hashtable]]::new()
        $cmtSkip  = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($svc in $famSvcMap.Keys) {
            if ($cmtSkip.Contains($svc)) { continue }
            # If this product IS a Commitment with no base counterpart, render as commit-only pair (0 / N)
            if ($famSvcMap[$svc].Display -match '-\s*(Postpaid\s+)?Commitment\s*$') {
                $famPairs.Add(@{ Base = $null; Commit = $svc; CommitOnly = $true; IsGB = ($svc -like '*Per GB*') })
                continue
            }
            # Fast path: exact same-prefix key match (works for Cove)
            $commitSvc = $null
            if ($famSvcMap.Contains("$svc - Commitment")) { $commitSvc = "$svc - Commitment" }
            elseif ($famSvcMap.Contains("$svc - Postpaid Commitment")) { $commitSvc = "$svc - Postpaid Commitment" }
            # Fallback: normalize display names — strips vendor prefix (e.g. 'N-able ') and '- Commitment' suffix
            if (-not $commitSvc) {
                $baseNorm = $famSvcMap[$svc].Display -replace '(?i)^N-able\s+', ''
                foreach ($other in $famSvcMap.Keys) {
                    if ($other -eq $svc -or $cmtSkip.Contains($other)) { continue }
                    if ($famSvcMap[$other].Display -notmatch '-\s*(Postpaid\s+)?Commitment\s*$') { continue }
                    $otherNorm = $famSvcMap[$other].Display -replace '(?i)^N-able\s+', '' -replace '\s*-\s*(Postpaid\s+)?Commitment\s*$', ''
                    if ($baseNorm -eq $otherNorm) { $commitSvc = $other; break }
                }
            }
            if ($commitSvc) {
                $famPairs.Add(@{ Base = $svc; Commit = $commitSvc; IsGB = ($svc -like '*Per GB*') })
                $null = $cmtSkip.Add($commitSvc)
            } else {
                $famPairs.Add(@{ Base = $svc; Commit = $null; IsGB = ($svc -like '*Per GB*') })
            }
        }
        $padCols = [Math]::Max(0, 4 - $famPairs.Count)

        # Period label for the tab button subtitle
        $famMethods  = $famRows | Select-Object -ExpandProperty 'Rating Method' -Unique | Sort-Object
        $periodLabel = if     ($famMethods -contains 'Usage' -and $famMethods -contains 'Subscription') { 'Usage + Subscription' }
                       elseif ($famMethods -contains 'Usage')                                          { 'Usage' }
                       elseif ($famMethods -contains 'Subscription')                                   { 'Subscription' }
                       else                                                                            { ($famMethods -join ', ') }

        # Build summary: Total (all rows) then one EndCustomer row per Usage Customer
        $famSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
        $famSummary.Add((New-EFSummaryRow $family 'Total' 0 $family $famRows $famSvcMap))

        $custRows = $famRows | Where-Object { $_.'Usage Customer' -and $_.'Usage Customer'.Trim() -ne '' }
        foreach ($ecGroup in ($custRows | Group-Object 'Usage Customer' | Sort-Object Name)) {
            $ecName = if ($ecGroup.Name) { $ecGroup.Name } else { '(No Customer)' }
            $famSummary.Add((New-EFSummaryRow $family 'EndCustomer' 1 $ecName $ecGroup.Group $famSvcMap))
        }

        # Collect customer names for cross-ref
        $EFileCustomers[$family] = @($custRows | Select-Object -ExpandProperty 'Usage Customer' -Unique |
                                      Where-Object { $_ } | Sort-Object)

        # Pre-compute famChurned + rollupPrior BEFORE the rendering loop so the Total row delta
        # is always the exact rollup of all visible child rows (active + churned), never diverging.
        $famChurned  = [System.Collections.Generic.List[hashtable]]::new()
        $rollupPrior = @{}   # rollupPrior[svc] = sum(active matched prior) + sum(churned prior)
        if ($hasPrior) {
            $curFamCusts = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($ec in $EFileCustomers[$family]) { [void]$curFamCusts.Add($ec) }
            $curFamKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($ec in $curFamCusts) { foreach ($k in (Get-MatchKeys $ec)) { [void]$curFamKeys.Add($k) } }
            $priFamCusts = @($EFPriorRows | Where-Object { $_.'Usage Customer' -and (Get-EFileFamily $_.'Product') -eq $family } |
                               Select-Object -ExpandProperty 'Usage Customer' -Unique)
            foreach ($pc in $priFamCusts) {
                $pKeys = Get-MatchKeys $pc; $found = $false
                foreach ($pk in $pKeys) { if ($curFamKeys.Contains($pk)) { $found = $true; break } }
                if (-not $found) {
                    $pcRows = @($EFPriorRows | Where-Object { $_.'Usage Customer' -eq $pc -and (Get-EFileFamily $_.'Product') -eq $family })
                    $famChurned.Add(@{ Name = $pc; Rows = $pcRows })
                }
            }
            # Rollup prior per product: active (matched) + churned
            foreach ($svc in $famSvcMap.Keys) {
                $actSum = [decimal]0
                foreach ($r in ($famSummary | Where-Object { $_.Level -eq 'EndCustomer' })) {
                    if ($priorLookup[$r.Name] -and $priorLookup[$r.Name][$svc]) { $actSum += [decimal]$priorLookup[$r.Name][$svc] }
                }
                $churnSum = [decimal]0
                foreach ($ch in $famChurned) {
                    $cv = ($ch.Rows | Where-Object { $_.Product -eq $svc } | ForEach-Object { [decimal]$_.'Quantity' } | Measure-Object -Sum).Sum
                    if ($cv) { $churnSum += $cv }
                }
                $rollupPrior[$svc] = $actSum + $churnSum
            }
        }

        # ── Build HTML rows ──────────────────────────────────────────────
        $rowId       = 0
        $depthParent = @{}
        $htmlRows    = [System.Text.StringBuilder]::new()
        $pfx         = "t$tabIdx"

        for ($i = 0; $i -lt $famSummary.Count; $i++) {
            $row   = $famSummary[$i]
            $depth = [int]$row.Indent
            $rid   = "${pfx}_$(if ($depth -eq 0){'p'}else{'r'})$rowId"
            $rowId++
            $depthParent[$depth] = $rid

            $isParent     = ($i + 1 -lt $famSummary.Count) -and ([int]$famSummary[$i+1].Indent -gt $depth)
            $parentAttr   = if ($depth -eq 0) { '' } else { "data-parent='$($depthParent[$depth-1])'" }
            $displayStyle = if ($depth -eq 0) { '' } else { "style='display:none'" }
            $collClass    = if ($isParent)    { 'collapsible' } else { '' }
            $clickAttr    = if ($isParent)    { "onclick='toggle(this)'" } else { '' }
            $togSpan      = if ($isParent)    { "<span class='tog'>&#9658;</span> " } else { "<span class='tog-leaf'></span> " }

            $cls = switch ($row.Level) {
                'Total'       { 'row-total' }
                'EndCustomer' { 'row-endcustomer' }
                default       { 'row-other' }
            }
            $levelBadge = switch ($row.Level) {
                'Total'       { "<span class='badge badge-contract'>$([System.Web.HttpUtility]::HtmlEncode($family))</span>" }
                'EndCustomer' { '<span class="badge badge-endcustomer">End Customer</span>' }
                default       { '' }
            }

            $null = $htmlRows.AppendLine("<tr class='$cls $collClass' data-id='$rid' $parentAttr data-depth='$depth' data-open='0' $displayStyle $clickAttr>")
            # Determine if this is a new customer (not in prior period)
            $isNewCust  = $hasPrior -and ($row.Level -eq 'EndCustomer') -and (-not $cur2prior[$row.Name])
            $newBadge   = if ($isNewCust) { " <span class='badge-new'>NEW</span>" } else { '' }
            $matchCheck = if ($hasPrior -and $row.Level -eq 'EndCustomer' -and $cur2prior[$row.Name]) {
                              $priName = [System.Web.HttpUtility]::HtmlEncode($cur2prior[$row.Name])
                              " <span class='recon-check' title='Prior matched: $priName'>&#10003;</span>"
                          } else { '' }
            $nameWords = $row.Name -split '\s+', 2
            $nameHtml  = if ($nameWords.Count -gt 1) {
                "$([System.Web.HttpUtility]::HtmlEncode($nameWords[0])) <span class='priv-rest'>$([System.Web.HttpUtility]::HtmlEncode($nameWords[1]))</span>"
            } else { [System.Web.HttpUtility]::HtmlEncode($row.Name) }
            $null = $htmlRows.AppendLine("  <td class='name-col depth$depth'>$togSpan$nameHtml $levelBadge$newBadge$matchCheck</td>")
            foreach ($pair in $famPairs) {
                $isGB = $pair.IsGB
                # Compute prior value for delta
                $priorVal = [decimal]0
                if ($hasPrior) {
                    $baseSvc = if ($pair.Base) { $pair.Base } else { $pair.Commit }
                    if ($row.Level -eq 'Total') {
                        # rollupPrior = named-customer prior (active + churned); add global SKU prior for families like N-central
                        $priorVal  = if ($rollupPrior[$baseSvc])    { [decimal]$rollupPrior[$baseSvc]    } else { [decimal]0 }
                        $priorVal += if ($priorGlobalSku[$baseSvc]) { [decimal]$priorGlobalSku[$baseSvc] } else { [decimal]0 }
                    } elseif ($row.Level -eq 'EndCustomer' -and $priorLookup[$row.Name]) {
                        $priorVal = if ($priorLookup[$row.Name][$baseSvc]) { [decimal]$priorLookup[$row.Name][$baseSvc] } else { 0 }
                    }
                    if ($isGB) { $priorVal = [math]::Round($priorVal, 1) }
                }
                if ($pair.CommitOnly) {
                    $cmtCol = $famSvcMap[$pair.Commit].Col
                    $cmtVal = $row.$cmtCol
                    $cHtml  = if ($cmtVal -and [decimal]$cmtVal -ne 0) {
                        $cFmt = if ($isGB) { Format-GB $cmtVal } else { Format-Num ([decimal]$cmtVal) }
                        " <span class='commit-qty' title='Commitment (prepay): $cmtVal'>/ $cFmt</span>"
                    } else { '' }
                    $null = $htmlRows.AppendLine("  <td><span class='zero'>0</span>$cHtml</td>")
                } elseif ($pair.Commit) {
                    $baseCol = $famSvcMap[$pair.Base].Col
                    $baseVal = $row.$baseCol
                    $cmtCol = $famSvcMap[$pair.Commit].Col
                    $cmtVal = $row.$cmtCol
                    $bFmt   = if ($isGB) { Format-GB $baseVal } else { Format-Num ([decimal]$baseVal) }
                    $dHtml  = if ($hasPrior) { Format-Delta ([decimal]$baseVal) $priorVal $isGB } else { '' }
                    $cHtml  = if ($cmtVal -and [decimal]$cmtVal -ne 0) {
                        $cFmt = if ($isGB) { Format-GB $cmtVal } else { Format-Num ([decimal]$cmtVal) }
                        " <span class='commit-qty' title='Commitment (prepay): $cmtVal'>/ $cFmt</span>"
                    } else { '' }
                    $null = $htmlRows.AppendLine("  <td>$bFmt$dHtml$cHtml</td>")
                } else {
                    $baseCol = $famSvcMap[$pair.Base].Col
                    $baseVal = $row.$baseCol
                    $valHtml = if ($isGB) { Format-GB $baseVal } else { Format-Num ([decimal]$baseVal) }
                    $dHtml   = if ($hasPrior) { Format-Delta ([decimal]$baseVal) $priorVal $isGB } else { '' }
                    $null = $htmlRows.AppendLine("  <td>$valHtml$dHtml</td>")
                }
            }
            for ($p = 0; $p -lt $padCols; $p++) { $null = $htmlRows.AppendLine('  <td></td>') }
            $null = $htmlRows.AppendLine('</tr>')
        }

        # ── Churned customer rows (famChurned pre-computed above) ──────────────────────────────────
        $reconHtml = ''
        if ($hasPrior) {
            foreach ($churned in ($famChurned | Sort-Object { $_.Name })) {
                $rid   = "${pfx}_r$rowId"; $rowId++
                $parentId = $depthParent[0]
                $null = $htmlRows.AppendLine("<tr class='row-endcustomer row-churned collapsible' data-id='$rid' data-parent='$parentId' data-depth='1' data-open='0' style='display:none'>")
                $cnWords = $churned.Name -split '\s+', 2
                $cnHtml  = if ($cnWords.Count -gt 1) {
                    "$([System.Web.HttpUtility]::HtmlEncode($cnWords[0])) <span class='priv-rest'>$([System.Web.HttpUtility]::HtmlEncode($cnWords[1]))</span>"
                } else { [System.Web.HttpUtility]::HtmlEncode($churned.Name) }
                $null = $htmlRows.AppendLine("  <td class='name-col depth1'><span class='tog-leaf'></span> <span class='churned-name'>$cnHtml</span> <span class='badge-churned'>CHURNED</span></td>")
                foreach ($pair in $famPairs) {
                    $isGB = $pair.IsGB
                    $svc  = if ($pair.Base) { $pair.Base } else { $pair.Commit }
                    $pVal = ($churned.Rows | Where-Object { $_.Product -eq $svc } | ForEach-Object { [decimal]$_.'Quantity' } | Measure-Object -Sum).Sum
                    if (-not $pVal) { $pVal = 0 }
                    if ($isGB) { $pVal = [math]::Round($pVal, 1) }
                    $fVal = if ($pVal -eq 0) { '<span class="zero">0</span>' } elseif ($isGB) { Format-GB $pVal } else { Format-Num $pVal }
                    $dHtml = if ($pVal -gt 0) { " <span class='delta-down' title='Churned'>&#9660;$(if($isGB){'{0:N1}' -f $pVal}else{'{0:N0}' -f $pVal})</span>" } else { '' }
                    $null = $htmlRows.AppendLine("  <td class='churned-val'>$fVal$dHtml</td>")
                }
                for ($p = 0; $p -lt $padCols; $p++) { $null = $htmlRows.AppendLine('  <td></td>') }
                $null = $htmlRows.AppendLine('</tr>')
            }

            # ── Reconciliation block ──────────────────────────────────────────────────────────────────
            # The Total row delta = rollupPrior (active+churned), so it is always self-consistent.
            # This block checks whether any prior quantity was unmatched (not reflected in any visible row).
            $reconRows    = [System.Text.StringBuilder]::new()
            $allReconPass = $true
            foreach ($pair in $famPairs) {
                if ($pair.CommitOnly) { continue }
                $baseSvc      = $pair.Base
                $isGB         = $pair.IsGB
                $colKey       = $famSvcMap[$baseSvc].Col
                $fmt          = if ($isGB) { 'N1' } else { 'N0' }
                $gsqPrior     = if ($priorGlobalSku[$baseSvc]) { [decimal]$priorGlobalSku[$baseSvc] } else { [decimal]0 }
                $curTotal     = [decimal]($famSummary[0].$colKey)  # Total row already includes all rows incl. global SKUs
                $priFileTotal = if ($priorTotals[$baseSvc]) { [decimal]$priorTotals[$baseSvc] } else { [decimal]0 }
                $acctForBase  = if ($rollupPrior[$baseSvc]) { [decimal]$rollupPrior[$baseSvc] } else { [decimal]0 }
                $acctFor      = $acctForBase + $gsqPrior
                $unmatched    = $priFileTotal - $acctFor
                if ($isGB) { $unmatched = [math]::Round($unmatched, 1) }
                $pass         = ($unmatched -eq 0)
                $allReconPass = $allReconPass -and $pass
                $checkIcon    = if ($pass) { "<span class='recon-ok'>&#10003;</span>" }
                               else        { "<span class='recon-fail'>&#10007; $($unmatched.ToString($fmt)) unmatched</span>" }
                $netChange    = $curTotal - $acctFor
                $netFmt       = if ($netChange -eq 0)   { '&mdash;' }
                               elseif ($netChange -gt 0) { "<span class='recon-delta-up'>+$($netChange.ToString($fmt))</span>" }
                               else                      { "<span class='recon-delta-dn'>$($netChange.ToString($fmt))</span>" }
                $dispName = $famSvcMap[$baseSvc].Display
                $null = $reconRows.AppendLine("<tr><td>$([System.Web.HttpUtility]::HtmlEncode($dispName))</td>")
                $null = $reconRows.AppendLine("  <td>$($priFileTotal.ToString($fmt))</td>")
                $null = $reconRows.AppendLine("  <td>$($curTotal.ToString($fmt))</td>")
                $null = $reconRows.AppendLine("  <td>$($acctFor.ToString($fmt))</td>")
                $null = $reconRows.AppendLine("  <td>$checkIcon</td>")
                $null = $reconRows.AppendLine("  <td>$netFmt</td></tr>")
            }
            if ($reconRows.Length -gt 0) {
                $hdr = if ($allReconPass) { "Reconciliation <span class='recon-ok'>&#10003; all prior quantities accounted for</span>" }
                        else              { "Reconciliation <span class='recon-fail'>&#10007; some prior quantities unmatched</span>" }
                $curHdr   = if ($EFPeriod)    { "Current&nbsp;<span class='recon-period'>($EFPeriod)</span>" }   else { 'Current' }
                $priHdr   = if ($priorPeriod) { "Prior&nbsp;(file)&nbsp;<span class='recon-period'>($priorPeriod)</span>" } else { 'Prior&nbsp;(file)' }
                $reconHtml = "<div class='recon-box'><div class='recon-title recon-toggle' onclick='toggleRecon(this)' title='Click to expand/collapse'>$hdr <button class='recon-btn' tabindex='-1'>RECONCILE&nbsp;&#9660;&nbsp;COLLAPSE</button></div><div class='recon-body'><table class='recon-table'><thead><tr><th>Product</th><th>$priHdr</th><th>$curHdr</th><th>Accounted&nbsp;for&#185;</th><th>Check</th><th>Net&nbsp;change</th></tr></thead><tbody>$($reconRows.ToString())</tbody></table><p class='recon-note'>&#185; Rollup of all visible rows: active customers&rsquo; matched prior + churned customers&rsquo; prior. Any gap means prior customers that could not be matched to current or churned.</p></div></div>"
            }
        }

        # Tab nav button
        $efActiveClass = if ($tabIdx -eq 0) { 'active' } else { '' }
        $null = $tabNavHtml.AppendLine("<button class='tab-btn $efActiveClass' onclick='showTab(""tab$tabIdx"",this)'>$([System.Web.HttpUtility]::HtmlEncode($family))<br><span style='font-size:10px;font-weight:400'>$([System.Web.HttpUtility]::HtmlEncode($periodLabel))</span></button>")

        # Column headers (one <th> per pair)
        $famThHtml = ($famPairs | ForEach-Object {
            $d    = if ($_.CommitOnly) { $famSvcMap[$_.Commit].Display -replace '\s*-\s*(Postpaid\s+)?Commitment\s*$','' }
                    else               { $famSvcMap[$_.Base].Display }
            $isGB = $_.IsGB
            $lbl  = Split-DisplayLabel $d
            $cBadge = if ($_.Commit) { "<br><span class='commit-badge'>+ Commitment</span>" } else { '' }
            $wAttr  = if ($isGB) { " style='min-width:110px'" } else { " style='min-width:120px'" }
            "      <th$wAttr>$lbl$cBadge</th>"
        }) -join "`n"
        if ($padCols -gt 0) { $famThHtml += "`n" + ((1..$padCols | ForEach-Object { "      <th style='min-width:120px'></th>" }) -join "`n") }

        $efPanelStyle = if ($tabIdx -eq 0) { '' } else { "style='display:none'" }
        $null = $tabPanelHtml.AppendLine("<div id='tab$tabIdx' class='tab-panel' $efPanelStyle>")
        $null = $tabPanelHtml.AppendLine("<div class='tab-actions'><button class='btn' onclick='expandAll(""tab$tabIdx"")'>Expand All</button> <button class='btn' onclick='collapseAll(""tab$tabIdx"")'>Collapse All</button> <input type='text' class='search-box' id='search-tab$tabIdx' placeholder='Filter rows...' oninput='filterRows(""tab$tabIdx"",this)'> <button class='btn btn-clear' onclick='clearFilter(""tab$tabIdx"")' title='Clear'>&#x2715;</button> <label class='preserve-lbl'><input type='checkbox' class='preserve-cb' onchange='syncPreserve(this)'> Preserve filter</label></div>")
        $null = $tabPanelHtml.AppendLine("<div class='tab-scroll'><table><thead><tr>")
        $null = $tabPanelHtml.AppendLine("  <th>Customer</th>")
        $null = $tabPanelHtml.AppendLine($famThHtml)
        $null = $tabPanelHtml.AppendLine("</tr></thead><tbody>")
        $null = $tabPanelHtml.AppendLine($htmlRows.ToString())
        $null = $tabPanelHtml.AppendLine("</tbody></table></div>")
        if ($reconHtml) { $null = $tabPanelHtml.AppendLine($reconHtml) }
        $null = $tabPanelHtml.AppendLine("<p class='efile-footnote'>&#9432;&nbsp; Commitment figures are contract-level totals &mdash; not broken down by customer.</p>")
        $null = $tabPanelHtml.AppendLine("</div>")

        $tabIdx++
    }
} else {
    if ($EFilePath) { Write-Warning "EFile not found: $EFilePath" }
}

# ── Customer Cross-Reference Tab ─────────────────────────────────────────────
# Collect unique EndCustomer names per product (Cove/EDR)
$xrefProdNames = [ordered]@{}
foreach ($pg in $prodGroups) {
    $xrefProdNames[$pg.Name] = @($Summary |
        Where-Object { $_.ProductName -eq $pg.Name -and $_.Level -eq 'EndCustomer' } |
        Select-Object -ExpandProperty Name |
        Where-Object { $_ -and $_.Trim() -ne '' } |
        Sort-Object -Unique)
}
# Add EFile families to cross-ref (only families that have customer-level data)
foreach ($fam in $EFileCustomers.Keys) {
    if ($EFileCustomers[$fam].Count -gt 0) {
        $xrefProdNames[$fam] = $EFileCustomers[$fam]
    }
}
$xrefNumProds = $xrefProdNames.Keys.Count

# Build match groups — Tier 1: multi-key match (norm/camel/strip/plural)  Tier 2: token overlap >= 80%
$xrefGroups = [System.Collections.Generic.List[hashtable]]::new()
foreach ($prod in $xrefProdNames.Keys) {
    foreach ($name in $xrefProdNames[$prod]) {
        $myKeys     = Get-MatchKeys $name
        $normName   = Norm-CustName $name
        $camelNorm  = Norm-CustName ($name -replace '([a-z]{2,})([A-Z])', '$1 $2')
        $matchedGrp = $null
        $matchType  = 'none'
        foreach ($grp in $xrefGroups) {
            if ($grp.Names[$prod]) { continue }
            foreach ($existName in ($grp.Names.Values | Where-Object { $_ })) {
                $existKeys = Get-MatchKeys $existName
                if ($myKeys.Overlaps($existKeys))       { $matchedGrp = $grp; $matchType = 'norm';  break }
                $existNorm  = Norm-CustName $existName
                $existCamel = Norm-CustName ($existName -replace '([a-z]{2,})([A-Z])', '$1 $2')
                # Token overlap on both plain norm and CamelCase-split norm — take the best
                $sc = [math]::Max(
                    (Get-TokenOverlap $normName $existNorm),
                    [math]::Max(
                        (Get-TokenOverlap $camelNorm $existNorm),
                        (Get-TokenOverlap $camelNorm $existCamel)
                    )
                )
                if ($sc -ge 0.80)                      { $matchedGrp = $grp; $matchType = 'fuzzy'; break }
            }
            if ($matchedGrp) { break }
        }
        if ($matchedGrp) {
            $matchedGrp.Names[$prod] = $name
            if ($matchType -eq 'fuzzy') { $matchedGrp.HasFuzzy = $true }
        } else {
            $newGrp = [ordered]@{ Canonical = $name; HasFuzzy = $false; Names = [ordered]@{} }
            foreach ($p in $xrefProdNames.Keys) { $newGrp.Names[$p] = $null }
            $newGrp.Names[$prod] = $name
            $xrefGroups.Add($newGrp)
        }
    }
}

# Compute status per group
foreach ($grp in $xrefGroups) {
    $present     = ($grp.Names.Values | Where-Object { $_ }).Count
    $grp.Status  = if     ($grp.HasFuzzy -and $present -lt $xrefNumProds) { 'fuzzy-partial' }
                   elseif ($grp.HasFuzzy)                                  { 'fuzzy' }
                   elseif ($present -lt $xrefNumProds)                     { 'partial' }
                   else                                                    { 'matched' }
}

$xrefAll     = ($xrefGroups | Where-Object { $_.Status -eq 'matched' }).Count
$xrefFuzzy   = ($xrefGroups | Where-Object { $_.Status -eq 'fuzzy' -or $_.Status -eq 'fuzzy-partial' }).Count
$xrefPartial = ($xrefGroups | Where-Object { $_.Status -eq 'partial' -or $_.Status -eq 'fuzzy-partial' }).Count
$xrefTotal   = $xrefGroups.Count
Write-Output "Cross-ref groups: $xrefTotal  (All: $xrefAll  Fuzzy: $xrefFuzzy  Partial: $xrefPartial)"

# Product column header HTML
$xrefProdThHtml = ($xrefProdNames.Keys | ForEach-Object {
    $parts = $_ -split '\s*\|\s*'
    if ($parts.Count -ge 2) {
        $midP = [int][math]::Round($parts.Count / 2)
        $l1   = [System.Web.HttpUtility]::HtmlEncode(($parts[0..($midP-1)] -join ' | '))
        $l2   = [System.Web.HttpUtility]::HtmlEncode(($parts[$midP..($parts.Count-1)] -join ' | '))
        "      <th class='xref-prod-th'>$l1<br><span style='font-weight:400;font-size:11px'>$l2</span></th>"
    } else {
        "      <th class='xref-prod-th'>$([System.Web.HttpUtility]::HtmlEncode($_))</th>"
    }
}) -join "`n"

# Table rows — alphabetical by canonical name
$xrefRows   = [System.Text.StringBuilder]::new()
$xrefRowNum = 0
foreach ($grp in ($xrefGroups | Sort-Object { $_.Canonical })) {
    $xrefRowNum++
    $pillClass = switch ($grp.Status) {
        'matched'       { 'xref-pill-green'   }
        'fuzzy'         { 'xref-pill-amber'   }
        'fuzzy-partial' { 'xref-pill-amber'   }
        default         { 'xref-pill-partial' }
    }
    $pillText = switch ($grp.Status) {
        'matched'       { 'All Products'    }
        'fuzzy'         { 'Fuzzy Match'     }
        'fuzzy-partial' { 'Fuzzy / Partial' }
        default         { 'Partial'         }
    }
    $null = $xrefRows.Append("<tr data-xstatus='$($grp.Status)'>")
    $null = $xrefRows.Append("<td class='xref-idx'>$xrefRowNum</td>")
    $null = $xrefRows.Append("<td><span class='xref-pill $pillClass'>$([System.Web.HttpUtility]::HtmlEncode($pillText))</span></td>")
    $null = $xrefRows.Append("<td class='xref-canonical'>$([System.Web.HttpUtility]::HtmlEncode($grp.Canonical))</td>")
    foreach ($prod in $xrefProdNames.Keys) {
        $n = $grp.Names[$prod]
        if (-not $n) {
            $null = $xrefRows.Append("<td class='xref-absent'><span title='Not in this product'>&mdash;</span></td>")
        } else {
            $normN   = Norm-CustName $n
            $normCan = Norm-CustName $grp.Canonical
            if ($n -eq $grp.Canonical) {
                $null = $xrefRows.Append("<td class='xref-match-exact'>$([System.Web.HttpUtility]::HtmlEncode($n))</td>")
            } elseif ($normN -eq $normCan) {
                $null = $xrefRows.Append("<td class='xref-match-norm'>$([System.Web.HttpUtility]::HtmlEncode($n)) <span class='xref-conf'>(norm)</span></td>")
            } else {
                $sc2 = [math]::Round((Get-TokenOverlap $normN $normCan) * 100)
                $null = $xrefRows.Append("<td class='xref-match-fuzzy'>$([System.Web.HttpUtility]::HtmlEncode($n)) <span class='xref-conf'>${sc2}%</span></td>")
            }
        }
    }
    $null = $xrefRows.AppendLine("</tr>")
}

# ── Unit Prices Tab ───────────────────────────────────────────────────────────
if ($EFilePath -and (Test-Path $EFilePath)) {
    # Collect unique product → rate from current file
    $priceMap = [ordered]@{}   # product → @{ Rate; UOM; Method; PriorRate; PriorUOM }
    foreach ($r in ($EFRaw | Sort-Object Product)) {
        $prod = $r.Product
        if (-not $priceMap.Contains($prod)) {
            $priceMap[$prod] = @{ Rate = $r.Rate; UOM = $r.UOM; Method = $r.'Rating Method'; PriorRate = ''; PriorUOM = ''; Period = "$($r.'Period From') thru $($r.'Period To')" }
        }
    }
    # Overlay prior rates
    if ($hasPrior) {
        foreach ($r in ($EFPriorRaw | Sort-Object Product)) {
            $prod = $r.Product
            if ($priceMap.Contains($prod)) {
                if (-not $priceMap[$prod].PriorRate) { $priceMap[$prod].PriorRate = $r.Rate; $priceMap[$prod].PriorUOM = $r.UOM }
            } else {
                $priceMap[$prod] = @{ Rate = ''; UOM = $r.UOM; Method = $r.'Rating Method'; PriorRate = $r.Rate; PriorUOM = $r.UOM; Period = "$($r.'Period From') thru $($r.'Period To')" }
            }
        }
    }

    # Build tab
    $priceSubtitle = if ($hasPrior) { "$priorPeriod vs $EFPeriod" } else { $EFPeriod }
    $null = $tabNavHtml.AppendLine("<button class='tab-btn' onclick='showTab(""tab$tabIdx"",this)'>&#128176; Unit Prices<br><span style='font-size:10px;font-weight:400'>$([System.Web.HttpUtility]::HtmlEncode($priceSubtitle))</span></button>")
    $null = $tabPanelHtml.AppendLine("<div id='tab$tabIdx' class='tab-panel price-panel' style='display:none'>")
    $null = $tabPanelHtml.AppendLine("<div class='tab-actions'><input type='text' class='search-box' id='search-price' placeholder='Filter products...' oninput='filterPrice(this)'> <button class='btn btn-clear' onclick='this.previousElementSibling.value="""";filterPrice(this.previousElementSibling)' title='Clear'>&#x2715;</button></div>")
    $null = $tabPanelHtml.AppendLine("<div class='price-scroll'><table class='price-tbl'><thead><tr>")
    $null = $tabPanelHtml.AppendLine("  <th class='price-prod-th'>Product</th>")
    $null = $tabPanelHtml.AppendLine("  <th class='price-ctr-th price-period-th'>Period</th>")
    $null = $tabPanelHtml.AppendLine("  <th class='price-ctr-th'>Method</th>")
    $null = $tabPanelHtml.AppendLine("  <th class='price-ctr-th'>UOM</th>")
    $curRateHdr  = if ($EFPeriod)    { "Rate&nbsp;<span class='recon-period'>($EFPeriod)</span>" }   else { 'Rate' }
    $priRateHdr  = if ($hasPrior -and $priorPeriod) { "Prior Rate&nbsp;<span class='recon-period'>($priorPeriod)</span>" } else { 'Prior Rate' }
    if ($hasPrior) { $null = $tabPanelHtml.AppendLine("  <th class='price-num-th'>$priRateHdr</th><th class='price-num-th'>$curRateHdr</th><th class='price-num-th'>Change</th>") } else { $null = $tabPanelHtml.AppendLine("  <th class='price-num-th'>$curRateHdr</th>") }
    $null = $tabPanelHtml.AppendLine("</tr></thead><tbody>")

    # Sort by normalized display name so renamed items (e.g. Adlumin MDR) group correctly,
    # with Subscription/Commitment rows immediately above their Usage sibling
    $priceSorted = $priceMap.Keys | Sort-Object {
        $base = $_ -replace '\s*-\s*(Postpaid\s+)?Commitment\s*$', '' -replace '\s*-\s*(Postpaid\s+)?Freemium\s*$', ''
        $base = $base -replace '^N-able Security \| N-able (Advanced MDR)', 'Adlumin | $1' `
                      -replace 'Endpoint Detection and Response', 'EDR' `
                      -replace 'Virtual Machine Server', 'Virtual Server' `
                      -replace 'Cove Continuity', 'Continuity'
        $subFirst = if ($_ -match 'Commitment|Freemium') { '0' } else { '1' }
        "$base`t$subFirst"
    }
    foreach ($prod in $priceSorted) {
        $p        = $priceMap[$prod]
        $method     = [System.Web.HttpUtility]::HtmlEncode($p.Method)
        $uom        = [System.Web.HttpUtility]::HtmlEncode($p.UOM)
        $periodEnc  = [System.Web.HttpUtility]::HtmlEncode($p.Period)
        # Normalize display name: MDR items → Adlumin branding; EDR verbose → short; VM Server → Virtual Server; Cove Continuity → Continuity
        $dispProd = $prod -replace '^N-able Security \| N-able (Advanced MDR)', 'Adlumin | $1' `
                         -replace 'Endpoint Detection and Response', 'EDR' `
                         -replace 'Virtual Machine Server', 'Virtual Server' `
                         -replace 'Cove Continuity', 'Continuity'
        $prodEnc  = [System.Web.HttpUtility]::HtmlEncode($dispProd)
        $rateDisp = if ($p.Rate -ne '') { $p.Rate } else { '<span class="price-absent">&mdash;</span>' }
        $rowClass = 'price-row'

        if ($hasPrior) {
            $priorDisp = if ($p.PriorRate -ne '') { $p.PriorRate } else { '<span class="price-absent">&mdash;</span>' }
            $changeFmt = '&mdash;'
            if ($p.Rate -ne '' -and $p.PriorRate -ne '') {
                $diff = [decimal]$p.Rate - [decimal]$p.PriorRate
                if ($diff -gt 0)      { $changeFmt = "<span class='price-up'>&#9650; $($diff.ToString('G6'))</span>" }
                elseif ($diff -lt 0)  { $changeFmt = "<span class='price-dn'>&#9660; $([math]::Abs($diff).ToString('G6'))</span>" }
            } elseif ($p.Rate -ne '' -and $p.PriorRate -eq '') {
                $changeFmt = "<span class='price-new'>NEW</span>"
                $rowClass += ' price-row-new'
            } elseif ($p.Rate -eq '' -and $p.PriorRate -ne '') {
                $changeFmt = "<span class='price-gone'>REMOVED</span>"
                $rowClass += ' price-row-gone'
            }
            $null = $tabPanelHtml.AppendLine("<tr class='$rowClass'><td class='price-prod'>$prodEnc</td><td class='price-ctr price-period-td'>$periodEnc</td><td class='price-ctr'>$method</td><td class='price-ctr'>$uom</td><td class='price-val'>$priorDisp</td><td class='price-val'>$rateDisp</td><td class='price-chg'>$changeFmt</td></tr>")
        } else {
            $null = $tabPanelHtml.AppendLine("<tr class='$rowClass'><td class='price-prod'>$prodEnc</td><td class='price-ctr price-period-td'>$periodEnc</td><td class='price-ctr'>$method</td><td class='price-ctr'>$uom</td><td class='price-val'>$rateDisp</td></tr>")
        }
    }

    $null = $tabPanelHtml.AppendLine("</tbody></table></div>")
    $null = $tabPanelHtml.AppendLine("<div class='price-footnote'>&#42; Line items with usage below commitment values may show 0 or null per-unit rates.</div>")
    $null = $tabPanelHtml.AppendLine("</div>")
    $tabIdx++
}

# Cross-ref nav button (always last tab)
$null = $tabNavHtml.AppendLine("<button class='tab-btn' onclick='showTab(""tab$tabIdx"",this)'>&#128269; Customer<br>Cross-Reference</button>")

# Cross-ref tab panel
$null = $tabPanelHtml.AppendLine("<div id='tab$tabIdx' class='tab-panel xref-panel' style='display:none'>")
$null = $tabPanelHtml.AppendLine("<div class='tab-actions xref-actions'>")
$null = $tabPanelHtml.AppendLine("  <input type='text' class='search-box' id='search-xref' placeholder='Filter customer name...' oninput='filterXref(this)'>")
$null = $tabPanelHtml.AppendLine("  <button class='btn btn-clear' onclick='clearXref()' title='Clear'>&#x2715;</button>")
$null = $tabPanelHtml.AppendLine("  <label class='xref-chk'><input type='checkbox' id='xchk-fuzzy' onchange='filterXref()'>&nbsp;Fuzzy only</label>")
$null = $tabPanelHtml.AppendLine("  <label class='xref-chk'><input type='checkbox' id='xchk-partial' onchange='filterXref()'>&nbsp;Partial only</label>")
$null = $tabPanelHtml.AppendLine("  <span class='xref-stats-line'>")
$null = $tabPanelHtml.AppendLine("    <span class='xsi-box xsi-green'>$xrefAll All</span>")
$null = $tabPanelHtml.AppendLine("    <span class='xsi-box xsi-amber'>$xrefFuzzy Fuzzy</span>")
$null = $tabPanelHtml.AppendLine("    <span class='xsi-box xsi-partial'>$xrefPartial Partial</span>")
$null = $tabPanelHtml.AppendLine("    <span class='xsi-box'>$xrefTotal Total</span>")
$null = $tabPanelHtml.AppendLine("  </span>")
$null = $tabPanelHtml.AppendLine("</div>")
$null = $tabPanelHtml.AppendLine("<div class='xref-scroll'><table class='xref-tbl'><thead><tr>")
$null = $tabPanelHtml.AppendLine("  <th class='xref-idx-th'>#</th>")
$null = $tabPanelHtml.AppendLine("  <th class='xref-status-th'>Match Status</th>")
$null = $tabPanelHtml.AppendLine("  <th class='xref-canonical-th'>Canonical Name</th>")
$null = $tabPanelHtml.AppendLine($xrefProdThHtml)
$null = $tabPanelHtml.AppendLine("</tr></thead><tbody>")
$null = $tabPanelHtml.AppendLine($xrefRows.ToString())
$null = $tabPanelHtml.AppendLine("</tbody></table></div></div>")

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>$([System.Web.HttpUtility]::HtmlEncode($Account)) | $Period | Usage Summary</title>
<style>
  body { font-family:'Segoe UI',Arial,sans-serif; font-size:13px; background:#f0f2f5; margin:0; padding:0;
         display:flex; flex-direction:column; height:100vh; overflow:hidden; box-sizing:border-box; }
  .page-header { padding:10px 16px 0 16px; flex-shrink:0; }
  h2   { color:#1a3a5c; margin:0 0 2px 0; font-size:18px; }
  p.sub{ color:#666; margin:0 0 6px 0; font-size:12px; }
  /* Tabs */
  .tab-nav  { display:flex; gap:4px; flex-wrap:wrap; flex-shrink:0; padding:0 16px; margin-bottom:0; }
  .tab-btn  { padding:6px 14px; font-size:12px; font-family:inherit; cursor:pointer;
              border:1px solid #b0c4d8; border-bottom:none; border-radius:6px 6px 0 0;
              background:#d6e8f5; color:#1a3a5c; font-weight:500;
              white-space:normal; text-align:center; line-height:1.3; }
  .tab-btn:hover  { background:#c0d8ee; }
  .tab-btn.active { background:#1a5276; color:#fff; border-color:#1a5276; }
  .tab-wrap { flex:1; min-height:0; padding:0 16px 12px 16px; display:flex; flex-direction:column; }
  .tab-panel { background:#fff; border:1px solid #b0c4d8; border-radius:0 6px 6px 6px;
               box-shadow:0 1px 4px rgba(0,0,0,.12); padding:12px 0 0 0;
               flex:1; min-height:0; overflow:hidden; }
  .tab-actions { padding:0 12px 10px 12px; display:flex; align-items:center; gap:6px; flex-wrap:wrap; }
  .search-box { padding:4px 8px; font-size:12px; border:1px solid #b0c4d8; border-radius:4px; font-family:inherit; width:200px; }
  .search-box:focus { outline:none; border-color:#1a5276; box-shadow:0 0 0 2px rgba(26,82,118,.15); }
  .btn-clear { padding:4px 8px; min-width:auto; }
  .preserve-lbl { font-size:11px; color:#555; cursor:pointer; display:flex; align-items:center; gap:3px; margin-left:4px; user-select:none; }
  .preserve-lbl input { margin:0; cursor:pointer; }
  /* Table */
  table { border-collapse:separate; border-spacing:0; width:100%; background:#fff; }
  thead tr { background:#1a5276; color:#fff; }
  thead th { position:sticky; top:0; z-index:2; background:#1a5276;
             padding:9px 12px; text-align:right; font-weight:600; font-size:12px;
             white-space:normal; word-break:break-word; min-width:80px;
             vertical-align:bottom; line-height:1.3; }
  thead th:first-child { text-align:left; min-width:280px; vertical-align:middle; }
  tbody tr { border-bottom:1px solid #e8ecf0; transition:background .1s; }
  tbody tr:hover { background:#eaf3fb; }
  tbody td { padding:6px 12px; text-align:right; white-space:nowrap; vertical-align:middle; }
  tbody td:first-child { text-align:left; }
  /* Row depth / type */
  .row-total td      { font-weight:700; background:#d6eaf8 !important; }
  .row-reseller td   { background:#fdfefe; }
  .row-serviceorg td { background:#fdfefe; }
  .row-site td       { font-style:italic; background:#fafafa; }
  td.depth0 { padding-left:8px !important;  font-weight:700; }
  td.depth1 { padding-left:28px !important; }
  td.depth2 { padding-left:48px !important; }
  td.depth3 { padding-left:68px !important; }
  td.depth4 { padding-left:88px !important; }
  .tog      { display:inline-block; width:14px; color:#1a5276; font-size:11px; user-select:none; }
  .tog-leaf { display:inline-block; width:14px; }
  .collapsible { cursor:pointer; }
  .zero { color:#ccc; }
  .na   { color:#aaa; }
  .btn { padding:5px 12px; font-size:12px; cursor:pointer;
         border:1px solid #1a5276; border-radius:4px; background:#eaf3fb; color:#1a5276; }
  .btn:hover { background:#1a5276; color:#fff; }
  .ref-col { text-align:left !important; color:#555; font-size:11px; min-width:120px; }
  .ref-id  { background:#f0f4f8; border:1px solid #d0dae4; border-radius:3px;
             padding:1px 5px; font-family:monospace; font-size:11px; }
  .badge          { display:inline-block; font-size:10px; font-weight:600; border-radius:3px;
                    padding:1px 6px; margin-left:5px; vertical-align:middle; opacity:0.75; }
  .badge-reseller    { background:#d4e6f1; color:#1a5276; }
  .badge-serviceorg  { background:#d5f5e3; color:#1e8449; }
  .badge-endcustomer { background:#fdebd0; color:#a04000; }
  .badge-site        { background:#f2f3f4; color:#555; }
  .badge-plan        { background:#e8f4f8; color:#0e6680; }
  .badge-contract    { background:#eae4f5; color:#5a3a8a; }
  .badge-count       { background:#e0e7ef; color:#1a3a5c; font-weight:700; opacity:1; min-width:18px; text-align:center; }
  .commit-qty        { color:#8b6914; font-size:11px; cursor:default; white-space:nowrap; }
  .commit-badge      { display:inline-block; font-size:9px; font-weight:500; color:#c8a84b;
                       background:rgba(200,168,75,.15); border-radius:3px; padding:0 4px;
                       white-space:nowrap; margin-top:2px; }
  /* ── Month-over-month delta styles ── */
  .delta-up          { color:#1a8f3c; font-size:10px; font-weight:600; white-space:nowrap; cursor:default; }
  .delta-down        { color:#c0392b; font-size:10px; font-weight:600; white-space:nowrap; cursor:default; }
  .badge-new         { display:inline-block; font-size:8px; font-weight:700; color:#fff; background:#27ae60;
                       border-radius:3px; padding:1px 4px; margin-left:4px; vertical-align:middle; letter-spacing:.5px; }
  .badge-churned     { display:inline-block; font-size:8px; font-weight:700; color:#fff; background:#95a5a6;
                       border-radius:3px; padding:1px 4px; margin-left:4px; vertical-align:middle; letter-spacing:.5px; }
  .row-churned td    { opacity:0.5; }
  .churned-name      { text-decoration:line-through; color:#999; }
  .churned-val       { color:#999 !important; }
  .efile-footnote    { margin:6px 14px 10px 14px; font-size:11px; color:#777; font-style:italic; }
  /* ── Reconciliation block ── */
  .recon-check       { color:#1a8f3c; font-size:10px; font-weight:700; margin-left:3px; cursor:default; }
  .recon-ok          { color:#1a8f3c; font-weight:700; }
  .recon-fail        { color:#c0392b; font-weight:700; }
  .recon-delta-up    { color:#1a8f3c; font-weight:600; }
  .recon-delta-dn    { color:#c0392b; font-weight:600; }
  .recon-box         { margin:8px 14px 0 14px; background:#f8fbff; border:1px solid #c8ddf0; border-radius:5px; padding:0; }
  .recon-title       { font-size:11px; font-weight:600; color:#2c5282; }
  .recon-toggle      { cursor:pointer; user-select:none; display:flex; align-items:center; justify-content:space-between; width:100%; padding:7px 12px; border-radius:5px; box-sizing:border-box; }
  .recon-toggle:hover { background:#dceaf7; color:#1a3a6e; }
  .recon-btn         { pointer-events:none; flex-shrink:0; margin-left:12px; padding:2px 9px; font-size:10px; font-weight:700; letter-spacing:0.04em; color:#2c5282; background:#dceaf7; border:1px solid #a8c4df; border-radius:3px; white-space:nowrap; }
  .recon-toggle:hover .recon-btn { background:#c4d9ef; }
  .recon-body        { padding:0 12px 10px 12px; }
  .recon-body.collapsed { display:none; }
  .recon-table       { font-size:10px; border-collapse:collapse; width:auto; }
  .recon-table th    { background:#e8f0f8; color:#2c5282; font-weight:600; padding:3px 8px; border:1px solid #c8ddf0; white-space:nowrap; }
  .recon-table td    { padding:2px 8px; border:1px solid #dde8f0; white-space:nowrap; }
  .recon-table tbody tr:hover td { background:#eef4fc; }
  .recon-note        { font-size:9px; color:#888; margin:4px 0 0 0; font-style:italic; }
  .recon-period      { font-weight:400; color:#7a9bb5; font-size:9px; }
  /* ── Unit Prices tab ── */
  .price-panel  { display:flex; flex-direction:column; overflow:hidden; }
  .price-scroll { flex:1; min-height:0; overflow:auto; border-top:1px solid #e0e8f0; }
  .price-tbl    { border-collapse:separate; border-spacing:0; width:max-content; min-width:100%; background:#fff; }
  .price-tbl thead tr { background:#1a5276; color:#fff; }
  .price-tbl th { position:sticky; top:0; z-index:2; background:#1a5276; padding:7px 12px; font-size:11px; font-weight:600; text-align:left; white-space:nowrap; }
  .price-prod-th { min-width:340px; text-align:left; }
  .price-ctr-th  { min-width:110px; text-align:center; }
  .price-num-th  { min-width:120px; text-align:right; }
  .price-tbl td { padding:5px 12px; font-size:11px; border-bottom:1px solid #edf0f4; white-space:nowrap; background:#fff; text-align:left; }
  .price-tbl tbody tr:hover td { background:#eef5fb !important; }
  .price-prod   { min-width:340px; text-align:left; }
  .price-ctr    { min-width:110px; text-align:center; }
  .price-val    { text-align:right; min-width:120px; font-variant-numeric:tabular-nums; }
  .price-chg    { text-align:right; min-width:100px; font-weight:600; }
  .price-absent { color:#bbb; }
  .price-up     { color:#c0392b; }
  .price-dn     { color:#1a8f3c; }
  .price-new    { color:#1a5276; font-size:10px; font-weight:700; letter-spacing:.04em; }
  .price-gone   { color:#888; font-size:10px; font-weight:700; letter-spacing:.04em; }
  .price-row-gone td { color:#aaa; }
  .price-period-th  { min-width:175px; text-align:center; }
  .price-period-td  { min-width:175px; text-align:center; color:#66809a; font-size:10.5px; }
  .price-footnote   { flex-shrink:0; padding:5px 14px 7px 14px; font-size:10.5px; color:#7a9bb5; border-top:1px solid #e8eef4; background:#fafdff; }
  /* ── Cross-reference tab ── */
  /* All panels are flex columns — toolbar fixed, table scrolls */
  .tab-panel { display:flex; flex-direction:column; overflow:hidden; }
  .tab-scroll { flex:1; min-height:0; overflow:auto; border-top:1px solid #e0e8f0; }
  /* xref panel — same pattern, different scroll div name */
  .xref-panel { display:flex; flex-direction:column; overflow:hidden; }
  .xref-scroll { flex:1; min-height:0; overflow:auto; border-top:1px solid #e0e8f0; }
  .xref-tbl { border-collapse:separate; border-spacing:0; width:max-content; min-width:100%; background:#fff; }
  .xref-tbl thead tr { background:#1a5276; color:#fff; }
  .xref-tbl th { position:sticky; top:0; z-index:2; background:#1a5276; padding:9px 12px;
                 text-align:left; font-weight:600; font-size:12px; white-space:normal; word-break:break-word; }
  .xref-tbl .xref-idx-th    { width:auto; min-width:22px; max-width:36px; text-align:center; position:sticky; left:0; z-index:3; padding:9px 1px !important; font-size:10px; }
  .xref-tbl .xref-status-th { width:110px; position:sticky; left:36px; z-index:3; }
  .xref-tbl .xref-canonical-th { min-width:260px; position:sticky; left:146px; z-index:3; box-shadow:2px 0 4px rgba(0,0,0,.08); }
  .xref-tbl .xref-prod-th   { min-width:240px; }
  .xref-tbl tbody tr { border-bottom:1px solid #e8ecf0; }
  .xref-tbl tbody tr:hover td { background:#eef5fb !important; }
  .xref-tbl td { padding:6px 12px; text-align:left; white-space:nowrap; background:#fff; }
  .xref-idx      { text-align:center !important; color:#aaa; font-size:10px; padding:6px 1px !important; position:sticky; left:0; z-index:1; }
  .xref-status   { position:sticky; left:36px; z-index:1; }
  .xref-canonical{ font-weight:600; position:sticky; left:146px; z-index:1; box-shadow:2px 0 4px rgba(0,0,0,.08); }
  .xref-absent     { text-align:center !important; color:#ccc; font-size:18px; line-height:1; }
  .xref-match-exact { color:#1a6b1a; }
  .xref-match-norm  { color:#1a6b1a; font-style:italic; }
  .xref-match-fuzzy { color:#7d5a00; font-style:italic; }
  .xref-conf { font-size:10px; color:#aaa; margin-left:3px; }
  .xref-pill { display:inline-block; padding:2px 9px; border-radius:10px; font-size:11px; font-weight:600; white-space:nowrap; }
  .xref-pill-green   { background:#d5f0d5; color:#1a6b1a; }
  .xref-pill-amber   { background:#fef3cd; color:#7d5a00; }
  .xref-pill-partial { background:#fde8cc; color:#8b4000; }
  .xref-actions { display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:0 12px 10px 12px; }
  .xref-chk { font-size:12px; color:#1a3a5c; display:flex; align-items:center; gap:4px; cursor:pointer; white-space:nowrap; }
  .xref-stats-line { display:flex; gap:6px; margin-left:8px; align-items:center; }
  .xsi-box { font-size:11px; font-weight:600; padding:2px 9px; border-radius:10px;
             background:#e8edf2; color:#1a3a5c; white-space:nowrap; }
  .xsi-green   { background:#d5f0d5; color:#1a6b1a; }
  .xsi-amber   { background:#fef3cd; color:#7d5a00; }
  .xsi-partial { background:#fde8cc; color:#8b4000; }
  /* ── Privacy blur toggle ── */
  .privacy-blur                { filter:blur(7px); transition:filter .15s; user-select:none; pointer-events:none; }
  body.privacy-on .priv-rest  { filter:blur(7px); transition:filter .15s; user-select:none; pointer-events:none; }
  body.privacy-on #acct-rest  { filter:blur(7px); transition:filter .15s; user-select:none; }
  .privacy-btn      { padding:4px 12px; font-size:12px; cursor:pointer; border:1px solid #7b8fa6;
                      border-radius:4px; background:#eaf3fb; color:#1a3a5c; font-family:inherit; white-space:nowrap; }
  .privacy-btn:hover  { background:#d0e4f2; }
  .privacy-btn.active { background:#c0392b; color:#fff; border-color:#922b21; }
</style>
</head>
<body>
<div class="page-header">
<div style="display:flex;align-items:center;gap:10px;margin-bottom:2px">
$(  $acctWords = $Account -split '\s+', 2
    if ($acctWords.Count -gt 1) {
        "<h2 style='margin:0'>$([System.Web.HttpUtility]::HtmlEncode($acctWords[0])) <span id='acct-rest'>$([System.Web.HttpUtility]::HtmlEncode($acctWords[1]))</span> &mdash; $Period Usage Summary</h2>"
    } else {
        "<h2 style='margin:0'>$([System.Web.HttpUtility]::HtmlEncode($Account)) &mdash; $Period Usage Summary</h2>"
    }
)
<button id="privacy-btn" class="privacy-btn" onclick="togglePrivacy()" title="Blur customer names and price data">&#128274; Privacy</button>
</div>
<p class="sub">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp; $(if ($Token) { "$($Rows.Count) devices &nbsp;|&nbsp; $($Summary.Count) summary rows &nbsp;|&nbsp; API: ${ApiSeconds}s &nbsp;|&nbsp; " })Total: $([math]::Round(((Get-Date) - $scriptStart).TotalSeconds, 1))s &nbsp;|&nbsp; <a href="https://developer.n-able.com/n-able-billing-api/reference" target="_blank" style="color:#5b8abf;text-decoration:none" title="N-able Billing API Reference">&#128196; Billing API Docs</a></p>
</div>
<div class="tab-nav">
$($tabNavHtml.ToString())
</div>
<div class="tab-wrap">
$($tabPanelHtml.ToString())
</div>
<script>
function toggle(row) {
    var id   = row.dataset.id;
    var open = row.dataset.open === '1';
    row.dataset.open = open ? '0' : '1';
    var tog = row.querySelector('.tog');
    if (tog) tog.textContent = open ? '\u25BA' : '\u25BC';
    document.querySelectorAll('tr[data-parent="' + id + '"]').forEach(function(child) {
        if (open) {
            child.style.display = 'none';
            child.dataset.open  = '0';
            var ct = child.querySelector('.tog');
            if (ct) ct.textContent = '\u25BA';
            collapseDesc(child.dataset.id);
        } else {
            child.style.display = '';
        }
    });
}
function collapseDesc(id) {
    document.querySelectorAll('tr[data-parent="' + id + '"]').forEach(function(child) {
        child.style.display = 'none';
        child.dataset.open  = '0';
        var ct = child.querySelector('.tog');
        if (ct) ct.textContent = '\u25BA';
        collapseDesc(child.dataset.id);
    });
}
function expandAll(tabId) {
    var panel = document.getElementById(tabId);
    panel.querySelectorAll('tbody tr').forEach(function(r) { r.style.display=''; r.dataset.open='1'; });
    panel.querySelectorAll('.tog').forEach(function(t) { t.textContent='\u25BC'; });
}
function collapseAll(tabId) {
    var panel = document.getElementById(tabId);
    panel.querySelectorAll('tbody tr[data-parent]').forEach(function(r) { r.style.display='none'; r.dataset.open='0'; });
    panel.querySelectorAll('.tog').forEach(function(t) { t.textContent='\u25BA'; });
}
function filterRows(tabId, input) {
    var term = input.value.toLowerCase().trim();
    var panel = document.getElementById(tabId);
    if (!term) { clearFilter(tabId); return; }
    panel.querySelectorAll('tbody tr').forEach(function(r) {
        var cell = r.querySelector('.name-col');
        var text = cell ? cell.textContent.toLowerCase() : '';
        r.style.display = text.indexOf(term) >= 0 ? '' : 'none';
    });
}
function clearFilter(tabId) {
    var panel = document.getElementById(tabId);
    var input = document.getElementById('search-' + tabId);
    if (input) input.value = '';
    panel.querySelectorAll('tbody tr').forEach(function(r) {
        r.style.display = parseInt(r.dataset.depth || '0') === 0 ? '' : 'none';
        r.dataset.open = '0';
    });
    panel.querySelectorAll('.tog').forEach(function(t) { t.textContent = '\u25BA'; });
}
function toggleRecon(titleEl) {
    var body = titleEl.nextElementSibling;
    var btn  = titleEl.querySelector('.recon-btn');
    if (body.classList.contains('collapsed')) {
        body.classList.remove('collapsed');
        if (btn) btn.innerHTML = 'RECONCILE&nbsp;&#9660;&nbsp;COLLAPSE';
    } else {
        body.classList.add('collapsed');
        if (btn) btn.innerHTML = 'RECONCILE&nbsp;&#9654;&nbsp;EXPAND';
    }
}
function syncPreserve(cb) {
    document.querySelectorAll('.preserve-cb').forEach(function(el) { el.checked = cb.checked; });
}
function showTab(tabId, btn) {
    var preserve = document.querySelector('.preserve-cb');
    var filterTerm = '';
    if (preserve && preserve.checked) {
        var activePanel = document.querySelector('.tab-panel[style*="flex"], .tab-panel:not([style*="none"])');
        if (activePanel) {
            var activeInput = activePanel.querySelector('.search-box');
            if (activeInput) filterTerm = activeInput.value;
        }
    }
    document.querySelectorAll('.tab-panel').forEach(function(p) { p.style.display='none'; });
    document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    document.getElementById(tabId).style.display = '';
    btn.classList.add('active');
    if (filterTerm) {
        var newInput = document.getElementById('search-' + tabId);
        if (newInput) { newInput.value = filterTerm; filterRows(tabId, newInput); }
    }
}
function addCountBadges() {
    var counts = {};
    document.querySelectorAll('tr[data-parent]').forEach(function(r) {
        var p = r.dataset.parent;
        counts[p] = (counts[p] || 0) + 1;
    });
    document.querySelectorAll('tr[data-id]').forEach(function(r) {
        var cnt = counts[r.dataset.id];
        if (cnt) {
            var cell = r.querySelector('.name-col');
            if (cell) {
                var b = document.createElement('span');
                b.className = 'badge badge-count';
                b.textContent = cnt;
                cell.appendChild(b);
            }
        }
    });
}
addCountBadges();
// Expand all collapsible tabs on load
document.querySelectorAll('.tab-panel[id^="tab"]').forEach(function(panel) {
    if (!panel.classList.contains('xref-panel')) { expandAll(panel.id); }
});
function filterPrice(src) {
    var term = src.value.toLowerCase().trim();
    document.querySelectorAll('.price-tbl tbody tr').forEach(function(r) {
        var prod = (r.cells[0] || {textContent:''}).textContent.toLowerCase();
        r.style.display = (!term || prod.indexOf(term) >= 0) ? '' : 'none';
    });
}
function filterXref(src) {
    var term    = (document.getElementById('search-xref') || {value:''}).value.toLowerCase().trim();
    var fuzOnly = !!(document.getElementById('xchk-fuzzy')   && document.getElementById('xchk-fuzzy').checked);
    var parOnly = !!(document.getElementById('xchk-partial') && document.getElementById('xchk-partial').checked);
    document.querySelectorAll('.xref-tbl tbody tr').forEach(function(r) {
        var name   = (r.cells[2] || {textContent:''}).textContent.toLowerCase();
        var status = r.dataset.xstatus || '';
        var isFuz  = status === 'fuzzy' || status === 'fuzzy-partial';
        var isPar  = status === 'partial' || status === 'fuzzy-partial';
        var show   = (!term    || name.indexOf(term) >= 0)
                  && (!fuzOnly || isFuz)
                  && (!parOnly || isPar);
        r.style.display = show ? '' : 'none';
    });
}
function clearXref() {
    var inp = document.getElementById('search-xref');
    if (inp) inp.value = '';
    document.querySelectorAll('#xchk-fuzzy,#xchk-partial').forEach(function(c) { c.checked = false; });
    filterXref();
}
function togglePrivacy() {
    var btn = document.getElementById('privacy-btn');
    var on = btn.classList.toggle('active');
    btn.innerHTML = on ? '&#128275; Reveal' : '&#128274; Privacy';
    // Blur .priv-rest spans inside name cells + #acct-rest in title via CSS body class
    document.body.classList.toggle('privacy-on', on);
    // Blur all data in the Unit Prices tab (per-cell)
    document.querySelectorAll('.price-tbl tbody td').forEach(function(el) {
        el.classList.toggle('privacy-blur', on);
    });
}
</script>
</body>
</html>
"@

$HtmlFile = Join-Path $OutputDir "${Timestamp}_${SafeName}_Summary_${Timeframe}.html"
$html | Set-Content -Path $HtmlFile -Encoding UTF8
Write-Output "HTML report saved: $HtmlFile"
Invoke-Item $HtmlFile
