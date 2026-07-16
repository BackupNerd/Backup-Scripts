# ????????????????????????????????????????????????????????????????????????????
# ? Get-CoveVSSReport.ps1 - Standalone VSS Health Report Generator           ?
# ? Zero-MCP: Direct API auth -> Device enum -> InternalInfo fetch -> HTML      ?
# ????????????????????????????????????????????????????????????????????????????
#Requires -Version 7.0
#
# SYNOPSIS
#   Generates a dark-theme HTML VSS health report for a Cove partner.
#   Authenticates directly against api.backup.management/jsonapi (no MCP).
#   Fetches per-device InternalInfo diagnostic pages, parses VSS data,
#   and produces a self-contained HTML report matching the MCP-generated version.
#
# CREDENTIAL STORAGE
#   First run:  prompts via Get-Credential, saves DPAPI-encrypted XML.
#   Later runs: loads from XML silently (DPAPI bound to user + machine).
#   Use -ClearCredentials to re-prompt and overwrite the stored file.
#
# ARCHITECTURE
#   1. Load/prompt credentials -> Login -> visa
#   2. EnumerateAccountStatistics -> device list (AU AN AR MN T0 OS)
#   3. Per device: EnumerateAccountRemoteAccessEndpoints -> InternalInfoPageUrl
#   4. GET InternalInfoPageUrl (skip TLS) -> HTML page
#   5. Split HTML by <a name="anchor"> -> parse 6 VSS sections
#   6. Analyze each device (writers, storage, snapshots, services)
#   7. Generate dark-theme HTML report
#
# USAGE
#   .\Get-CoveVSSReport.ps1
#   .\Get-CoveVSSReport.ps1 -PartnerId 175407 -CredentialFile "C:\ProgramData\MXB\mycreds.xml"
#   .\Get-CoveVSSReport.ps1 -ClearCredentials   # re-prompt for creds
#   .\Get-CoveVSSReport.ps1 -OpenReport         # auto-open HTML in browser
#
<#
.SYNOPSIS
  Standalone VSS health report generator for Cove Data Protection.

.PARAMETER PartnerId
  Cove partner/customer hierarchical ID to scope device enumeration.
  Default: 175407 (Impact360)

.PARAMETER PartnerLabel
  Display label for report header. Default: "Impact360"

.PARAMETER CredentialFile
  DPAPI-encrypted PSCredential XML (Export-Clixml format).
  Default: C:\ProgramData\MXB\mcpcred-iaso.xml

.PARAMETER OutputPath
  Directory for HTML report output.
  Default: .\output\reports\<date>

.PARAMETER CachePath
  Directory for InternalInfo HTML cache files. Self-describing filenames include
  device name, AccountId, and fetch timestamp. 15-minute TTL - reuses cached
  file if younger than 15 min. Defaults to .\cache\internalinfo

.PARAMETER ClearCredentials
  Delete stored credential file and re-prompt on this run.

.PARAMETER OpenReport
  Open the generated HTML report in the default browser when done.

.PARAMETER DiagTimeoutSec
  Per-device InternalInfo page fetch timeout in seconds. Default: 20
#>

param(
    [int]   $PartnerId       = 000000, TenantID
    [string]$PartnerLabel    = "companyname",
    [string]$CredentialFile  = "C:\ProgramData\MXB\${env:computername}_${env:username}_API_Credentials.Secure.xml",
    [string]$OutputPath      = "",
    [string]$CachePath        = "",
    [switch]$ClearCredentials,
    [switch]$OpenReport = $true,
    [int]   $DiagTimeoutSec  = 60,
    [int[]] $TestAccountIds  = @(),  # Optional: test specific device AccountIds directly (bypasses enumeration)
    # T0 values to include. 0=OK 1=InProcess 2=Failed 5=Completed 8=CompletedWithErrors 9=InProgressWithFaults 12=Restarted
    [int[]] $T0Filter        = @(2, 8, 0, 5, 1, 9, 12),
    [switch]$AllDevices,
    [int]   $MaxDevices      = 500,
    [int]   $ParallelCount   = 15,
    [switch]$ShowEmailIcon   = $true,
    [switch]$ShowOpenFolderIcon = $true,
    [string]$EmailTo         = "",
    [string]$EmailSubjectTemplate = "VSS Health Report - {PartnerLabel} - {ReportDate}",
    [string]$EmailBodyTemplate = "Hello,`r`n`r`nPlease review the attached VSS Health Report.`r`n`r`nPartner: {PartnerLabel} ({PartnerId})`r`nDate: {ReportDate}`r`nSummary: RED={RedCount}, YELLOW={YellowCount}, GREEN={GreenCount}, OFFLINE={OfflineCount}`r`nReport path: {ReportPath}`r`nGenerated: {GeneratedUtc}`r`n",
    [switch]$CopyReportFileToClipboard,
    [switch]$GenerateEmailAndCopyFile,
    [switch]$OpenEmailDraftWithAttachment,
    [switch]$DebugTimeline,
    # S1 rollback window thresholds
    [int]$RollbackWarnBelow  = 1,   # WARN if S1 rollback coverage < N days
    [int]$RollbackTarget     = 7,   # INFO if coverage >= WarnBelow but < Target; silent if >= Target
    [int]$RollbackWarnAbove  = 30,  # WARN if coverage > N days (possibly excessive retention)
    # Shadow storage allocation threshold
    [int]$ShadowAllocMinPct  = 10,  # WARN if shadow max cap < N% of volume (Microsoft/S1 minimum)
    # Snapshot count threshold
    [int]$MaxSnapPerVol      = 180  # WARN if total snapshots on any volume exceeds N
)

$ErrorActionPreference = "Stop"
$COVE_API_URL = "https://api.backup.management/jsonapi"
$REPORT_DATE  = Get-Date -Format "yyyy-MM-dd"

if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot "output\reports\$REPORT_DATE"
}

# ============================================================================
# Credential helpers  (same DPAPI pattern as Archive-Cleanup.ps1)
# ============================================================================

function ConvertFrom-SecureString2 {
    param([System.Security.SecureString]$SecureString)
    $ptr  = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
    $text = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)
    return $text
}

function Expand-TemplateText {
    param(
        [string]$Template,
        [hashtable]$Tokens
    )
    $out = if ($null -ne $Template) { [string]$Template } else { "" }
    foreach ($k in $Tokens.Keys) {
        $out = $out.Replace("{$k}", [string]$Tokens[$k])
    }
    return $out
}

function New-EmailDraftEml {
    param(
        [string]$DraftPath,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$AttachmentName,
        [string]$AttachmentContent
    )

    $boundary = "----=_Part_" + [Guid]::NewGuid().ToString('N')
    $safeSubject = ($Subject ?? '') -replace '[\r\n]+', ' '
    $bodyB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($Body ?? '')), [Base64FormattingOptions]::InsertLineBreaks)
    $attachB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($AttachmentContent ?? '')), [Base64FormattingOptions]::InsertLineBreaks)

    $lines = @(
        'From: '
        "To: $To"
        "Subject: $safeSubject"
        'MIME-Version: 1.0'
        "Content-Type: multipart/mixed; boundary=`"$boundary`""
        ''
        "--$boundary"
        'Content-Type: text/plain; charset=utf-8'
        'Content-Transfer-Encoding: base64'
        ''
        $bodyB64
        ''
        "--$boundary"
        "Content-Type: text/html; name=`"$AttachmentName`""
        'Content-Transfer-Encoding: base64'
        "Content-Disposition: attachment; filename=`"$AttachmentName`""
        ''
        $attachB64
        ''
        "--$boundary--"
        ''
    )

    $eml = $lines -join "`r`n"
    [IO.File]::WriteAllText($DraftPath, $eml, [Text.UTF8Encoding]::new($false))
    return $DraftPath
}

function Open-OutlookDraftWithAttachment {
    param(
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$AttachmentPath
    )
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        if ($To) { $mail.To = $To }
        $mail.Subject = ($Subject ?? '')
        $mail.Body = ($Body ?? '')
        if ($AttachmentPath -and (Test-Path $AttachmentPath)) {
            $mail.Attachments.Add($AttachmentPath) | Out-Null
        }
        $mail.Display($true)
        return $true
    } catch {
        Write-Host "Outlook draft with attachment failed: $_" -ForegroundColor Yellow
        return $false
    }
}

function Set-ReportFileClipboard {
    param([string]$FilePath)

    if (-not $FilePath -or -not (Test-Path $FilePath)) { return $false }
    $fp = $FilePath.Replace("'", "''")
    $cmd = @"
Add-Type -AssemblyName System.Windows.Forms
`$files = New-Object System.Collections.Specialized.StringCollection
`$null = `$files.Add('$fp')
[System.Windows.Forms.Clipboard]::SetFileDropList(`$files)
"@
    try {
        & pwsh -NoProfile -Sta -Command $cmd | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Get-CoveCredentials {
    param([string]$CredFile, [switch]$Clear)

    if ($Clear -and (Test-Path $CredFile)) {
        Remove-Item $CredFile -Force
        Write-Host "Credential file removed. Re-prompting..." -ForegroundColor Yellow
    }

    if (Test-Path $CredFile) {
        Write-Host "Loading credentials from $CredFile" -ForegroundColor Cyan
        $stored = Import-Clixml -Path $CredFile

        # PSA v10 format: PSCustomObject with PartnerName, Username, Password (ConvertFrom-SecureString)
        if ($stored.PSObject.Properties['PartnerName']) {
            $secPwd   = $stored.Password | ConvertTo-SecureString
            $password = ConvertFrom-SecureString2 -SecureString $secPwd
            Write-Host "  Partner : $($stored.PartnerName)" -ForegroundColor DarkCyan
            Write-Host "  User    : $($stored.Username)" -ForegroundColor DarkCyan
            return @{ PartnerName = $stored.PartnerName; Username = $stored.Username; Password = $password }
        }

        # Legacy PSCredential format (mcpcred-*.xml) - no PartnerName stored
        Write-Host "  User: $($stored.UserName)  (legacy format)" -ForegroundColor DarkCyan
        return @{ PartnerName = ''; Username = $stored.UserName; Password = ConvertFrom-SecureString2 -SecureString $stored.Password }

    } else {
        Write-Host "No credential file found at $CredFile" -ForegroundColor Yellow
        Write-Host "  Creating new credential file (PSA v10 format)..." -ForegroundColor Yellow

        $partnerName = ''
        do { $partnerName = Read-Host "  Enter Backup.Management Partner Name (e.g. 'Acme, Inc')" }
        while ($partnerName.Length -eq 0)

        $cred = Get-Credential -Message 'Enter Backup.Management login (email + password)'
        if (-not $cred) { throw "No credentials provided." }

        $dir = Split-Path $CredFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        [PSCustomObject]@{
            PartnerName = $partnerName
            Username    = $cred.UserName
            Password    = ($cred.Password | ConvertFrom-SecureString)
        } | Export-Clixml -Path $CredFile -Force

        Write-Host "Credentials saved to $CredFile" -ForegroundColor Green
        return @{ PartnerName = $partnerName; Username = $cred.UserName; Password = ConvertFrom-SecureString2 -SecureString $cred.Password }
    }
}

# ============================================================================
# Cove JSON-RPC API
# ============================================================================

function Invoke-CoveAPI {
    param(
        [string]   $Method,
        [hashtable]$Params,
        [string]   $Visa = $null
    )

    $payload = @{ id = 1; jsonrpc = "2.0"; method = $Method; params = $Params }
    if ($Visa) { $payload.visa = $Visa }

    $json = ConvertTo-Json -InputObject $payload -Depth 10

    try {
        $resp = Invoke-RestMethod -Uri $COVE_API_URL -Method Post `
            -Headers @{ "Content-Type" = "application/json" } -Body $json -ErrorAction Stop
        return $resp
    } catch {
        return @{ error = @{ message = $_.Exception.Message; code = -1 }; result = $null }
    }
}

function Invoke-CoveLogin {
    param([hashtable]$Creds)
    $display = if ($Creds.PartnerName) { "$($Creds.Username) (partner: $($Creds.PartnerName))" } else { $Creds.Username }
    Write-Host "Authenticating as $display..." -ForegroundColor Cyan
    $resp = Invoke-CoveAPI -Method "Login" -Params @{
        username = $Creds.Username
        password = $Creds.Password
    }
    if ($resp.error) { throw "Login failed: $($resp.error.message)" }
    $visa = $resp.visa
    if (-not $visa) { throw "Login returned no visa token." }
    Write-Host "Login OK." -ForegroundColor Green
    return $visa
}

# ============================================================================
# Device Enumeration
# ============================================================================

function Get-CoveDevices {
    param([string]$Visa, [int]$PartnerId)

    Write-Host "Enumerating devices for partner $PartnerId..." -ForegroundColor Cyan

    $resp = Invoke-CoveAPI -Method "EnumerateAccountStatistics" -Visa $Visa -Params @{
        query = @{
            PartnerId          = $PartnerId
            RecurseSubPartners = $true
            Columns            = @("AU", "AN", "AR", "MN", "T0", "OS", "OT", "IP", "VN", "PD", "PN", "OI", "OP", "RU", "MF", "MO", "I84", "I85", "CD", "TS")
            StartRecordNumber  = 0
            RecordsCount       = 5000
        }
    }

    if ($resp.error) { throw "EnumerateAccountStatistics failed: $($resp.error.message)" }

    # Double-wrapped result: result.result
    $raw = $resp.result
    if ($raw -and $raw.PSObject.Properties['result']) { $raw = $raw.result }

    if (-not $raw) { throw "No devices returned." }

    $devices = @()
    foreach ($entry in $raw) {
        $flat = @{}
        foreach ($setting in $entry.Settings) {
            foreach ($key in $setting.PSObject.Properties.Name) {
                $flat[$key] = $setting.$key
            }
        }
        $devices += [PSCustomObject]@{
            AccountId  = [int]($flat['AU'] ?? 0)
            DeviceName = (($flat['AN'] ?? '') -replace '\s','').Trim()
            AR         = ($flat['AR'] ?? '').Trim()
            MN         = ($flat['MN'] ?? '').Trim()
            T0         = [int]($flat['T0'] ?? 0)
            OT         = [int]($flat['OT'] ?? 0)
            OS         = ($flat['OS'] ?? '').Trim()
            IP         = ($flat['IP'] ?? '').Trim()
            VN         = ($flat['VN'] ?? '').Trim()
            PD         = ($flat['PD'] ?? '').Trim()
            PN         = ($flat['PN'] ?? '').Trim()
            OI         = ($flat['OI'] ?? '').Trim()
            OP         = ($flat['OP'] ?? '').Trim()
            RU         = ($flat['RU'] ?? '').Trim()
            CD         = ($flat['CD'] ?? '').Trim()   # device creation date (unix seconds or ISO string)
        }
    }

    Write-Host "Found $($devices.Count) device(s) total." -ForegroundColor Green
    return $devices
}

# ============================================================================
# InternalInfo Page Fetch
# ============================================================================

function Get-InternalInfoUrl {
    param([string]$Visa, [int]$AccountId)

    $resp = Invoke-CoveAPI -Method "EnumerateAccountRemoteAccessEndpoints" -Visa $Visa -Params @{
        accountId = $AccountId
    }

    if ($resp.error) { return $null }

    # Unwrap double-nested result: result.result = [...]
    $endpoints = $resp.result
    if ($endpoints -and $endpoints.PSObject.Properties['result']) {
        $endpoints = $endpoints.result
    }

    if (-not $endpoints -or $endpoints.Count -eq 0) { return $null }

    # Endpoint object has InternalInfoPageUrl directly (no 'ep' wrapper)
    $ep = $endpoints[0]
    return $ep.InternalInfoPageUrl
}

function Get-DeviceRepserv {
    param([string]$Visa, [int]$AccountId)

    $resp = Invoke-CoveAPI -Method "GetAccountInfoById" -Visa $Visa -Params @{
        accountId = $AccountId
    }

    if ($resp.error) { return $null }
    $info = $resp.result
    if (-not $info) { return $null }

    $hostStr = ($info.homeNodeInfo?.CommonInfo?.Host ?? '')
    if (-not $hostStr) { return $null }

    $hostNoPort = $hostStr.Split(':')[0].ToLower()
    return @{
        RepUrl = "https://$hostNoPort/repserv_json"
        Token  = ($info.result?.Token ?? '')
        Name   = ($info.result?.Name  ?? '')
    }
}

function Get-DeviceLastSessions {
    # Returns hashtable: { pluginId (int) -> lastCompletedEndTime (unix) }
    param([string]$Visa, [hashtable]$Repserv, [int]$AccountId)

    $payload = @{
        jsonrpc = '2.0'
        id      = 1
        method  = 'QuerySessions'
        visa    = $Visa
        params  = @{
            accountId = $AccountId
            orderBy   = 'Id DESC'
            query     = '0 != 1'
            range     = @{ Offset = 0; Size = 500 }
            account   = $Repserv.Name
            token     = $Repserv.Token
        }
    } | ConvertTo-Json -Depth 10

    try {
        $resp = Invoke-RestMethod -Uri $Repserv.RepUrl -Method POST `
            -ContentType 'application/json' -Body $payload -ErrorAction Stop
    } catch {
        return @{}
    }

    if ($resp.error) { return @{} }

    $sessions = if ($resp.result -and $resp.result.PSObject.Properties['result']) {
        $resp.result.result
    } else { $resp.result }
    if (-not $sessions) { return @{} }

    # PluginId can be an integer or string enum - normalize to int
    $PLUGIN_ID_STR = @{
        "WorkstationFileSystem" = 1; "FilesAndFolders" = 1; "FileSystem" = 1
        "Exchange" = 4
        "NetworkShares" = 6
        "SystemState" = 7; "VssSystemState" = 7
        "VMware" = 8
        "VssMsSql" = 10; "MsSql" = 10
        "HyperV" = 14
        "MySql" = 15
    }

    $STATUS_INT = @{
        "Completed"=5; "CompletedWithErrors"=3; "Failed"=1; "Aborted"=2
        "InProgress"=0; "NotStarted"=4; "Overdue"=6; "PartiallyCompleted"=3
    }
    $completedPlugin = @{}
    $failedPlugin    = @{}
    $sessFullPlugin  = @{}
    foreach ($s in $sessions) {
        $rawPid  = $s.PluginId
        $plugId  = if ($rawPid -match '^\d+$') { [int]$rawPid }
                   elseif ($PLUGIN_ID_STR.ContainsKey("$rawPid")) { $PLUGIN_ID_STR["$rawPid"] }
                   else { 0 }
        $end       = [long]($s.EndTime ?? 0)
        $start     = [long]($s.StartTime ?? 0)
        $rawStatus = $s.Status
        $status    = if ($rawStatus -match '^\d+$') { [int]$rawStatus }
                     elseif ($STATUS_INT.ContainsKey("$rawStatus")) { $STATUS_INT["$rawStatus"] }
                     else { -1 }
        $statusStr = switch ($status) { 5 { 'Completed' } 3 { 'CompletedWithErrors' } 1 { 'Failed' } 2 { 'Aborted' } 6 { 'Skipped' } default { "$rawStatus" } }
        if ($plugId -gt 0 -and $end -gt 0) {
            if (-not $sessFullPlugin.ContainsKey($plugId)) { $sessFullPlugin[$plugId] = [System.Collections.Generic.List[hashtable]]::new() }
            $sessFullPlugin[$plugId].Add(@{ start = $start; end = $end; status = $statusStr; cleaned = ($s.Flags -contains 'Cleaned'); accelerated = ($s.Flags -contains 'Accelerated'); buildVersion = "$($s.BuildVersion)".Trim(); changedCount = [int]($s.ChangedCount ?? 0); changedBytes = [long]($s.ChangedSize ?? 0); errorsCount = [int]($s.ErrorsCount ?? 0); selectedCount = [int]($s.SelectedCount ?? 0); selectedBytes = [long]($s.SelectedSize ?? 0); newCount = [int]($s.NewCount ?? 0); newBytes = [long]($s.NewSize ?? 0); sentBytes = [long]($s.SentSize ?? 0); deletedCount = [int]($s.DeletedCount ?? 0) })
            if ($status -eq 5) {
                if (-not $completedPlugin.ContainsKey($plugId)) { $completedPlugin[$plugId] = [System.Collections.Generic.List[long]]::new() }
                $completedPlugin[$plugId].Add($end)
            } elseif ($status -in @(1, 2, 3)) {
                if (-not $failedPlugin.ContainsKey($plugId)) { $failedPlugin[$plugId] = [System.Collections.Generic.List[long]]::new() }
                $failedPlugin[$plugId].Add($end)
            }
        }
    }
    # Merge all plugin IDs; sort each list descending
    $allPlugins = @($completedPlugin.Keys + $failedPlugin.Keys + $sessFullPlugin.Keys | Sort-Object -Unique)
    $result = @{}
    foreach ($k in $allPlugins) {
        $result[$k] = @{
            completed     = if ($completedPlugin.ContainsKey($k)) { @($completedPlugin[$k] | Sort-Object -Descending) } else { @() }
            failed        = if ($failedPlugin.ContainsKey($k))    { @($failedPlugin[$k]    | Sort-Object -Descending) } else { @() }
            sessions_full = if ($sessFullPlugin.ContainsKey($k))  { @($sessFullPlugin[$k]  | Sort-Object { $_.end } -Descending) } else { @() }
        }
    }
    return $result
}

function Save-SessionsToCache {
    # Saves parsed session data to CSV for later analysis
    param([string]$CacheDir, [string]$DeviceName, [int]$AccountId, [hashtable]$LastSessions)
    
    if (-not $LastSessions -or $LastSessions.Count -eq 0) { return }
    
    $cacheSessionDir = Join-Path $CacheDir 'sessions'
    if (-not (Test-Path $cacheSessionDir)) { New-Item -ItemType Directory -Path $cacheSessionDir -Force | Out-Null }
    
    $safeName  = $DeviceName -replace '[^\w.-]', '_'
    $stamp     = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $csvFile   = Join-Path $cacheSessionDir "$safeName-$AccountId-sessions-$stamp.csv"
    
    $csvRows = @()
    foreach ($plugId in $LastSessions.Keys) {
        foreach ($sess in $LastSessions[$plugId].sessions_full) {
            $startDt = [DateTimeOffset]::FromUnixTimeSeconds($sess.start).LocalDateTime
            $endDt   = [DateTimeOffset]::FromUnixTimeSeconds($sess.end).LocalDateTime
            $dur     = [int](($endDt - $startDt).TotalSeconds / 60)
            $csvRows += [PSCustomObject]@{
                PluginId    = $plugId
                StartTime   = $startDt.ToString('yyyy-MM-dd HH:mm:ss')
                EndTime     = $endDt.ToString('yyyy-MM-dd HH:mm:ss')
                DurationMin = $dur
                Status      = $sess.status
                StartUnix   = $sess.start
                EndUnix     = $sess.end
            }
        }
    }
    
    if ($csvRows.Count -gt 0) {
        $csvRows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Force | Out-Null
    }
}

function Load-SessionsFromCache {
    # Loads cached session data for a device if available (TTL-based)
    param([string]$CacheDir, [string]$DeviceName, [int]$AccountId, [int]$TtlMin = 5)
    
    $cacheSessionDir = Join-Path $CacheDir 'sessions'
    if (-not (Test-Path $cacheSessionDir)) { return $null }
    
    $safeName = $DeviceName -replace '[^\w.-]', '_'
    $pattern  = "$safeName-$AccountId-sessions-*.csv"
    
    $cachedFile = Get-ChildItem -Path $cacheSessionDir -Filter $pattern -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($cachedFile -and ((Get-Date) - $cachedFile.LastWriteTime).TotalMinutes -lt $TtlMin) {
        try {
            $data = Import-Csv -Path $cachedFile.FullName -ErrorAction Stop
            return @{ data = $data; file = $cachedFile.Name }
        } catch {
            return $null
        }
    }
    return $null
}

function Invoke-InternalInfoPage {
    param([string]$Url, [int]$TimeoutSec)

    try {
        # Skip cert validation (agent uses self-signed cert)
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -SkipCertificateCheck `
                -TimeoutSec $TimeoutSec -ErrorAction Stop
        } else {
            # PowerShell 5 - use ServicePointManager bypass
            [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing `
                -TimeoutSec $TimeoutSec -ErrorAction Stop
        }
        return $resp.Content
    } catch {
        return $null
    }
}

# ============================================================================
# HTML Parsing
# ============================================================================

function Strip-HtmlTags {
    param([string]$Html)
    if (-not $Html) { return '' }
    $text = [regex]::Replace($Html, '<[^>]+>', '')
    # Decode common HTML entities
    $text = $text -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' `
                  -replace '&nbsp;',' ' -replace '&#39;',"'" -replace '&quot;','"'
    return $text.Trim()
}

function Split-InternalInfoSections {
    param([string]$Html)

    $sections = @{}
    if (-not $Html) { return $sections }

    # Split on <a name="sectionname"> anchors (case-insensitive)
    $anchorPattern = [regex]'<a\s+name\s*=\s*"([^"]+)"\s*>\s*</a>'
    $matches = $anchorPattern.Matches($Html)

    if ($matches.Count -eq 0) { return $sections }

    for ($i = 0; $i -lt $matches.Count; $i++) {
        $name  = $matches[$i].Groups[1].Value.ToLower()
        $start = $matches[$i].Index + $matches[$i].Length
        $end   = if ($i + 1 -lt $matches.Count) { $matches[$i+1].Index } else { $Html.Length }
        $sections[$name] = $Html.Substring($start, $end - $start)
    }

    return $sections
}

function Get-HtmlTableData {
    # Returns array of string-arrays: first row may be headers (TH), rest are data (TD)
    param([string]$SectionHtml)

    $trOpts = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor
              [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $trPattern = [regex]::new('<tr\b[^>]*>(.*?)</tr>', $trOpts)
    $cellPat   = [regex]::new('<t[dh]\b[^>]*>(.*?)</t[dh]>', $trOpts)

    $tableRows = @()
    foreach ($tr in $trPattern.Matches($SectionHtml)) {
        $cells = @($cellPat.Matches($tr.Groups[1].Value) |
                   ForEach-Object { Strip-HtmlTags $_.Groups[1].Value })
        if ($cells.Count -gt 0) { $tableRows += ,@($cells) }
    }
    return $tableRows
}

function Parse-DlSection {
    # Parses a <dl>/<dt>/<dd> definition-list HTML section into an ordered hashtable.
    # Matches server.py _kv_pairs() behavior for vsswriter_* sections.
    param([string]$SectionHtml)

    $opts  = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor
             [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $dtPat = [regex]::new('<dt[^>]*>(.*?)</dt>', $opts)
    $ddPat = [regex]::new('<dd[^>]*>(.*?)</dd>', $opts)

    $dts = @($dtPat.Matches($SectionHtml) | ForEach-Object { (Strip-HtmlTags $_.Groups[1].Value).Trim() })
    $dds = @($ddPat.Matches($SectionHtml) | ForEach-Object { (Strip-HtmlTags $_.Groups[1].Value).Trim() })

    $result = [ordered]@{}
    for ($i = 0; $i -lt $dts.Count; $i++) {
        $key = $dts[$i]
        $val = if ($i -lt $dds.Count) { $dds[$i] } else { '' }
        if ($key -and $key -notmatch '^components?$') { $result[$key] = $val }
    }
    return $result
}

function Parse-TableSection {
    # Parses an HTML table section into PSCustomObject array using first row as headers.
    param([string]$SectionHtml)

    $rows = Get-HtmlTableData $SectionHtml
    if ($rows.Count -lt 2) { return @() }

    $headers = $rows[0]
    $result  = @()

    for ($i = 1; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        $obj = [ordered]@{}
        for ($j = 0; $j -lt $headers.Count; $j++) {
            $obj[$headers[$j]] = if ($j -lt $row.Count) { $row[$j] } else { '' }
        }
        # Skip empty rows
        $vals = @($obj.Values | Where-Object { $_ -ne '' })
        if ($vals.Count -gt 0) { $result += [PSCustomObject]$obj }
    }
    return $result
}

function Parse-EventLogs {
    # Parses all log_* sections from the InternalInfo Sections hashtable.
    # Returns a flat list of PSCustomObjects: Time, Message, Level, Provider
    # Filters to rows where Level matches a known Windows event log level to avoid
    # picking up non-event rows from other tables on the page.
    param([hashtable]$Sections)

    $events   = [System.Collections.Generic.List[object]]::new()
    $logKeys  = @($Sections.Keys | Where-Object { $_ -match '^log_' })
    $trOpts   = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $trRx     = [regex]::new('<tr\b[^>]*>(.*?)</tr>', $trOpts)
    $tdRx     = [regex]::new('<td\b[^>]*>(.*?)</td>', $trOpts)
    $validLvl = @('Information','Warning','Error','Critical','Verbose')

    foreach ($key in $logKeys) {
        $html = $Sections[$key]
        foreach ($tr in $trRx.Matches($html)) {
            $cells = @($tdRx.Matches($tr.Groups[1].Value) |
                       ForEach-Object { Strip-HtmlTags $_.Groups[1].Value })
            # Event rows have exactly 4 cells: Time, Message, Level, Provider
            if ($cells.Count -eq 4 -and $cells[2].Trim() -in $validLvl) {
                $events.Add([PSCustomObject]@{
                    Time     = $cells[0].Trim()
                    Message  = $cells[1].Trim()
                    Level    = $cells[2].Trim()
                    Provider = $cells[3].Trim()
                }) | Out-Null
            }
        }
    }
    return $events
}

function Format-GB {
    param([double]$GB)
    if ($GB -lt 1.0) { return "$([math]::Round($GB * 1024, 0)) MB" }
    return "$($GB.ToString('F1')) GB"
}

function Format-HM {
    param([double]$Hours)
    if ($Hours -lt 1) { return "$([int][math]::Round($Hours * 60))m" }
    $h = [int][math]::Floor($Hours)
    $m = [int][math]::Round(($Hours - $h) * 60)
    if ($m -eq 0) { return "${h}h" }
    return "${h}h ${m}m"
}

function Get-GBValue {
    param([string]$Str)
    if (-not $Str) { return 0.0 }
    $m = [regex]::Match($Str, '([\d.]+)\s*(TB|GB|MB|KB|B)?', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return 0.0 }
    $val  = [double]$m.Groups[1].Value
    $unit = $m.Groups[2].Value.ToUpper()
    switch ($unit) {
        'TB' { return $val * 1024 }
        'GB' { return $val }
        'MB' { return $val / 1024 }
        'KB' { return $val / 1048576 }
        'B'  { return $val / 1073741824 }
        default { return $val }  # no unit - assume GB (legacy)
    }
}

function Parse-UtcDateTime {
    # Note: InternalInfo timestamps are device-local time, not UTC - returns Kind=Unspecified
    param([string]$Str)
    if (-not $Str) { return $null }
    $s = $Str -replace ' UTC','' -replace '\s+$',''
    $formats = @('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'M/d/yyyy h:mm:ss tt', 'M/d/yyyy H:mm:ss')
    foreach ($fmt in $formats) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParseExact($s, $fmt, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt  # Kind=Unspecified = device local time
        }
    }
    try { return [DateTime]::Parse($s) } catch { return $null }
}

# ============================================================================
# Device Analysis
# ============================================================================

function Analyze-Device {
    param(
        [PSCustomObject]$Device,
        [hashtable]$Sections,
        [hashtable]$LastSessions     = @{},
        [int]$RollbackWarnBelow      = 1,
        [int]$RollbackTarget         = 7,
        [int]$RollbackWarnAbove      = 30,
        [int]$ShadowAllocMinPct      = 10,
        [int]$MaxSnapPerVol          = 180
    )

    $aid         = $Device.AccountId
    $hasData     = ($Sections -and $Sections.Count -gt 0)
    $severity    = "GREEN"
    $issues      = [System.Collections.Generic.List[string]]::new()
    $remediations= [System.Collections.Generic.List[hashtable]]::new()

    # ?? Defaults when no data ??????????????????????????????????????????????
    $vssService = "Unknown"
    $writers    = @()
    $unhealthy  = @()
    $providers  = @()
    $thirdParty = @()
    $volumes    = @()
    $snapshots  = @()
    $manualRunning = @()
    $snapsTimeline = @{}
    $deviceTzLabel = 'local'
    $backupManagerVersion = $null
    $backupInProgress = $Device.T0 -in @(1, 9, 12)
    $s1ServicePresent = $false
    $s1Installed      = $false
    $detectedSoftware = [System.Collections.Generic.List[object]]::new()

    $secCatalog = @(
        # ---- Security / EDR ----
        @{ Cat='security'; Mfg='Microsoft';        Product='Defender / MDE';          SvcPat='windefend|wdnissvc|^sense$|mdcoresvc|mssecflt|webthreatdefsvc'; ProcPat='Windows Defender\\|MsSense\.exe|MsMpEng\.exe' }
        @{ Cat='security'; Mfg='SentinelOne';      Product='SentinelOne Agent';       SvcPat='sentinelagent|sentinelstaticengine|sentinelhelperservice'; ProcPat='SentinelOne\\|SentinelAgent\.exe|SentinelServiceHost\.exe'; PathExcl='SafeNet' }
        @{ Cat='security'; Mfg='CrowdStrike';      Product='Falcon Sensor';           SvcPat='csfalconservice|csagent|csc_vpnagent|csc_swgagent'; ProcPat='CrowdStrike\\|CSFalconService\.exe' }
        @{ Cat='security'; Mfg='Sophos';           Product='Sophos Endpoint';         SvcPat='sophos|savservice|sntpservice|sophisfim|hmpalertsvc'; ProcPat='Sophos\\|SavService\.exe|hmpalert\.exe' }
        @{ Cat='security'; Mfg='ESET';             Product='ESET Endpoint';           SvcPat='^ekrn$|^epfw$|epfwwfp|ehttpsrv|efwd|esetbridge'; ProcPat='ESET\\|ekrn\.exe|egui\.exe' }
        @{ Cat='security'; Mfg='Webroot';          Product='SecureAnywhere';          SvcPat='wrsvc|wrcoreservice|wrwtssvc'; ProcPat='Webroot\\|WRSA\.exe' }
        @{ Cat='security'; Mfg='Malwarebytes';     Product='Malwarebytes';            SvcPat='mbamservice|mbamprotection'; ProcPat='Malwarebytes\\|MBAMService\.exe' }
        @{ Cat='security'; Mfg='Cylance';          Product='CylancePROTECT';          SvcPat='cylancesvc'; ProcPat='Cylance\\|CylanceSvc\.exe' }
        @{ Cat='security'; Mfg='Carbon Black';     Product='CB Endpoint';             SvcPat='carbonblack|repmgr'; ProcPat='CarbonBlack\\|cb\.exe$|RepMgr\.exe' }
        @{ Cat='security'; Mfg='Trend Micro';      Product='Trend Micro';             SvcPat='tmbmsrv|ntrtscan|tmproxy'; ProcPat='Trend Micro\\|TMBMSRV\.exe|NTRtScan\.exe' }
        @{ Cat='security'; Mfg='Broadcom';         Product='Symantec Endpoint';       SvcPat='sepmasterservice|symantec endpoint protection'; ProcPat='Symantec\\|ccSvcHst\.exe|Smc\.exe' }
        @{ Cat='security'; Mfg='Trellix';          Product='Trellix / McAfee';        SvcPat='mcafeeframework|mfemms|mfevtp'; ProcPat='McAfee\\|Trellix\\|McShield\.exe' }
        @{ Cat='security'; Mfg='Fortinet';         Product='FortiClient';             SvcPat='fct_secsrv|forticlient|forticlientservice'; ProcPat='Fortinet\\|FcmDaemon\.exe|FortiTray\.exe' }
        @{ Cat='security'; Mfg='Palo Alto';        Product='Cortex XDR';             SvcPat='cyserver|cyvera|tlaservice'; ProcPat='Cortex XDR\\|cyserver\.exe|cyvera\.exe' }
        @{ Cat='security'; Mfg='ThreatLocker';     Product='ThreatLocker';            SvcPat='threatlockerservice'; ProcPat='ThreatLocker\\' }
        @{ Cat='security'; Mfg='Elastic';          Product='Elastic Security';        SvcPat='elasticendpoint'; ProcPat='Elastic\\' }
        @{ Cat='security'; Mfg='Bitdefender';      Product='Bitdefender';             SvcPat='bdredline|bdagent|epsecurityservice|bdnc'; ProcPat='Bitdefender\\' }
        @{ Cat='security'; Mfg='Huntress';         Product='Huntress Agent';          SvcPat='huntressagent|huntressupdater'; ProcPat='Huntress\\' }
        @{ Cat='security'; Mfg='Panda/WatchGuard'; Product='EPDR';                   SvcPat='psuaservice|nanoservicemain|nanosrv'; ProcPat='Panda\\|WatchGuard\\Endpoint' }
        @{ Cat='security'; Mfg='Kaseya';           Product='RocketCyber MDR';         SvcPat='^rocketagent$'; ProcPat='RocketAgent\\' }
        @{ Cat='security'; Mfg='Kaseya';           Product='Infocyte';                SvcPat='^huntagent$'; ProcPat='infocyte\\' }
        # ---- Backup & DR ----
        @{ Cat='backup';   Mfg='Veeam';            Product='Veeam Backup';            SvcPat='veeamdeploymentservice|veeamagentsvc|veeamendpointbackupsvc|veeamdeploysvc|veeamtransportsvc|veeammanagementagentsvc|veeamhvintegrationsvc'; ProcPat='Veeam\\' }
        @{ Cat='backup';   Mfg='Acronis';          Product='Acronis Backup';          SvcPat='acronisagent|acrsch2svc|acronisactiveprotectionservice'; ProcPat='Acronis\\' }
        @{ Cat='backup';   Mfg='Datto';            Product='Datto Backup';            SvcPat='dattobackupagentservice|dattocloudcontinuityservice'; ProcPat='DattoBackupAgent\.exe|DattoCloudContinuity\.exe' }
        @{ Cat='backup';   Mfg='Code42';           Product='Code42';                  SvcPat='code42service'; ProcPat='Code42\\' }
        @{ Cat='backup';   Mfg='Rubrik';           Product='Rubrik';                  SvcPat='rubrik'; ProcPat='RubrikVssProvider' }
        @{ Cat='backup';   Mfg='Unitrends';        Product='Unitrends';               SvcPat='unitrendsvc'; ProcPat='Unitrends\\' }
        @{ Cat='backup';   Mfg='Synology';         Product='Active Backup';           SvcPat='synology active backup|activebackupagent'; ProcPat='ActiveBackupAgentService\.exe' }
        @{ Cat='backup';   Mfg='Arcserve';         Product='Arcserve UDP';            SvcPat='casad2dwebsvc|casadserver|arcserveudp'; ProcPat='Arcserve\\|CA ARCserve\\' }
        @{ Cat='backup';   Mfg='Commvault';        Product='Commvault';               SvcPat='^gxcvd$|^gxfwd$|^clbackup$'; ProcPat='CommVault\\' }
        @{ Cat='backup';   Mfg='Microsoft';        Product='Windows Backup';          SvcPat='^wbengine$|^sdrsvc$'; ProcPat='wbengine\.exe' }
        @{ Cat='backup';   Mfg='Microsoft';        Product='Azure Backup';            SvcPat='^obengine$|mabextensionservice|windowsazurebackup'; ProcPat='Microsoft Azure Backup\\|MABWorker' }
        @{ Cat='backup';   Mfg='Veritas';          Product='Backup Exec';             SvcPat='backupexecjobengine|backupexecrpcservice|backupexecagentaccelerator'; ProcPat='Backup Exec\\|Veritas\\' }
        @{ Cat='backup';   Mfg='Druva';            Product='Druva inSync';            SvcPat='druvainsync|druvaphoenix'; ProcPat='Druva\\' }
        @{ Cat='backup';   Mfg='Barracuda';        Product='Barracuda Backup';        SvcPat='barracudabackup|barracudaagent'; ProcPat='Barracuda\\' }
        @{ Cat='backup';   Mfg='Zerto';            Product='Zerto';                   SvcPat='zertoservice|zerto'; ProcPat='Zerto\\' }
        @{ Cat='backup';   Mfg='Replibit';         Product='Replibit';                SvcPat='replibitagentservice|replibitmanagementservice|replibitUpdaterservice'; ProcPat='Replibit\\' }
        @{ Cat='backup';   Mfg='SQL Backup Master'; Product='SQL Backup Master';      SvcPat='sqlbackupmaster'; ProcPat='SQLBackupMaster\\' }
        # ---- RMM & Monitoring ----
        @{ Cat='rmm';      Mfg='N-able';           Product='N-central Agent';         SvcPat='windows agent|basupSrvcUpdater|nableupdateservice|n-able technologies windows agent'; ProcPat='N-able Technologies\\Windows Agent\\|bpagent\.exe|baagent\.exe' }
        @{ Cat='rmm';      Mfg='N-able';           Product='N-able Automation Mgr';  SvcPat='automationmanageragent'; ProcPat='AutomationManager\.AgentService\.exe' }
        @{ Cat='rmm';      Mfg='N-able';           Product='N-able PME';              SvcPat='pmesclient'; ProcPat='PME\.Agent\.exe|FileCacheServiceAgent\.exe|RequestHandlerAgent\.exe' }
        @{ Cat='rmm';      Mfg='ConnectWise';      Product='Automate (LabTech)';      SvcPat='^ltservice$|^ltsvcmon$'; ProcPat='LabTech\\|LTService\.exe|LTSvcMon\.exe' }
        @{ Cat='rmm';      Mfg='ConnectWise';      Product='SAAZ / Continuum';        SvcPat='saazserverplus|saazscheduler|saazremotesupport|saazwatchdog|itsplatformmanager'; ProcPat='SAAZ\\|DMPHelpDesk\.exe' }
        @{ Cat='rmm';      Mfg='Datto';            Product='Datto RMM';               SvcPat='^cagservice$'; ProcPat='CagService\.exe' }
        @{ Cat='rmm';      Mfg='NinjaRMM';         Product='NinjaOne';                SvcPat='ninjarmmagent|ninjavmagent'; ProcPat='NinjaRMMAgent\.exe|NinjaRMMAgentPatcher\.exe' }
        @{ Cat='rmm';      Mfg='Kaseya';           Product='Kaseya VSA';              SvcPat='kaseya agent|kaseyasnmptraphandler'; ProcPat='Kaseya\\|AgentMon\.exe' }
        @{ Cat='rmm';      Mfg='Atera';            Product='Atera RMM';               SvcPat='ateraagent'; ProcPat='AEMAgent\.exe|Atera\\' }
        @{ Cat='rmm';      Mfg='Syncro';           Product='Syncro';                  SvcPat='syncro|syncrolive|syncroovermind'; ProcPat='Syncro\\' }
        @{ Cat='rmm';      Mfg='Pulseway';         Product='Pulseway';                SvcPat='pulseway'; ProcPat='pbeagent\.exe|Pulseway\\' }
        @{ Cat='rmm';      Mfg='Auvik';            Product='Auvik Collector';         SvcPat='auvikcollector|auvikagent'; ProcPat='AuvikAgent|AuvikWatchdog' }
        @{ Cat='rmm';      Mfg='CyberCNS';         Product='CyberCNS';               SvcPat='cybercnsagentmonitor|cybercnsagent'; ProcPat='CyberCNS\\' }
        @{ Cat='rmm';      Mfg='Liongard';         Product='Liongard';                SvcPat='liongardagent'; ProcPat='LiongardAgent\.exe' }
        @{ Cat='rmm';      Mfg='Action1';          Product='Action1 RMM';             SvcPat='action1'; ProcPat='action1_agent\.exe' }
        @{ Cat='rmm';      Mfg='Microsoft';        Product='Intune / SCCM';           SvcPat='intunemanagementextension|healthtlservice'; ProcPat='HealthService\.exe' }
        @{ Cat='rmm';      Mfg='Splunk';           Product='Splunk Forwarder';        SvcPat='splunkforwarder|splunkd'; ProcPat='Splunk\\' }
        # ---- Remote Access ----
        @{ Cat='remote';   Mfg='TeamViewer';       Product='TeamViewer';              SvcPat='^teamviewer$|teamviewer_service'; ProcPat='TeamViewer\\|TeamViewer\.exe' }
        @{ Cat='remote';   Mfg='AnyDesk';          Product='AnyDesk';                 SvcPat='^anydesk$'; ProcPat='AnyDesk\.exe' }
        @{ Cat='remote';   Mfg='GoTo';             Product='LogMeIn / GoTo';          SvcPat='logmein|lmiguardiansvc|gotoassist'; ProcPat='LogMeIn\\|LMIGuardianSvc\.exe' }
        @{ Cat='remote';   Mfg='ConnectWise';      Product='Control (ScreenConnect)'; SvcPat='screenconnect'; ProcPat='ScreenConnect\\|ScreenConnect\.ClientService\.exe' }
        @{ Cat='remote';   Mfg='Splashtop';        Product='Splashtop';               SvcPat='splashtopremoteservice|srvctrl'; ProcPat='Splashtop\\|SRService\.exe' }
        @{ Cat='remote';   Mfg='UltraVNC';         Product='UltraVNC';                SvcPat='uvnc_service'; ProcPat='uvnc\\|winvnc\.exe' }
        @{ Cat='remote';   Mfg='SupRemo';          Product='SupRemo';                 SvcPat='supremoservice'; ProcPat='Supremo\.exe' }
        @{ Cat='remote';   Mfg='N-able';           Product='TakeControl (BeAnywhere)'; SvcPat='basupservice_n_central|basupstandaloneservice_n_central|basupexpressservice|^basupservice$'; ProcPat='BeAnywhere\\|BASupSrvc\.exe|BASupSrvcCnfg\.exe|BASupSrvcUpdater\.exe|GetSupportService' }
        @{ Cat='backup';   Mfg='N-able';           Product='Cove Backup Manager';      SvcPat='^backup service controller$'; ProcPat='BackupFP\.exe|Backup Manager\\BackupFP' }
    )

    if (-not $hasData) {
        $severity = "OFFLINE"
        $issues.Add("Device unreachable via diagnostics endpoint") | Out-Null
    } else {

        # ?? Device local time (used as reference for snapshot age - avoids UTC offset errors) ??
        $deviceNow = $null
        if ($Sections.ContainsKey('timeinfo')) {
            # Section uses <h3>Current time</h3><table><tr><td>value</td>... structure
            $opts = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            $m = [regex]::Match($Sections['timeinfo'],
                '<h3>\s*Current time\s*</h3>\s*<table[^>]*>.*?<td[^>]*>\s*([^<]+)\s*</td>', $opts)
            if ($m.Success) {
                $ctRaw = $m.Groups[1].Value.Trim()
                $dt = [datetime]::MinValue
                if ([datetime]::TryParseExact($ctRaw, 'yyyy-MMM-dd HH:mm:ss',
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
                    $deviceNow = $dt
                }
            }
            # Parse timezone label (e.g. "Central Daylight Time" -> "CDT")
            $tzM = [regex]::Match($Sections['timeinfo'],
                '<h3>\s*Time zone\s*</h3>\s*<table[^>]*>.*?<td[^>]*>\s*([^<]+)\s*</td>', $opts)
            if ($tzM.Success) {
                $deviceTzFull  = $tzM.Groups[1].Value.Trim()
                $deviceTzLabel = ($deviceTzFull -split '\s+' |
                    ForEach-Object { if ($_) { $_[0].ToString().ToUpper() } }) -join ''
            }
        }
        if (-not $deviceNow) { $deviceNow = [datetime]::Now }

        # ?? Backup Manager version ????????????????????????????????????????
        foreach ($sectionKey in @('accountprofile', 'appsettings')) {
            if ($backupManagerVersion -or -not $Sections.ContainsKey($sectionKey)) { continue }
            $rows = @(Parse-TableSection $Sections[$sectionKey])
            foreach ($row in $rows) {
                $settingKey = $row.PSObject.Properties.Name | Where-Object { $_ -match '^(Setting|Name|Key)$' } | Select-Object -First 1
                $valueKey   = $row.PSObject.Properties.Name | Where-Object { $_ -match '^(Value|Version)$' } | Select-Object -First 1
                if (-not $settingKey -or -not $valueKey) { continue }
                $setting = ($row.$settingKey ?? '').Trim()
                if ($setting -in @('LastSuccessfullyInitializedApplicationVersion', 'ProfileBackupManagerVersion')) {
                    $backupManagerVersion = ($row.$valueKey ?? '').Trim()
                    if ($backupManagerVersion) { break }
                }
            }
        }

        # ?? Hardware information (manufacturer, model, CPU, RAM) ????????????????????????
        $hwManufacturer = ''
        $hwModel = ''
        $hwCpu = ''
        $hwRam = ''
        # machineinfo section: single-column cells with "Key: Value" text
        if ($Sections.ContainsKey('machineinfo')) {
            $miHtml = $Sections['machineinfo']
            $miMatches = [regex]::Matches($miHtml, '<td[^>]*>([^<]+)</td>', 'IgnoreCase')
            foreach ($m in $miMatches) {
                $cell = $m.Groups[1].Value.Trim()
                if ($cell -match '^Manufacturer:\s*(.+)$') { $hwManufacturer = $Matches[1].Trim() }
                elseif ($cell -match '^Model:\s*(.+)$')        { $hwModel = $Matches[1].Trim() }
            }
        }
        # hardwareinfo section: RAM in first <td>, processor count in "Number of processors" column
        if ($Sections.ContainsKey('hardwareinfo')) {
            $hiHtml = $Sections['hardwareinfo']
            # RAM: first <td> value before any thead (appears before "Processor info" table)
            if ($hiHtml -match '<td[^>]*>\s*([\d\.]+ [KMGT]B)\s*</td>') { $hwRam = $Matches[1].Trim() }
            # Processor count: row after Architecture/Number header
            if ($hiHtml -match 'Number of processors[\s\S]*?<td>[^<]+</td>\s*<td>(\d+)</td>') {
                $hwCpu = "$($Matches[1]) cores"
            }
        }

        # ?? VSS service status ?????????????????????????????????????????????
        if ($Sections.ContainsKey('vssservicestatus')) {
            $text = (Strip-HtmlTags $Sections['vssservicestatus']).Trim()
            if ($text) { $vssService = $text }
        }

        # ?? VSS Writers (individual vsswriter_* dl sections, matching server.py _kv_pairs) ?
        $writerSectionKeys = @($Sections.Keys | Where-Object { $_ -match '^vsswriter_' })
        foreach ($wsKey in $writerSectionKeys) {
            $kv      = Parse-DlSection $Sections[$wsKey]
            # Use the Name field for display; fall back to section key suffix
            $name    = if ($kv['Name']) { $kv['Name'].Trim() } else { $wsKey.Substring('vsswriter_'.Length) }
            $status  = if ($null -ne $kv['Status']) { $kv['Status'].Trim() }
                       elseif ($null -ne $kv['State']) { $kv['State'].Trim() }
                       else { $null }
            $err     = if ($kv['Last error']) { $kv['Last error'].Trim() } else { '' }
            $usage   = if ($kv['Usage type']) { $kv['Usage type'].Trim() } else { '' }
            $writers += @{ name = $name; status = if ($status) { $status } else { '' }; error = $err; usage = $usage }
            # Unhealthy: status present and not 'stable', OR error present and not a healthy value
            $statusLow = if ($status) { $status.ToLower() } else { $null }
            $errLow    = $err.ToLower()
            # All "Waiting for *" states are normal VSS lifecycle transients during a backup.
            # InternalInfo captures writer state at collection time - a writer mid-backup shows
            # "Waiting for freeze/thaw/etc." which is healthy. Only flag if error code is also bad.
            # "Waiting for Post-snapshot" outside a backup may indicate a stuck snapshot set,
            # but that is caught by the SWPRV/VSS service suspicious detector instead.
            $waitingForBackup = $statusLow -match '^waiting for'
            $stateUnhealthy = ($null -ne $statusLow) -and
                              ($statusLow -ne 'stable') -and
                              (-not $waitingForBackup)
            $errUnhealthy   = ($errLow -ne '') -and ($errLow -notin @('no error', '0x00000000'))
            if ($stateUnhealthy -or $errUnhealthy) {
                $unhealthy += @{ name = $name; status = if ($status) { $status } else { 'Unknown' }; error = $err }
            }
        }
        # Detect SentinelOne writer presence - used later for snapshot tool classification
        # Match 'SentinelOne', 'Sentinel One', or 'Sentinel Agent' (S1 VSS writer names); exclude SafeNet Sentinel (DRM)
        $s1WriterPresent = [bool]($writers | Where-Object { $_.name -match 'sentinel[\s-]?one|sentinel\s+agent' })

        # VSS writer -> host service mapping
        # Writers hosted inside vssvc.exe only need a VSS restart.
        # Writers with an external service need that service restarted too.
        $WriterHostSvc = @{
            'ASR Writer'                      = @('vss',         $null)
            'BITS Writer'                     = @('BITS',        'BITS')
            'COM+ REGDB Writer'               = @('EventSystem', 'COM+ Event System')
            'IIS Config Writer'               = @('iisadmin',    'IIS Admin')
            'IIS Metabase Writer'             = @('iisadmin',    'IIS Admin')
            'Microsoft Hyper-V VSS Writer'    = @('vmms',        'Hyper-V VM Management')
            'Performance Counters Writer'     = @('vss',         $null)
            'Registry Writer'                 = @('vss',         $null)
            'Shadow Copy Optimization Writer' = @('vss',         $null)
            'SqlServerWriter'                 = @('sqlwriter',   'SQL Server VSS Writer')
            'System Writer'                   = @('cryptsvc',    'Cryptographic Services')
            'Task Scheduler Writer'           = @('Schedule',    'Task Scheduler')
            'VSS Metadata Store Writer'       = @('vss',         $null)
            'WIDWriter'                       = @('MSSQL$MICROSOFT##WID', 'Windows Internal Database')
            'WMI Writer'                      = @('winmgmt',     'WMI')
        }

        function Get-WriterHostSvc {
            param([string]$Name)
            foreach ($k in $WriterHostSvc.Keys) {
                if ($Name -like "*$k*") { return $WriterHostSvc[$k] }
            }
            if ($Name -match 'sentinel[\s-]?one|sentinel\s+agent') { return @('SentinelAgent', 'SentinelOne Agent') }
            return @('vss', $null)  # default: VSS-hosted
        }

        # Enrich each writer with its host service name (for the writers table service column)
        foreach ($w in $writers) {
            $pair = Get-WriterHostSvc $w.name
            $w['service_name'] = $pair[0]
        }

        # Issue per unhealthy writer; ONE consolidated remediation for all of them
        foreach ($uw in $unhealthy) {
            $severity = "RED"
            $err = if ($uw.error) { $uw.error } else { $uw.status }
            $issues.Add("Writer '$($uw.name)' -> $($uw.status) ($err)") | Out-Null
        }

        if ($unhealthy.Count -gt 0) {
            # Collect external (non-VSS-hosted) services that need restarting
            $extSvcs = [ordered]@{}  # svc -> label
            foreach ($uw in $unhealthy) {
                $pair = Get-WriterHostSvc $uw.name
                $svc  = $pair[0]; $lbl = $pair[1]
                if ($svc -ne 'vss' -and -not $extSvcs.Contains($svc)) {
                    $extSvcs[$svc] = if ($lbl) { $lbl } else { $svc }
                }
            }

            $writerNamesStr = ($unhealthy | ForEach-Object { $_.name }) -join ', '
            $cmds = @("# Failing writers: $writerNamesStr", "vssadmin list writers", "")
            $cmds += "# Stop: SQL writer first (hard dependency), then other external services, then VSS"
            if ($extSvcs.Contains('sqlwriter')) { $cmds += 'net stop sqlwriter' }
            foreach ($svc in $extSvcs.Keys) {
                if ($svc -ne 'sqlwriter') { $cmds += "net stop $svc" }
            }
            $cmds += @('net stop vss', 'net start vss')
            foreach ($svc in $extSvcs.Keys) {
                if ($svc -ne 'sqlwriter') { $cmds += "net start $svc" }
            }
            if ($extSvcs.Contains('sqlwriter')) {
                $cmds += @('net start sqlwriter', 'vssadmin delete shadows /all /quiet')
            }
            $cmds += @('', '# Verify all writers recovered:', 'vssadmin list writers')

            $remediations.Add(@{
                sev   = 'CRITICAL'
                title = "VSS writer failure ($($unhealthy.Count) writer(s))"
                cmds  = $cmds
            }) | Out-Null
        }

        # ?? VSS Providers ??????????????????????????????????????????????????
        if ($Sections.ContainsKey('vssproviders')) {
            $providers = @(Parse-TableSection $Sections['vssproviders'])
            $thirdParty = @($providers | Where-Object {
                $n = $_.Name ?? $_.'Provider Name' ?? ''
                $n -notmatch 'Microsoft|Hyper-V IC|File Share'
            })
        }

        foreach ($p in $thirdParty) {
            if ($severity -eq "GREEN") { $severity = "YELLOW" }
            $issues.Add("Third-party VSS provider: $($p.Name) - potential conflict") | Out-Null
        }

        # ?? Shadow Storages ????????????????????????????????????????????????
        $storages = @()
        if ($Sections.ContainsKey('vssshadowstorages')) {
            $storages = @(Parse-TableSection $Sections['vssshadowstorages'])
        }

        foreach ($st in $storages) {
            $volLabel = ($st.Volume ?? $st.'Volume Name' ?? '?').Trim()
            $usedKey  = $st.PSObject.Properties.Name | Where-Object { $_ -match 'Used' }        | Select-Object -First 1
            $allocKey = $st.PSObject.Properties.Name | Where-Object { $_ -match 'Alloc' }       | Select-Object -First 1
            $maxKey   = $st.PSObject.Properties.Name | Where-Object { $_ -match 'Max|Maximum' } | Select-Object -First 1
            # "Shadow copy volume" column contains the GUID volume path that matches
            # "Original volume" in the vsssnapshots table (e.g. \\?\Volume{guid}\)
            $guidKey  = $st.PSObject.Properties.Name | Where-Object { $_ -match 'Shadow copy volume' } | Select-Object -First 1
            $volGuid  = if ($guidKey) { ($st.$guidKey).Trim() } else { '' }

            $maxRaw    = if ($maxKey) { ($st.$maxKey ?? '').Trim() } else { '' }
            $unbounded = $maxRaw -match 'UNBOUNDED|unlimited' -or (-not $maxKey)
            $usedGB    = Get-GBValue ($st.$usedKey)
            $allocGB   = if ($allocKey) { Get-GBValue ($st.$allocKey) } else { 0.0 }
            $maxGB     = Get-GBValue $maxRaw
            $pctUsed   = if ($maxGB -gt 0) { [math]::Round($usedGB / $maxGB * 100, 1) } else { 0 }

            $volEntry = @{
                drive       = $volLabel
                guid        = $volGuid
                used_gb     = $usedGB
                alloc_gb    = $allocGB
                max_gb      = $maxGB
                unbounded   = $unbounded
                pct_used    = $pctUsed
                s1_count    = 0
                s1_coverage_days = 0.0
                s1_newest   = $null
                s1_oldest   = $null
                behind_schedule = $false
                minutes_since_newest = 0.0
                count        = 0
                cove_count   = 0
                native_count = 0
                writer_count = 0
                other_count   = 0
                unknown_count = 0
                orphan_count  = 0
                orphan_ids    = [System.Collections.Generic.List[string]]::new()
                orphan_oldest_time = $null
                disk_total_gb = 0.0
                alloc_pct     = 0.0
                s1_median_interval_h = $null
            }
            $volumes += $volEntry
            # Shadow storage checks are deferred to after snapshot classification
            # so we know whether S1/native snapshots exist on each volume.
        }

        # ?? Disk configuration (drive sizes -> shadow storage allocation % check) ??????????????
        # S1 recommends shadow storage allocation ?10% of the protected volume.
        # Parse diskconfiguration JSON to get each drive letter's total size.
        $diskSizeByDrive  = @{}
        $volumeGuidToDrive = @{}   # normalised GUID (lowercase, no braces) -> "C:\"
        if ($Sections.ContainsKey('diskconfiguration')) {
            $opts = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            $thM = [regex]::Match($Sections['diskconfiguration'], '<div[^>]*class="jjson"[^>]*>(.*?)</div>', $opts)
            if ($thM.Success) {
                $jsonRaw = [System.Net.WebUtility]::HtmlDecode($thM.Groups[1].Value.Trim())
                try {
                    $diskData = $jsonRaw | ConvertFrom-Json -ErrorAction Stop
                    foreach ($disk in $diskData.DiskArray) {
                        foreach ($part in $disk.PartitionArray) {
                            $dl = ($part.driveLetter ?? '').Trim()
                            $driveKey = if ($dl) { ($dl.TrimEnd('\') + '\').ToUpper() } else { '' }

                            # Size by drive letter
                            if ($driveKey -and $part.length -gt 0) {
                                $gb = [math]::Round($part.length / 1073741824.0, 1)
                                if (-not $diskSizeByDrive.ContainsKey($driveKey) -or $gb -gt $diskSizeByDrive[$driveKey]) {
                                    $diskSizeByDrive[$driveKey] = $gb
                                }
                            }

                            # GUID -> drive letter lookup (from deviceName: \\??\Volume{guid} or \\?\Volume{guid}\)
                            $devName = ($part.deviceName ?? '').Trim()
                            $guidM = [regex]::Match($devName, '\{([0-9a-fA-F\-]{36})\}')
                            if ($guidM.Success -and $driveKey) {
                                $volumeGuidToDrive[$guidM.Groups[1].Value.ToLower()] = $driveKey
                            }
                            # Also check deviceNamesArray entries
                            if ($part.deviceNamesArray) {
                                foreach ($dn in $part.deviceNamesArray) {
                                    $gm2 = [regex]::Match($dn, '\{([0-9a-fA-F\-]{36})\}')
                                    if ($gm2.Success -and $driveKey -and -not $volumeGuidToDrive.ContainsKey($gm2.Groups[1].Value.ToLower())) {
                                        $volumeGuidToDrive[$gm2.Groups[1].Value.ToLower()] = $driveKey
                                    }
                                }
                            }
                        }
                    }
                } catch {}
            }
        }

        # Back-fill disk_total_gb and alloc_pct on each volume entry (checks deferred to after snapshots)
        foreach ($ve in $volumes) {
            $key = (($ve.drive -replace '[/\\]+$', '') + '\').ToUpper()
            $diskGB = if ($diskSizeByDrive.ContainsKey($key)) { $diskSizeByDrive[$key] } else { 0.0 }
            $ve.disk_total_gb = $diskGB
            $ve.alloc_pct = if ($diskGB -gt 0 -and $ve.max_gb -gt 0) {
                [math]::Round($ve.max_gb / $diskGB * 100, 1)
            } else { 0.0 }
        }

        # ?? Snapshots (classify by attribute pattern - matches server.py _classify_snap_tool) ?
        # s1     = writers were invoked (no "no writers") + SentinelOne writer registered
        # writer = writers were invoked but no S1 writer detected
        # cove   = no writers + client accessible + auto recover  (crash-consistent, active session)
        # native = no writers + client accessible + no auto release  (WSB / Veeam / etc.)
        # other  = everything else
        # orphan = other + no auto release absent (unmanaged leftover)
        if ($Sections.ContainsKey('vsssnapshots')) {
            $snapRows = @(Parse-TableSection $Sections['vsssnapshots'])
            $timeKey  = if ($snapRows.Count -gt 0) {
                $snapRows[0].PSObject.Properties.Name |
                    Where-Object { $_ -match 'Creation|Time|Date' } | Select-Object -First 1
            }

            # Build GUID and drive-letter maps for matching snapshots to volumes.
            # InternalInfo "Original volume" is a raw GUID path: \\?\Volume{guid}\
            # (no drive-letter prefix) - must resolve via shadow storage GUID map.
            $volByGuid  = @{}
            $volByDrive = @{}
            foreach ($v in $volumes) {
                if ($v.guid)  { $volByGuid[$v.guid]   = $v }
                if ($v.drive) { $volByDrive[$v.drive] = $v }
            }

            # Group snapshots by resolved drive letter
            $snapsByDrive = @{}  # drive -> {s1,cove,native,writer,other,orphan,all}
            foreach ($sn in $snapRows) {
                $attrs      = ($sn.Attributes ?? $sn.Type ?? $sn.SnapshotType ?? '').Trim().ToLower()
                $hasWriters = $attrs -notmatch 'no writers'
                $hasClient  = $attrs -match 'client accessible'
                $hasNar     = $attrs -match 'no auto release'
                $hasAr      = $attrs -match 'auto recover'

                $snapTool = if ($hasWriters) {
                    # S1: writer registered + no auto release. Without the S1 writer present,
                    # no-auto-release snaps are from another tool - classify as 'writer' to avoid false positives.
                    if ($hasNar -and $s1WriterPresent) { 's1' }
                    elseif ($hasNar)   { 'writer' }
                    elseif ($hasAr)    { 'unknown' }
                    else               { 'writer' }
                } elseif ($hasClient) {
                    if ($hasAr)      { 'cove' }
                    elseif ($hasNar) { 'native' }
                    else             { 'other' }
                } else { 'other' }

                $isOrphan = ($snapTool -eq 'other' -and -not $hasNar)
                $origVol  = ($sn.'Original volume' ?? $sn.'OriginalVolume' ?? '').Trim()
                # Resolve drive: GUID match first, then "(C:)\..." regex fallback
                $driveLet = if ($volByGuid.ContainsKey($origVol)) {
                    $volByGuid[$origVol].drive
                } elseif ($origVol -match '\(([A-Z]:\\)') {
                    $Matches[1]
                } else { '_' }

                if (-not $snapsByDrive.ContainsKey($driveLet)) {
                    $snapsByDrive[$driveLet] = @{
                        s1      = [System.Collections.Generic.List[object]]::new()
                        cove    = [System.Collections.Generic.List[object]]::new()
                        native  = [System.Collections.Generic.List[object]]::new()
                        writer  = [System.Collections.Generic.List[object]]::new()
                        other   = [System.Collections.Generic.List[object]]::new()
                        unknown = [System.Collections.Generic.List[object]]::new()
                        orphan  = [System.Collections.Generic.List[object]]::new()
                        all     = [System.Collections.Generic.List[object]]::new()
                    }
                }
                $snapsByDrive[$driveLet].all.Add($sn)        | Out-Null
                $snapsByDrive[$driveLet][$snapTool].Add($sn) | Out-Null
                if ($isOrphan) { $snapsByDrive[$driveLet].orphan.Add($sn) | Out-Null }
            }

            # Update each volume entry with per-volume snapshot counts
            $allS1Snaps = [System.Collections.Generic.List[object]]::new()
            foreach ($drv in $snapsByDrive.Keys) {
                $grp = $snapsByDrive[$drv]
                $ve  = $volByDrive[$drv]
                if (-not $ve) {
                    if ($drv -eq '_' -and $volumes.Count -gt 0) { $ve = $volumes[0] }
                    else {
                        $ve = @{ drive = $drv; used_gb = 0.0; alloc_gb = 0.0; max_gb = 0.0; unbounded = $false; pct_used = 0.0;
                                 s1_count = 0; s1_coverage_days = 0.0; s1_newest = $null; s1_oldest = $null;
                                 behind_schedule = $false; minutes_since_newest = 0.0;
                                 count = 0; cove_count = 0; native_count = 0; writer_count = 0;
                                 other_count = 0; unknown_count = 0; orphan_count = 0; orphan_ids = [System.Collections.Generic.List[string]]::new(); orphan_oldest_time = $null; disk_total_gb = 0.0; alloc_pct = 0.0; s1_median_interval_h = $null }
                        $volumes += $ve; $volByDrive[$drv] = $ve
                    }
                }
                $ve.count        = $grp.all.Count
                $ve.s1_count     = $grp.s1.Count
                $ve.cove_count   = $grp.cove.Count
                $ve.native_count = $grp.native.Count
                $ve.writer_count = $grp.writer.Count
                $ve.other_count   = $grp.other.Count
                $ve.unknown_count = $grp.unknown.Count
                $ve.orphan_count  = $grp.orphan.Count
                $ve.orphan_ids        = [System.Collections.Generic.List[string]]::new()
                $ve.orphan_oldest_time = $null
                if ($grp.orphan.Count -gt 0) {
                    # Identify Shadow Copy ID column - vssadmin output typically names it "Shadow Copy ID"
                    $idColName = if ($grp.orphan[0]) {
                        $grp.orphan[0].PSObject.Properties.Name |
                            Where-Object { $_ -match 'Shadow Copy ID|ShadowCopyId|^ID$' } |
                            Select-Object -First 1
                    } else { $null }
                    $orphTimes = [System.Collections.Generic.List[datetime]]::new()
                    foreach ($orphSnap in $grp.orphan) {
                        if ($idColName) {
                            $m = [regex]::Match(($orphSnap.$idColName ?? ''), '\{[0-9a-fA-F\-]{36}\}')
                            if ($m.Success) { $ve.orphan_ids.Add($m.Value) | Out-Null }
                        }
                        if ($timeKey) {
                            $ot = Parse-UtcDateTime ($orphSnap.$timeKey ?? '')
                            if ($ot) { $orphTimes.Add($ot) | Out-Null }
                        }
                    }
                    if ($orphTimes.Count -gt 0) {
                        $ve.orphan_oldest_time = ($orphTimes | Measure-Object -Minimum).Minimum
                    }
                }

                if ($grp.s1.Count -gt 0 -and $timeKey) {
                    $grp.s1 | ForEach-Object { $allS1Snaps.Add($_) | Out-Null }
                    $s1Times = @($grp.s1 | ForEach-Object { Parse-UtcDateTime ($_.$timeKey) } | Where-Object { $_ })
                    if ($s1Times.Count -gt 0) {
                        $oldest = ($s1Times | Measure-Object -Minimum).Minimum
                        $newest = ($s1Times | Measure-Object -Maximum).Maximum
                        $now    = $deviceNow
                        $ve.s1_coverage_days     = [math]::Round(($newest - $oldest).TotalDays, 1)
                        $ve.s1_oldest            = $oldest.ToString('yyyy-MM-dd HH:mm') + " $deviceTzLabel"
                        $ve.s1_newest            = $newest.ToString('yyyy-MM-dd HH:mm') + " $deviceTzLabel"
                        $ve.minutes_since_newest = [math]::Round(($now - $newest).TotalMinutes, 1)
                        # behind_schedule is set after median is computed below; placeholder
                        $ve.behind_schedule      = $false

                        # Median adjacent interval - robust cadence estimate unaffected by outliers
                        if ($s1Times.Count -ge 2) {
                            $sorted    = $s1Times | Sort-Object
                            $intervals = for ($i = 1; $i -lt $sorted.Count; $i++) {
                                ($sorted[$i] - $sorted[$i-1]).TotalHours
                            }
                            $sortedIntervals = @($intervals | Sort-Object)
                            $mid = [int][math]::Floor($sortedIntervals.Count / 2)
                            $median = if ($sortedIntervals.Count % 2 -eq 0) {
                                ($sortedIntervals[$mid - 1] + $sortedIntervals[$mid]) / 2
                            } else { $sortedIntervals[$mid] }
                            $ve.s1_median_interval_h = [math]::Round($median, 1)
                        }
                        # Set behind_schedule using cadence-aware threshold (1.5x median, min 4h)
                        $volGraceH = if ($ve.s1_median_interval_h -and $ve.s1_median_interval_h -gt 0) {
                            [math]::Max(4.0, $ve.s1_median_interval_h * 1.5)
                        } else { 4.0 }
                        $ve.behind_schedule = ($ve.minutes_since_newest / 60.0) -gt $volGraceH
                    }
                }
            }

            # ?? Per-volume shadow storage checks (deferred until snapshot counts are known) ????????
            foreach ($ve in $volumes) {
                $volLabel = $ve.drive
                $hasSnaps = ($ve.s1_count + $ve.native_count + $ve.cove_count) -gt 0
                if ($ve.unbounded) {
                    # UNBOUNDED max: VSS can consume all available disk space - set an explicit cap
                    if ($hasSnaps) {
                        if ($severity -eq 'GREEN') { $severity = 'YELLOW' }
                        $issues.Add("Shadow max UNBOUNDED on ${volLabel}: VSS can consume all free disk space. Set an explicit cap (recommended: 10% of volume).") | Out-Null
                        $remediations.Add(@{
                            sev  = "WARN"; title = "Shadow storage cap not set on $volLabel"
                            cmds = @("vssadmin resize shadowstorage /for=${volLabel} /on=${volLabel} /maxsize=10%")
                        }) | Out-Null
                    }
                } else {
                    # Fixed cap: check allocation vs disk and usage vs cap
                    if ($ve.alloc_pct -gt 0 -and $ve.alloc_pct -lt $ShadowAllocMinPct) {
                        # Allocation below configured minimum (default 10% per Microsoft/S1 recommendation)
                        if ($severity -eq 'GREEN') { $severity = 'YELLOW' }
                        $issues.Add("Shadow allocation low on ${volLabel}: $($ve.alloc_pct)% of disk ($($ve.max_gb.ToString('F1')) GB cap / $($ve.disk_total_gb.ToString('F1')) GB volume) - minimum is $ShadowAllocMinPct%") | Out-Null
                        $remediations.Add(@{
                            sev  = "WARN"; title = "Shadow storage allocation low on $volLabel"
                            cmds = @("vssadmin resize shadowstorage /for=${volLabel} /on=${volLabel} /maxsize=${ShadowAllocMinPct}%")
                        }) | Out-Null
                    } elseif ($ve.pct_used -ge 90 -and $hasSnaps) {
                        # Cap nearly full but allocation is adequate - S1 is actively managing expiry, expected behavior
                        $issues.Add("[i]${volLabel} shadow: $($ve.pct_used)% of cap used ($($ve.used_gb.ToString('F1')) / $($ve.max_gb.ToString('F1')) GB) - S1 managing expiry, allocation $($ve.alloc_pct)% of disk") | Out-Null
                    }
                }
                # Snapshot count check (all types combined)
                if ($ve.count -gt $MaxSnapPerVol) {
                    if ($severity -eq 'GREEN') { $severity = 'YELLOW' }
                    $issues.Add("${volLabel} has $($ve.count) snapshots (limit: $MaxSnapPerVol) - excessive snapshot accumulation may impact VSS performance") | Out-Null
                }
            }

            # Fleet-level S1 health checks (aggregate across all volumes for issue flags)
            if ($allS1Snaps.Count -gt 0 -and $timeKey) {
                $allS1Times = @($allS1Snaps | ForEach-Object { Parse-UtcDateTime ($_.$timeKey) } | Where-Object { $_ })
                if ($allS1Times.Count -gt 0) {
                    $oldest         = ($allS1Times | Measure-Object -Minimum).Minimum
                    $newest         = ($allS1Times | Measure-Object -Maximum).Maximum
                    $now            = $deviceNow
                    $coverageDays   = [math]::Round(($newest - $oldest).TotalDays, 1)
                    $minSinceNewest = [math]::Round(($now - $newest).TotalMinutes, 1)

                    # Compute median cadence from fleet-wide S1 times so the behind threshold is cadence-aware
                    $fleetMedianIntervalH = $null
                    if ($allS1Times.Count -ge 3) {
                        $sortedAll = @($allS1Times | Sort-Object)
                        $fleetIntervals = for ($fi = 1; $fi -lt $sortedAll.Count; $fi++) {
                            ($sortedAll[$fi] - $sortedAll[$fi-1]).TotalHours
                        }
                        $sortedFI = @($fleetIntervals | Sort-Object)
                        $fMid = [int][math]::Floor($sortedFI.Count / 2)
                        $fleetMedianIntervalH = if ($sortedFI.Count % 2 -eq 0) {
                            ($sortedFI[$fMid - 1] + $sortedFI[$fMid]) / 2
                        } else { $sortedFI[$fMid] }
                    }
                    # Grace window = 1.5x cadence, minimum 4h; anything beyond is "behind"
                    $behindThreshH  = if ($fleetMedianIntervalH -and $fleetMedianIntervalH -gt 0) {
                        [math]::Max(4.0, $fleetMedianIntervalH * 1.5)
                    } else { 4.0 }
                    $behindSchedule = ($minSinceNewest / 60.0) -gt $behindThreshH

                    if ($allS1Times.Count -gt 1) {
                        if ($coverageDays -lt $RollbackWarnBelow) {
                            if ($severity -eq 'GREEN') { $severity = 'YELLOW' }
                            $issues.Add("S1 rollback window: $coverageDays days - near zero. Shadow storage pressure may be expiring snapshots faster than S1 creates them.") | Out-Null
                        } elseif ($coverageDays -lt $RollbackTarget) {
                            $issues.Add("[i]S1 rollback: $coverageDays days - reduced but functional (target >= $RollbackTarget days)") | Out-Null
                        } elseif ($coverageDays -gt $RollbackWarnAbove) {
                            if ($severity -eq 'GREEN') { $severity = 'YELLOW' }
                            $issues.Add("S1 rollback: $coverageDays days - possibly excessive retention; shadow storage may be carrying more history than needed (threshold: $RollbackWarnAbove days)") | Out-Null
                        }
                        # RollbackTarget..RollbackWarnAbove: healthy range, no alert
                    } elseif ($allS1Times.Count -eq 1) {
                        if ($severity -eq "GREEN") { $severity = "YELLOW" }
                        $issues.Add("Only 1 S1 snapshot - coverage cannot be established") | Out-Null
                    }

                    if ($behindSchedule) {
                        $hrsLate  = [math]::Round($minSinceNewest / 60, 1)
                        # RED if overdue by more than 1x cadence beyond the grace window (or >16h absolute)
                        $redThreshH = if ($fleetMedianIntervalH -and $fleetMedianIntervalH -gt 0) {
                            [math]::Max(8.0, $fleetMedianIntervalH * 2.5)
                        } else { 8.0 }
                        if ($hrsLate -gt $redThreshH) {
                            $severity = "RED"
                            $issues.Add("S1 severely behind schedule: $(Format-HM $hrsLate) since last snapshot (expected cadence ~$(Format-HM $behindThreshH))") | Out-Null
                            $remediations.Add(@{
                                sev  = "CRITICAL"; title = "SentinelOne stale"
                                cmds = @("# Check SentinelOne agent status in S1 console","# If agent is degraded, reinstall or re-register")
                            }) | Out-Null
                        } else {
                            if ($severity -eq "GREEN") { $severity = "YELLOW" }
                            $issues.Add("S1 behind schedule: $(Format-HM $hrsLate) since last snapshot (expected cadence ~$(Format-HM $behindThreshH))") | Out-Null
                        }
                    }
                }
            } elseif ($snapRows.Count -gt 0 -and $s1WriterPresent) {
                if ($severity -eq "GREEN") { $severity = "YELLOW" }
                $issues.Add("No S1 (writer-based) snapshots found - rollback protection unknown") | Out-Null
            }
        }

        # ?? Services (detect manual VSS services running) ??????????????????
        $svcMap = @{}
        if ($Sections.ContainsKey('servicesinfo')) {
            $svcRows = @(Parse-TableSection $Sections['servicesinfo'])
            $manualVssSvcs = @('swprv','vss','sqlwriter')
            foreach ($svc in $svcRows) {
                $nameKey    = $svc.PSObject.Properties.Name | Where-Object { $_ -match '^Name$' } | Select-Object -First 1
                $stateKey   = $svc.PSObject.Properties.Name | Where-Object { $_ -match 'Status|State' } | Select-Object -First 1
                $startupKey = $svc.PSObject.Properties.Name | Where-Object { $_ -match 'Startup|Start' } | Select-Object -First 1
                $pathKey    = $svc.PSObject.Properties.Name | Where-Object { $_ -match '^Path$' } | Select-Object -First 1

                $svcName    = ($svc.$nameKey ?? '').Trim().ToLower()
                $svcState   = ($svc.$stateKey ?? '').Trim()
                $svcStartup = ($svc.$startupKey ?? '').Trim()
                $svcPath    = ($svc.$pathKey ?? '').Trim()

                if ($svcName) { $svcMap[$svcName] = @{ state = $svcState; startup = $svcStartup } }

                # Match SentinelOne services specifically; exclude SafeNet Sentinel (DRM licensing)
                if ($svcName -match 'sentinel[\s-]?one|^sentinelagent$|^sentinelstaticengine$|^sentinelhelperservice$' -and $svcPath -notmatch 'SafeNet') { $s1ServicePresent = $true }
                if ($svcName -in $manualVssSvcs -and $svcState -match 'Running') {
                    $manualRunning += $svcName.ToUpper()
                }
                # Software detection catalog match
                foreach ($tool in $secCatalog) {
                    $excl = $tool.PathExcl
                    if ($svcName -match $tool.SvcPat -and (-not $excl -or $svcPath -notmatch $excl)) {
                        $detectedSoftware.Add([PSCustomObject]@{
                            cat = $tool.Cat; mfg = $tool.Mfg; product = $tool.Product
                            via = 'Service'; name = $svc.$nameKey; state = $svcState
                        }) | Out-Null
                    }
                }
            }
        }
        $s1Installed = $s1WriterPresent -or $s1ServicePresent

        # ?? Processes (software detection only) ??????????????????????????????????????
        if ($Sections.ContainsKey('processesinfo')) {
            $procRows = @(Parse-TableSection $Sections['processesinfo'])
            if ($procRows.Count -gt 0) {
                $procNameKey = $procRows[0].PSObject.Properties.Name | Where-Object { $_ -match '^Name$' } | Select-Object -First 1
                $procPathKey = $procRows[0].PSObject.Properties.Name | Where-Object { $_ -match '^Path$' } | Select-Object -First 1
                $seenProcTools = [System.Collections.Generic.HashSet[string]]::new()
                foreach ($proc in $procRows) {
                    $pName = ($proc.$procNameKey ?? '').Trim()
                    $pPath = ($proc.$procPathKey ?? '').Trim()
                    foreach ($tool in $secCatalog) {
                        $key = "$($tool.Mfg)|$($tool.Product)"
                        if ($seenProcTools.Contains($key)) { continue }
                        $excl = $tool.PathExcl
                        if (($pName -match $tool.ProcPat -or $pPath -match $tool.ProcPat) -and (-not $excl -or $pPath -notmatch $excl)) {
                            # Only add as Process if not already detected via Service
                            $alreadyViaSvc = [bool]($detectedSoftware | Where-Object { $_.mfg -eq $tool.Mfg -and $_.product -eq $tool.Product })
                            if (-not $alreadyViaSvc) {
                                $detectedSoftware.Add([PSCustomObject]@{
                                    cat = $tool.Cat; mfg = $tool.Mfg; product = $tool.Product
                                    via = 'Process'; name = $pName; state = 'Running'
                                }) | Out-Null
                            }
                            $seenProcTools.Add($key) | Out-Null
                        }
                    }
                }
            }
        }

        # ?? Event log signals (VolSnap / VSS events from last backup session) ??
        # VolSnap Event 36: "oldest shadow copy deleted...below the user defined limit" = normal
        #   storage management working. Counter-evidence against "hung" - system is deleting old
        #   shadows to stay under limit, which is expected healthy behaviour.
        # VolSnap Event 14: "insufficient disk space...to grow the shadow copy storage" = storage
        #   exhaustion. Primary cause of stale snapshots; VSS services may appear "hung" because
        #   VSS gave up trying to create a snapshot that storage could not accommodate.
        $evtLogs            = if ($hasData) { @(Parse-EventLogs $Sections) } else { @() }
        $volsnapDeleteCount = @($evtLogs | Where-Object {
            $_.Provider -eq 'Volsnap' -and
            $_.Message  -match 'oldest shadow copy.*deleted.*below the user defined limit|oldest shadow copy.*deleted.*below a limit'
        }).Count
        $volsnapExhaustCount = @($evtLogs | Where-Object {
            $_.Provider -eq 'Volsnap' -and
            $_.Message  -match 'insufficient disk space.*to grow the shadow copy storage|shadow copy storage could not grow'
        }).Count
        $vssErrCount = @($evtLogs | Where-Object {
            $_.Provider -match '^VSS$' -and $_.Level -eq 'Error'
        }).Count

        # Storage exhaustion: VolSnap confirms disk pressure caused snapshot truncation
        if ($volsnapExhaustCount -gt 0) {
            $severity = "RED"
            $issues.Add("Shadow Storage Exhaustion: VolSnap logged $volsnapExhaustCount insufficient-disk-space event(s) - storage could not grow for new shadow copies") | Out-Null
            $remediations.Add(@{
                sev  = "CRITICAL"
                title = "Shadow storage exhaustion (VolSnap Event 14)"
                cmds = @(
                    "# VolSnap could not grow shadow storage - caused backup failures / stale snapshots",
                    "# Check current allocation:",
                    "vssadmin list shadowstorage",
                    "# Increase max storage (example: 20% of volume):",
                    "vssadmin resize shadowstorage /for=C: /on=C: /maxsize=20%",
                    "# Or free disk space on the volume hosting shadow storage"
                )
            }) | Out-Null
        }

        # Manual VSS services + no active backup = hung snapshot set
        $manualVssActionable = $false
        if ($manualRunning.Count -gt 0 -and -not $backupInProgress) {
            $svcStr     = $manualRunning -join " + "
            $hasVss     = $manualRunning -contains 'VSS'
            $staleVols  = @($volumes | Where-Object { $_.behind_schedule -and $_.minutes_since_newest -gt 240 })
            $staleHours = if ($staleVols.Count -gt 0) {
                [math]::Round(($staleVols | Measure-Object minutes_since_newest -Maximum).Maximum / 60, 1)
            } else { 0 }
            # High-confidence requires VSS service itself stuck (not just SWPRV, which can linger briefly after S1 cycle)
            $highConfidence = $hasVss -and ($staleVols.Count -gt 0)
            if ($highConfidence) {
                $severity = "RED"

                # Adjust confidence label based on VolSnap event evidence
                # Storage exhaustion events reduce confidence that services are "hung" - VSS may have
                # stopped because it could not grow storage, not because it froze mid-snapshot.
                $confidenceLabel = if ($volsnapExhaustCount -gt 0) {
                    "moderate-confidence unreleased snapshot state (storage exhaustion also active - may be primary cause; fix storage first)"
                } elseif ($volsnapDeleteCount -gt 5) {
                    "moderate-high-confidence unreleased snapshot state (aggressive VolSnap deletion active - storage pressure present)"
                } else {
                    "high-confidence unreleased snapshot state - persistent VSS/SWPRV following last successful snapshot strongly correlated with all subsequent snapshot creation failures"
                }
                $evtNote = if ($volsnapExhaustCount -gt 0) {
                    " | VolSnap storage-exhaustion events: $volsnapExhaustCount"
                } elseif ($volsnapDeleteCount -gt 0) {
                    " | VolSnap deletion events: $volsnapDeleteCount (normal cleanup)"
                } else { '' }

                # Estimate "hung since" from oldest orphan snapshot - that's the last Cove operation
                # that created a snapshot and never released it; all subsequent S1/Cove attempts fail
                # because SWPRV is still holding that set's COM objects.
                $orphanTimes = @($staleVols | Where-Object { $_.orphan_oldest_time } |
                    ForEach-Object { $_.orphan_oldest_time })
                $hungSince = if ($orphanTimes.Count -gt 0) {
                    $oldest = ($orphanTimes | Measure-Object -Minimum).Minimum
                    $hungH  = [math]::Round(($deviceNow - $oldest).TotalHours, 1)
                    " | hung since $($oldest.ToString('yyyy-MM-dd HH:mm')) $deviceTzLabel (~$(Format-HM $hungH) ago)"
                } else { '' }

                $issues.Add("$svcStr (Manual-start) running with no active backup + S1 stale ${staleHours}h - $confidenceLabel${hungSince}${evtNote}") | Out-Null

                # Build stop/start sequence from services actually detected running
                # Detection criteria: (1) service is Manual-start and currently Running,
                # (2) no backup in progress (T0 not 1/9/12), (3) newest S1 snap >4h stale
                $hasSqlwriter = $manualRunning -contains 'SQLWRITER'
                $hasSwprv     = $manualRunning -contains 'SWPRV'

                # Collect specific orphan shadow GUIDs parsed from report (for targeted deletion)
                $allOrphanIds = @($staleVols | ForEach-Object {
                    if ($_.orphan_ids) { $_.orphan_ids }
                } | Where-Object { $_ })

                $cmds = @(
                    "# Pattern: last S1 snapshot ~= last successful backup -> $svcStr stayed running -> no new snapshots created",
                    "# Writers are Stable - the problem is NOT a bad writer. It is the snapshot set itself,",
                    "# which was never released. Subsequent snapshot creation attempts by both S1 and Cove fail",
                    "# because the provider is still holding the previous set's COM objects open.",
                    "#",
                    "# WHICH SNAPSHOT? The orphaned set(s) are classified as type 'other' (no auto-release,",
                    "# not client-accessible) - creation time should match the last successful backup.",
                    "# Writers are Stable - no writer restart needed. Only the snapshot set requires cleanup.",
                    "#",
                    "# HOW MANY CYCLES? ONE stop/start cycle is sufficient - but ONLY if you delete",
                    "# shadows FIRST while VSS is still running (Step 2). Stopping without pre-deleting",
                    "# re-registers the orphaned state on restart. Do not attempt multiple cycles."
                )

                if ($s1WriterPresent) {
                    $cmds += @(
                        "",
                        "# ?????????????????????????????????????????????????????????????????",
                        "# SENTINELONE DETECTED: S1 shadow copy protection will block",
                        "# 'vssadmin delete shadows'. You must unprotect BEFORE Step 2.",
                        "# Get passphrase: S1 Console -> Sentinels -> [device] -> Actions -> Show Passphrase",
                        "# ?????????????????????????????????????????????????????????????????",
                        "",
                        "# Step 0: Disable S1 tamper protection (passphrase from S1 console required)",
                        '& "C:\Program Files\SentinelOne\Sentinel Agent\sentinelctl.exe" unprotect -k "PASTE-S1-PASSPHRASE-HERE"'
                    )
                }

                $cmds += @(
                    "",
                    "# Step 1: Audit current shadow and writer state",
                    "vssadmin list shadows",
                    "vssadmin list writers",
                    ""
                )

                if ($allOrphanIds.Count -gt 0) {
                    $cmds += "# Step 2: Delete specific orphaned shadow(s) by ID while VSS is STILL RUNNING"
                    $cmds += "#         Targeted deletion - preserves any non-orphaned shadows from other tools."
                    foreach ($id in $allOrphanIds) {
                        $cmds += "vssadmin delete shadows /shadow=$id /quiet"
                    }
                } else {
                    $cmds += "# Step 2: Delete ALL shadows while VSS is STILL RUNNING (required before stopping)"
                    $cmds += "#         Shadow IDs could not be parsed from this report - deletes all copies."
                    $cmds += "#         To target specific shadows: vssadmin list shadows, then:"
                    $cmds += "#           vssadmin delete shadows /shadow={GUID} /quiet"
                    $cmds += "vssadmin delete shadows /all /quiet"
                }

                $cmds += @(
                    "",
                    "# Step 3: Stop services - SQL writer first (hard dep on SWPRV/VSS), then SWPRV, then VSS"
                )
                if ($hasSqlwriter) { $cmds += 'net stop sqlwriter' }
                if ($hasSwprv)     { $cmds += 'net stop swprv' }
                $cmds += 'net stop vss'
                $cmds += @(
                    "",
                    "# Step 4: Start VSS clean - one cycle is sufficient after shadows were pre-deleted",
                    "net start vss"
                )
                if ($hasSwprv)     { $cmds += 'net start swprv' }
                if ($hasSqlwriter) { $cmds += 'net start sqlwriter' }
                $cmds += @(
                    "",
                    "# Step 5: Verify writers recovered and no orphaned shadows remain",
                    "vssadmin list writers",
                    "vssadmin list shadows"
                )

                if ($s1WriterPresent) {
                    $cmds += @(
                        "",
                        "# Step 6: Re-enable S1 tamper protection",
                        '& "C:\Program Files\SentinelOne\Sentinel Agent\sentinelctl.exe" protect'
                    )
                }

                $hangTitle = if ($volsnapExhaustCount -gt 0) {
                    "Unreleased VSS snapshot state (storage exhaustion also active - fix storage first)"
                } else {
                    "Unreleased VSS snapshot state (snapshot lifecycle not completed - writers stable, services persistent, no new snapshots)"
                }
                $remediations.Add(@{
                    sev  = "CRITICAL"; title = $hangTitle
                    cmds = $cmds
                }) | Out-Null
                $manualVssActionable = $true
            } else {
                # Only flag if S1 is also stale - if snapshots are current, services are just finishing normally
                if ($staleVols.Count -gt 0) {
                    if ($severity -eq "GREEN") { $severity = "YELLOW" }
                    $manualVssActionable = $true
                    $issues.Add("$svcStr (Manual-start) running with no active backup + S1 stale ${staleHours}h - moderate-confidence unreleased snapshot state (VSS service not among running; verify with vssadmin list shadows)") | Out-Null
                    $remediations.Add(@{
                        sev  = "WARN"; title = "Manual VSS services running - verify snapshot lifecycle completed (no unreleased state)"
                        cmds = @(
                            "# VSS/SWPRV are Manual-start; they stop automatically when snapshots complete.",
                            "# Both running with no active backup suggests a snapshot set was never released.",
                            "vssadmin list shadows",
                            "# Delete orphaned shadows while VSS is still running (required):",
                            "vssadmin delete shadows /all /quiet",
                            "net stop swprv",
                            "net stop vss"
                        )
                    }) | Out-Null
                }
                # If staleVols = 0: S1 snapshots are current - services are finishing normally, no action needed
            }
        }

        # Writers all gone (vsswriters section exists but no individual vsswriter_* sections)
        if ($writers.Count -eq 0 -and $Sections.ContainsKey('vsswriters')) {
            $severity = "RED"
            $issues.Add("VSS CRITICAL: No writers registered - VSS service may have crashed") | Out-Null
            $remediations.Add(@{
                sev  = "CRITICAL"; title = "All VSS writers gone"
                cmds = @("net stop vss","net start vss","# If writers don't re-register, reboot the server","sc query vss")
            }) | Out-Null
        }

        # Build snaps_by_drive for timeline rendering - reuse already-classified types from $snapsByDrive
        # to guarantee timeline and volume counts are always in sync.
        if ($timeKey) {
            foreach ($drv in $snapsByDrive.Keys) {
                $drvList = [System.Collections.Generic.List[hashtable]]::new()
                foreach ($snapType in @('s1','writer','cove','native','other','unknown')) {
                    foreach ($sn in $snapsByDrive[$drv][$snapType]) {
                        $dt = Parse-UtcDateTime ($sn.$timeKey ?? '')
                        if ($dt) { $drvList.Add(@{ ts = $dt; type = $snapType }) | Out-Null }
                    }
                }
                $snapsTimeline[$drv] = @($drvList | Sort-Object { $_.ts })
            }
        }
    }

        # Parse ESENT shadow copy creation events from event logs for timeline
        # Groups by instance number -> {start, end, status} pairs with device-local timestamps
        $vssEvents = [System.Collections.Generic.List[hashtable]]::new()
        if ($evtLogs.Count -gt 0) {
            $startMap = @{}
            foreach ($ev in ($evtLogs | Sort-Object { $_.Time })) {
                if ($ev.Provider -ne 'ESENT') { continue }
                $m = [regex]::Match($ev.Message, 'Shadow copy instance (\d+) (starting|completed successfully|freeze started|freeze ended)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if (-not $m.Success) { continue }
                $inst = $m.Groups[1].Value
                $verb = $m.Groups[2].Value.ToLower()
                $dt   = Parse-UtcDateTime $ev.Time
                if (-not $dt) { continue }
                if ($verb -eq 'starting') {
                    $startMap[$inst] = $dt
                } elseif ($verb -eq 'completed successfully' -and $startMap.ContainsKey($inst)) {
                    $vssEvents.Add(@{ start = $startMap[$inst]; end = $dt; label = "VSS instance $inst" })
                    $startMap.Remove($inst)
                }
            }
            # Add any unpaired starts as point events
            foreach ($inst in $startMap.Keys) {
                $vssEvents.Add(@{ start = $startMap[$inst]; end = $null; label = "VSS instance $inst (no end)" })
            }
        }

    $t0Label = switch ($Device.T0) {
        0  { "OK" }; 1 { "InProcess" }; 2 { "Failed" }; 3 { "CompletedWithErrors" }
        5  { "Completed" }; 6 { "Skipped" }; 8 { "CompletedWithErrors" }
        9  { "InProgressWithFaults" }; 12 { "Restarted" }; default { "Status($($Device.T0))" }
    }

    $hwMfr  = if ($hwManufacturer) { $hwManufacturer } else { $Device.MF ?? '' }
    $hwMod  = if ($hwModel)        { $hwModel }        else { $Device.MO ?? '' }
    $hwCpuF = if ($hwCpu)          { $hwCpu }          else { if ($Device.I84) { "$($Device.I84) cores" } else { '' } }
    $hwRb   = [long]($Device.I85 ?? 0)
    $hwRamF = if ($hwRam) { $hwRam } elseif ($hwRb -gt 0) {
        if ($hwRb -ge 1073741824) { "$([math]::Round($hwRb / 1073741824)) GB" } else { "$([math]::Round($hwRb / 1048576)) MB" }
    } else { '' }

    # Per-DS overdue check: flag severity if any data source hasn't completed successfully
    # within 2x its cadence (YELLOW) or 4x its cadence / 24h floor (RED).
    # Skipped when a backup is in progress (first completed session may not be posted yet).
    if (-not $backupInProgress -and $LastSessions -and $LastSessions.Count -gt 0) {
        $dsNames = @{ 1='FileSystem'; 4='Exchange'; 6='Network Shares'; 7='System State';
                      8='VMware'; 10='SQL Server'; 14='Hyper-V'; 15='MySQL' }
        $nowU = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        foreach ($plugId in $LastSessions.Keys) {
            $comp = @($LastSessions[$plugId].completed)   # sorted descending unix timestamps
            if ($comp.Count -eq 0) { continue }           # never succeeded in query window; skip (separate issue)
            $lastSuccH = ($nowU - $comp[0]) / 3600.0
            # Cadence: median inter-session interval from completed timestamps
            $dsName = $dsNames[$plugId] ?? "Plugin $plugId"
            $dsCadH = 0.0
            if ($comp.Count -ge 2) {
                $intervals = for ($ci = 0; $ci -lt $comp.Count - 1; $ci++) {
                    ($comp[$ci] - $comp[$ci+1]) / 3600.0
                }
                $sortedI = @($intervals | Where-Object { $_ -gt 0 } | Sort-Object)
                if ($sortedI.Count -gt 0) {
                    $mid = [int][math]::Floor($sortedI.Count / 2)
                    $dsCadH = if ($sortedI.Count % 2 -eq 0) { ($sortedI[$mid-1] + $sortedI[$mid]) / 2 } else { $sortedI[$mid] }
                }
            }
            $yellowThreshH = if ($dsCadH -gt 0) { [math]::Max(4.0, $dsCadH * 2.0) } else { 8.0 }
            $redThreshH    = if ($dsCadH -gt 0) { [math]::Max(24.0, $dsCadH * 4.0) } else { 24.0 }
            $yellowThreshH = [math]::Round($yellowThreshH, 2)
            $redThreshH    = [math]::Round($redThreshH, 2)
            $ageRnd = [math]::Round($lastSuccH, 1)
            if ($lastSuccH -gt $redThreshH) {
                $severity = 'RED'
                $issues.Add("[ds-overdue]$dsName last successful backup: ${ageRnd}h ago (threshold: ${redThreshH}h)") | Out-Null
            } elseif ($lastSuccH -gt $yellowThreshH -and $severity -eq 'GREEN') {
                $severity = 'YELLOW'
                $issues.Add("[ds-overdue]$dsName last successful backup: ${ageRnd}h ago (threshold: ${yellowThreshH}h)") | Out-Null
            }
        }
    }

    # Hardware threshold warnings
    $cpuCores = if ($hwCpuF -match '(\d+)\s*cores?') { [int]$Matches[1] } elseif ($Device.I84) { [int]$Device.I84 } else { 0 }
    $ramGb    = if ($hwRamF -match '([\d\.]+)\s*GB') { [double]$Matches[1] } elseif ($hwRb -gt 0) { [math]::Round($hwRb / 1073741824, 1) } else { 0 }
    if ($cpuCores -gt 0 -and $cpuCores -le 2) {
        $issues.Add("[i]Low CPU: $cpuCores core$(if($cpuCores -eq 1){''}else{'s'}) - backup performance may be limited") | Out-Null
    }
    if ($ramGb -gt 0 -and $ramGb -lt 8) {
        $issues.Add("[i]Low RAM: $($hwRamF) - minimum recommended is 8 GB for reliable VSS operations") | Out-Null
    }

    # Aggregate backup status: worst-case across most recent session from each datasource
    # Hierarchy: Failed > CompletedWithErrors > Completed > Skipped > Unknown
    $backupStatus = 'Unknown'
    if ($LastSessions -and $LastSessions.Count -gt 0) {
        $hierarchy = @('Failed', 'CompletedWithErrors', 'Completed', 'Skipped', 'Unknown')
        $worst = 'Unknown'
        foreach ($pluginId in $LastSessions.Keys) {
            $plugData = $LastSessions[$pluginId]
            # Get most recent session for this datasource
            if ($plugData.sessions_full -and $plugData.sessions_full.Count -gt 0) {
                $lastSess = $plugData.sessions_full[0]  # Most recent (sorted descending by end time)
                if ($lastSess -and $lastSess.status) {
                    $sessStatus = $lastSess.status
                    # Map session status to filter categories
                    if ($sessStatus -eq 'Failed' -or $sessStatus -eq 'Aborted') {
                        $sessStatus = 'Failed'
                    } elseif ($sessStatus -eq 'NotStarted') {
                        $sessStatus = 'Skipped'
                    }
                    # Compare against hierarchy: lower index = worse status
                    if ($hierarchy.Contains($sessStatus)) {
                        $sessIdx = $hierarchy.IndexOf($sessStatus)
                        $worstIdx = $hierarchy.IndexOf($worst)
                        if ($sessIdx -lt $worstIdx) {
                            $worst = $sessStatus
                        }
                    }
                }
            }
        }
        $backupStatus = $worst
    }

    return @{
        account_id         = $aid
        device_name        = $Device.DeviceName
        machine_name       = if ($Device.MN) { $Device.MN } else { $Device.DeviceName }
        customer_name      = $Device.AR
        os                 = $Device.OS
        ip_address         = $Device.IP ?? ''
        last_timestamp     = $Device.TS  # Unix timestamp of last device contact
        agent_version      = $Device.VN ?? ''
        hw_manufacturer    = $hwMfr
        hw_model           = $hwMod
        hw_cpu             = $hwCpuF
        hw_ram             = $hwRamF
        product_id         = $Device.PD ?? ''
        product_name       = $Device.PN ?? ''
        profile_id         = $Device.OI ?? ''
        profile_name       = $Device.OP ?? ''
        retention_units    = $Device.RU ?? ''
        t0                 = $Device.T0
        t0_label           = $t0Label
        severity           = $severity
        issues             = $issues
        remediations       = $remediations
        writers            = $writers
        unhealthy_writers  = $unhealthy
        providers          = $providers
        third_party_providers = $thirdParty
        volumes            = $volumes
        vss_service        = $vssService
        has_data           = $hasData
        backup_in_progress = $backupInProgress
        backup_status      = $backupStatus
        manual_vss_running = $manualRunning
        manual_vss_actionable = $manualVssActionable
        last_sessions      = $LastSessions
        snaps_by_drive     = $snapsTimeline
        vss_events         = $vssEvents
        system_events      = & {
            # Collect Error + Warning events, deduplicate at ~90% similarity, keep 30 most recent
            $knownProc = @{
                'backup service controller' = 'N-able Cove Backup Service'
                'backupfp'                  = 'N-able Cove Backup Manager'
                'basupservice'              = 'N-able TakeControl (BeAnywhere)'
                'basupstandaloneservice'    = 'N-able TakeControl (BeAnywhere)'
                'basupexpressservice'       = 'N-able TakeControl (BeAnywhere)'
                'beanywhere'                = 'N-able TakeControl (BeAnywhere)'
                'getsupportservice'         = 'N-able TakeControl (BeAnywhere)'
            }
            function Normalize-EvtMsg([string]$msg) {
                # Strip digits, GUIDs, paths, IPs, excess whitespace to get message skeleton
                $n = $msg.ToLower()
                $n = [regex]::Replace($n, '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', 'GUID')
                $n = [regex]::Replace($n, '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', 'IP')
                $n = [regex]::Replace($n, '[a-z]:\\[^\s,\.]+', 'PATH')
                $n = [regex]::Replace($n, '\b\d+\b', '#')
                $n = [regex]::Replace($n, '\s+', ' ').Trim()
                $n
            }
            $errWarn = @($evtLogs | Where-Object { $_.Level -in @('Error','Warning') } |
                         Sort-Object { $_.Time } -Descending | Select-Object -First 200)
            $seen = @{}; $deduped = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($ev in $errWarn) {
                $key = (Normalize-EvtMsg $ev.Message).Substring(0, [math]::Min(120, (Normalize-EvtMsg $ev.Message).Length))
                if ($seen.ContainsKey($key)) { $seen[$key]++ } else {
                    $seen[$key] = 1
                    # Friendly provider/process name
                    $prov = $ev.Provider
                    foreach ($k in $knownProc.Keys) {
                        if ($ev.Message -match [regex]::Escape($k) -or $ev.Provider -match [regex]::Escape($k)) {
                            $prov = "$prov ($($knownProc[$k]))"
                            break
                        }
                    }
                    $deduped.Add([PSCustomObject]@{
                        Time    = $ev.Time
                        Message = $ev.Message
                        Level   = $ev.Level
                        Provider= $prov
                        _key    = $key
                    }) | Out-Null
                }
            }
            # Set count on each entry
            foreach ($ev in $deduped) { $ev | Add-Member -NotePropertyName Count -NotePropertyValue $seen[$ev._key] -Force }
            @($deduped | Select-Object -First 30)
        }
        device_now         = $deviceNow
        device_tz          = $deviceTzLabel
        device_tz_full     = if ($null -ne $deviceTzFull) { $deviceTzFull } else { '' }
        backup_manager_version = $backupManagerVersion
        device_type            = [int]($Device.OT ?? 0)
        s1_installed           = $s1Installed
        s1_inconsistency       = $false
        is_posix               = ($Device.OS -match 'Mac|macOS|Linux')
        detected_software      = $detectedSoftware
        volume_guid_map        = $volumeGuidToDrive
        services_map           = $svcMap
        created_at             = & {
            $cdRaw = ($Device.CD ?? '').Trim()
            if ($cdRaw -match '^\d{10,13}$') {
                # Unix timestamp (seconds or ms)
                [long]$(if ($cdRaw.Length -gt 10) { [long]$cdRaw / 1000 } else { [long]$cdRaw })
            } elseif ($cdRaw) {
                try { [DateTimeOffset]::Parse($cdRaw).ToUnixTimeSeconds() } catch { 0L }
            } else { 0L }
        }
    }
}

# ============================================================================
# Session Age Helpers
# ============================================================================

$PLUGIN_NAMES = @{
    1 = 'Files & Folders';  4 = 'Exchange';       6 = 'Network Shares'
    7 = 'System State';     8 = 'VMware';          10 = 'MS SQL'
    14 = 'Hyper-V';         15 = 'MySQL';          18 = 'System State'
}

# ============================================================================
# VSS + Backup Timeline
# ============================================================================

function TlPct([datetime]$dt, [datetime]$winStart, [double]$winSec) {
    [math]::Round([math]::Max(0,[math]::Min(100,($dt - $winStart).TotalSeconds / $winSec * 100)), 3)
}

function Get-TlAxisHtml([datetime]$winStart, [datetime]$winEnd, [int]$days, [double]$axisShiftH = 0.0, [int]$bwPx = 8400, [int]$gridEveryH = 1) {
    $winSec = ($winEnd - $winStart).TotalSeconds
    $out = ''
    # Left-edge label (device local date)
    $devWinStart = $winStart.AddHours(-$axisShiftH)   # winStart in device local time
    $out += "<div class='tl-ax-e1'></div>"
    $out += "<div class='tl-ax-ds'>$($devWinStart.ToString('ddd'))<br>$($devWinStart.ToString('MM/dd'))</div>"

    # Hour labels centered in each gridEveryH-hour cell
    $hStart = $winStart.AddHours($gridEveryH - ($winStart.Hour % $gridEveryH)).AddMinutes(-$winStart.Minute).AddSeconds(-$winStart.Second)
    $pxPerCell = [math]::Round($gridEveryH * 3600.0 / $winSec * $bwPx, 1)
    $hTime = $hStart
    while ($hTime -lt $winEnd) {
        $cellCenterPx = [math]::Round(($hTime - $winStart).TotalSeconds / $winSec * $bwPx + $pxPerCell / 2, 1)
        if ($cellCenterPx -gt 0 -and $cellCenterPx -lt $bwPx) {
            $hrLabel = $hTime.AddHours(-$axisShiftH).ToString('HH')   # device local hour
            $out += "<div class='tl-ax-hl' style='left:${cellCenterPx}px'>$hrLabel</div>"
        }
        $hTime = $hTime.AddHours($gridEveryH)
    }

    # Midnight ticks in device local time - pixel coords to match gridline divs exactly
    $devMidnight = $devWinStart.Date.AddDays(1)
    while ($devMidnight.AddHours($axisShiftH) -lt $winEnd) {
        $tickInMach = $devMidnight.AddHours($axisShiftH)
        $px = [math]::Round(($tickInMach - $winStart).TotalSeconds / $winSec * $bwPx, 1)
        if ($px -gt 20 -and $px -lt ($bwPx - 20)) {
            $out += "<div class='tl-ax-t' style='left:${px}px'></div>"
            $out += "<div class='tl-ax-dl' style='left:${px}px'><div>$($devMidnight.ToString('ddd'))</div><div>$($devMidnight.ToString('MM/dd'))</div></div>"
        }
        $devMidnight = $devMidnight.AddDays(1)
    }
    # "now" marker
    $out += "<div class='tl-ax-er'></div>"
    $out += "<div class='tl-ax-nw'>now</div>"
    $out
}

function Get-TlSnapRowHtml([array]$snaps, [datetime]$winStart, [double]$winSec, [int]$bwPx = 8400, [double]$axisShiftH = 0.0) {
    # Snap timestamps are device-local (Kind=Unspecified). winStart is machine-local (Kind=Local).
    # Shift snaps by axisShiftH to convert device-local -> machine-local for position math.
    $winEnd = $winStart.AddSeconds($winSec)
    $inWin  = @($snaps | Where-Object { $_.ts.AddHours($axisShiftH) -ge $winStart -and $_.ts.AddHours($axisShiftH) -le $winEnd })
    if ($inWin.Count -eq 0) { return '' }
    $oldest   = ($inWin | Sort-Object { $_.ts } | Select-Object -First 1).ts
    $coverPx  = [math]::Round(($oldest.AddHours($axisShiftH) - $winStart).TotalSeconds / $winSec * $bwPx, 1)
    $purgedBg  = "<div class='tl-pbg' style='width:${coverPx}px'></div>"
    $liveBg    = "<div class='tl-lbg' style='left:${coverPx}px'></div>"
    $coverMark = "<div title='Oldest surviving snapshot' class='tl-cm' style='left:${coverPx}px'></div>"
    $snapLabels = @{ 's1'='S1 (SentinelOne)'; 'cove'='Cove Backup'; 'native'='Native VSS'; 'writer'='VSS Writer'; 'vss_client'='VSS Client'; 'other'='Other VSS'; 'unknown'='Unknown' }
    $snapClass  = @{ 's1'='tl-s-s1'; 'native'='tl-s-nat'; 'writer'='tl-s-wri'; 'vss_client'='tl-s-vsc'; 'other'='tl-s-oth'; 'cove'='tl-s-cov'; 'unknown'='tl-s-unk' }
    # Vertical lanes: top=S1, middle=native/writer/vss_client/unknown, bottom=cove/other
    $snapColors = @{ 's1'='#4ecdc4'; 'cove'='#52be80'; 'native'='#fd79a8'; 'writer'='#a29bfe'; 'vss_client'='#7fb3d3'; 'other'='#787878'; 'unknown'='#fdcb6e' }
    $snapTopPx  = @{ 's1'=0; 'writer'=10; 'native'=10; 'vss_client'=10; 'unknown'=10; 'cove'=21; 'other'=21 }
    $snapHtPx   = @{ 's1'=13; 'writer'=14; 'native'=14; 'vss_client'=14; 'unknown'=14; 'cove'=13; 'other'=13 }
    $ticks = ($inWin | ForEach-Object {
        $px  = [math]::Round(($_.ts.AddHours($axisShiftH) - $winStart).TotalSeconds / $winSec * $bwPx, 1)
        $lbl = $snapLabels[$_.type] ?? $_.type
        $tip = "$lbl $($_.ts.ToString('M/d HH:mm'))"
        $cls = $snapClass[$_.type]
        if ($_.type -eq 'cove' -or $_.type -eq 'unknown') {
            if ($cls) {
                "<div title='&#9888; ACTIVE JOB: $tip' class='$cls' style='left:${px}px'></div>"
            } else {
                $col = $snapColors[$_.type] ?? '#787878'
                "<div title='&#9888; ACTIVE JOB: $tip' style='position:absolute;left:${px}px;top:0;width:4px;height:100%;background:$col;opacity:0.9;transform:translateX(-2px);cursor:default;box-shadow:0 0 6px 1px ${col}88;border-radius:1px'></div>"
            }
        } else {
            if ($cls) {
                "<div title='$tip' class='$cls' style='left:${px}px'></div>"
            } else {
                $col = $snapColors[$_.type] ?? '#787878'
                $top = $snapTopPx[$_.type] ?? 10
                $ht  = $snapHtPx[$_.type]  ?? 14
                "<div title='$tip' style='position:absolute;left:${px}px;top:${top}px;width:12px;height:${ht}px;background:linear-gradient(to right,transparent 5px,$col 5px,$col 7px,transparent 7px);transform:translateX(-6px);cursor:default'></div>"
            }
        }
    }) -join ''
    $purgedBg + $liveBg + $coverMark + $ticks
}

function Get-TlBackupRowHtml([object[]]$sessFull, [string]$label, [datetime]$winStart, [double]$winSec, [int]$bwPx = 8400) {
    $winEnd  = $winStart.AddSeconds($winSec)
    $sorted  = @($sessFull | Sort-Object { $_.start })
    
    # Per-session rendering with vertical offset for skipped/failed jobs
    # Row is 20px tall: top 10px for running/completed, bottom 10px for skipped/failed
    $results = for ($i = 0; $i -lt $sorted.Count; $i++) {
        $origS = [DateTimeOffset]::FromUnixTimeSeconds($sorted[$i].start).LocalDateTime
        $origE = [DateTimeOffset]::FromUnixTimeSeconds($sorted[$i].end).LocalDateTime
        if ($origE -lt $winStart -or $origS -gt $winEnd) { continue }
        # Snap start DOWN to nearest 5-min boundary (component arithmetic, no epoch)
        $tSsnap = $origS.AddSeconds(-$origS.Second).AddMinutes(-($origS.Minute % 5))
        # Snap end UP to nearest 5-min boundary; minimum = snapped start + 5 min
        $eRem   = $origE.Minute % 5 * 60 + $origE.Second
        $tEsnap = if ($eRem -eq 0) { $origE } else { $origE.AddSeconds(300 - $eRem) }
        if ($tEsnap -lt $tSsnap.AddMinutes(5)) { $tEsnap = $tSsnap.AddMinutes(5) }
        if ($tSsnap -lt $winStart) { $tSsnap = $winStart }
        if ($tEsnap -gt $winEnd)   { $tEsnap = $winEnd   }
        # Pixel positioning - bypasses % rounding aliasing entirely
        $lPx    = [math]::Max(0, [math]::Round(($tSsnap - $winStart).TotalSeconds / $winSec * $bwPx, 1))
        $wPx    = [math]::Max(3, [math]::Round(($tEsnap - $tSsnap).TotalSeconds  / $winSec * $bwPx, 1) - 1)
        $cls    = switch ($sorted[$i].status) { 'Completed' { 'tl-bk-ok' } 'CompletedWithErrors' { 'tl-bk-err' } 'Skipped' { 'tl-bk-skp' } default { 'tl-bk-fai' } }
        $cls   += $(if($sorted[$i].cleaned){' tl-bk-cl'}) + $(if($sorted[$i].accelerated){' tl-bk-ac'})
        # Offset skipped/failed jobs to bottom half (top: 10px), others to top half (top: 0)
        $topPx  = if ($sorted[$i].status -in @('Skipped', 'Failed', 'Interrupted')) { '10px' } else { '0px' }
        $dur         = [int](($origE - $origS).TotalSeconds / 60)
        $bv          = $sorted[$i].buildVersion
        $flags       = @($(if($sorted[$i].cleaned){'Cleaned'}), $(if($sorted[$i].accelerated){'Accelerated'})) | Where-Object { $_ }
        $verChanged  = ($i -gt 0 -and $bv -and $sorted[$i-1].buildVersion -and $bv -ne $sorted[$i-1].buildVersion)
        $sc          = switch ($sorted[$i].status) { 'Completed' { 'jt-ok' } 'CompletedWithErrors' { 'jt-err' } 'Skipped' { 'jt-skp' } default { 'jt-fai' } }
        $flagsJson   = '[' + (($flags | ForEach-Object { '"' + $_ + '"' }) -join ',') + ']'
        $bvJson      = if ($bv) { '"' + $bv + '"' } else { 'null' }
        $prevBldJson = if ($verChanged) { '"' + $sorted[$i-1].buildVersion + '"' } else { 'null' }
        $safeLabel   = $label -replace '&', '&amp;'
        $dataTip     = '{"p":"' + $safeLabel + '","d":"' + $origS.ToString('M/d') + '","s":"' + $origS.ToString('HH:mm') + '","e":"' + $origE.ToString('HH:mm') + '","dur":' + $dur + ',"st":"' + $sorted[$i].status + '","sc":"' + $sc + '","fl":' + $flagsJson + ',"bv":' + $bvJson + ',"pb":' + $prevBldJson + ',"cc":' + ($sorted[$i].changedCount ?? 0) + ',"cb":' + ($sorted[$i].changedBytes ?? 0) + ',"err":' + ($sorted[$i].errorsCount ?? 0) + ',"sc2":' + ($sorted[$i].selectedCount ?? 0) + ',"sb":' + ($sorted[$i].selectedBytes ?? 0) + ',"nc":' + ($sorted[$i].newCount ?? 0) + ',"nb":' + ($sorted[$i].newBytes ?? 0) + ',"snt":' + ($sorted[$i].sentBytes ?? 0) + ',"dc":' + ($sorted[$i].deletedCount ?? 0) + '}'
        $verMark     = if ($verChanged) { "<div title='Build updated: $bv (was $($sorted[$i-1].buildVersion))' style='position:absolute;left:${lPx}px;top:0;font-size:10px;font-weight:bold;color:#e74c3c;cursor:help;z-index:2;line-height:1;pointer-events:all;transform:translateX(-100%)'># </div>" } else { '' }
        "<div data-tip='$dataTip' onmouseenter='_tlShowTip(event,this)' onmousemove='_tlMoveTip(event)' onmouseleave='_tlHideTip()' class='$cls' style='position:absolute;left:${lPx}px;width:${wPx}px;height:10px;top:${topPx}'></div>$verMark"
    }
    $results -join ''
}

function Get-TimelineHtml([hashtable]$D) {
    if (-not $D.snaps_by_drive -and (-not $D.last_sessions -or $D.last_sessions.Count -eq 0)) { return '' }

    # Helper: median cadence hours from sorted-desc unix timestamp array
    function TlCadH([long[]]$ts) {
        if ($ts.Count -lt 2) { return 24.0 }
        $gaps = @(); for ($i = 0; $i -lt [math]::Min($ts.Count-1,20); $i++) { $g = ($ts[$i]-$ts[$i+1])/3600.0; if ($g -gt 0) { $gaps += $g } }
        if ($gaps.Count -eq 0) { return 24.0 }
        (@($gaps | Sort-Object))[[int]($gaps.Count/2)]
    }
    function TlCadTxt([double]$h) {
        $m = [int][math]::Round($h*60)
        if ($h -lt 1.5) { return "~${m}m" } elseif ($h -lt 36) { return "~$([int][math]::Round($h))h" } else { return "~$([int]($h/24))d" }
    }
    function TlAgeStr([datetime]$dt, [datetime]$ref) {
        $a = ($ref - $dt).TotalHours
        if ($a -lt 1) { "$([int]($a*60))m ago" } elseif ($a -lt 48) { "$([int]$a)h ago" } else { "$([int]($a/24))d ago" }
    }

    # Determine window from fastest plugin cadence
    $medianH = 24.0
    foreach ($plugId in $D.last_sessions.Keys) {
        $comp = @($D.last_sessions[$plugId].completed)
        if ($comp.Count -ge 2) {
            $m = TlCadH $comp
            if ($m -lt $medianH) { $medianH = $m }
        }
    }
    $winDays   = if ($medianH -le 1.5) { 7 } elseif ($medianH -le 8) { 14 } else { 30 }
    $pxPerHour = if ($medianH -le 1.5) { 50 } elseif ($medianH -le 8) { 25 } else { 12 }
    $barWidthPx = [int]($winDays * 24 * $pxPerHour)   # ~8400px across all cadence tiers
    # Grid interval: every 2h for dense hourly views, 4h / 6h for sparser ones
    $gridEveryH = if ($pxPerHour -ge 40) { 1 } elseif ($pxPerHour -ge 20) { 2 } else { 4 }
    # Always use machine's current local time as winEnd - device_now from InternalInfo can be
    # stale (cached up to 5 min, or much longer on previous run) causing session clipping.
    # Sessions come from a live repserv call so anchoring to [datetime]::Now keeps them in-window.
    $winEnd   = [datetime]::Now
    $winStart = $winEnd.AddDays(-$winDays)
    $winSec   = ($winEnd - $winStart).TotalSeconds

    # Compute axis shift: device midnight may differ from machine midnight by TZ offset
    # axisShiftH = machUTCoffset - devUTCoffset  (positive = device is west/behind machine)
    $axisShiftH = 0.0
    if ($D.device_tz_full) {
        try {
            $machTzLoc = [System.TimeZoneInfo]::Local
            $devTzLoc  = [System.TimeZoneInfo]::GetSystemTimeZones() |
                         Where-Object { $_.Id -eq $D.device_tz_full -or
                                        $_.StandardName -eq $D.device_tz_full -or
                                        $_.DaylightName -eq $D.device_tz_full } |
                         Select-Object -First 1
            if ($devTzLoc) {
                $axisShiftH = $machTzLoc.GetUtcOffset($winEnd).TotalHours -
                              $devTzLoc.GetUtcOffset($winEnd).TotalHours
            }
        } catch {}
    }

    $PLUGIN_NAMES_TL = @{1='FileSystem';4='Exchange';6='Net Shares';7='Sys State';8='VMware';10='MS SQL';14='Hyper-V';15='MySQL';18='Sys State'}
    $lStyle = 'width:185px;min-width:185px;font-size:0.74em;text-align:right;padding-right:8px;flex-shrink:0'
    $rH = '34px'  # fixed row height - must fit 2-line label AND bar content

    $labelItems = ''
    $barItems   = ''
    # Snap rows per drive (snaps stored UTC -> convert to local)
    if ($D.snaps_by_drive) {
        foreach ($drv in ($D.snaps_by_drive.Keys | Sort-Object)) {
            $localSnaps = @($D.snaps_by_drive[$drv] | ForEach-Object { @{ ts = $_.ts; type = $_.type } })
            $inWin = @($localSnaps | Where-Object { $_.ts.AddHours($axisShiftH) -ge $winStart })
            if ($inWin.Count -eq 0) { continue }
            $r = Get-TlSnapRowHtml $inWin $winStart $winSec $barWidthPx $axisShiftH

            # ?? Cadence: S1 snaps calculated independently from all other types ??
            function SnapMedianH([array]$snaps) {
                $sorted = @($snaps | Sort-Object { $_.ts } -Descending)
                $gaps = @(); for ($gi = 0; $gi -lt [math]::Min($sorted.Count-1,20); $gi++) {
                    $g = ($sorted[$gi].ts - $sorted[$gi+1].ts).TotalHours; if ($g -gt 0) { $gaps += $g }
                }
                if ($gaps.Count -eq 0) { return 0.0 }
                (@($gaps | Sort-Object))[[int]($gaps.Count/2)]
            }
            # Detect cadence: interval OR time-of-day pattern
            # Returns @{ type; medH; label; dailyToD (nullable) }
            function DetectSnapCadence([array]$snaps) {
                if ($snaps.Count -lt 2) { return @{ type='unknown'; medH=0; label=''; dailyToD=$null } }
                $sorted = @($snaps | Sort-Object { $_.ts })
                $gaps = @(); for ($gi = 0; $gi -lt [math]::Min($sorted.Count-1,20); $gi++) {
                    $g = ($sorted[$gi+1].ts - $sorted[$gi].ts).TotalHours; if ($g -gt 0) { $gaps += $g }
                }
                if ($gaps.Count -eq 0) { return @{ type='unknown'; medH=0; label=''; dailyToD=$null } }
                $medH = (@($gaps | Sort-Object))[[int]($gaps.Count/2)]
                $mean = ($gaps | Measure-Object -Average).Average
                $cv   = if ($mean -gt 0) { [math]::Sqrt(($gaps | ForEach-Object {($_-$mean)*($_-$mean)} | Measure-Object -Average).Average) / $mean } else { 0 }
                # Regular interval: low coefficient of variation
                if ($cv -lt 0.35) { return @{ type='interval'; medH=$medH; label=(TlCadTxt $medH); dailyToD=$null } }
                # Irregular: check for time-of-day clustering (e.g. daily at 2 AM)
                if ($snaps.Count -ge 3 -and $medH -gt 18) {
                    $todMins = @($snaps | ForEach-Object { $_.ts.Hour * 60 + $_.ts.Minute })
                    $todMean = ($todMins | Measure-Object -Average).Average
                    $todSd   = [math]::Sqrt(($todMins | ForEach-Object {($_-$todMean)*($_-$todMean)} | Measure-Object -Average).Average)
                    if ($todSd -lt 45) {   # all snaps within ?45 min of same time of day
                        $h = [int]($todMean / 60); $m = [int]($todMean % 60)
                        $todLabel = "daily@$($h.ToString('D2')):$($m.ToString('D2'))"
                        return @{ type='daily'; medH=24.0; label=$todLabel; dailyToD=@{h=$h;m=$m;sdMin=$todSd} }
                    }
                }
                return @{ type='irregular'; medH=$medH; label=''; dailyToD=$null }
            }

            $s1Snaps       = @($localSnaps | Where-Object { $_.type -eq 's1' })
            $nonS1Snaps    = @($localSnaps | Where-Object { $_.type -ne 's1' })
            $s1MedH        = if ($s1Snaps.Count -ge 2) { SnapMedianH $s1Snaps } else { 0 }
            $nonS1Cadence  = if ($nonS1Snaps.Count -ge 2) { DetectSnapCadence $nonS1Snaps } else { @{ type='unknown'; medH=0; label=''; dailyToD=$null } }
            $nonS1MedH     = $nonS1Cadence.medH

            # For overdue/missed detection use the dominant cadence:
            $snapMedH = if ($s1MedH -gt 0) { $s1MedH } elseif ($nonS1MedH -gt 0) { $nonS1MedH } else { 0 }

            # Most recent snap (all types) for age display
            $sortedAll = @($localSnaps | Sort-Object { $_.ts } -Descending)
            $lastSnap     = $sortedAll[0].ts
            $lastSnapMach = $lastSnap.AddHours($axisShiftH)   # convert to machine-local for comparisons
            $ageStr    = TlAgeStr $lastSnapMach $winEnd
            $overdue   = $snapMedH -gt 0 -and ($winEnd - $lastSnapMach).TotalHours -gt 1.5 * $snapMedH
            $infoCol   = if ($overdue) { '#e74c3c' } else { '#555' }

            # Label: show S1 cadence only (non-S1 cadence omitted from display)
            $nonS1LabelPart = if ($nonS1Cadence.label) { " ? $($nonS1Cadence.label)" } else { '' }
            $cadTxt  = if ($s1MedH -gt 0) { "$(TlCadTxt $s1MedH) S1" } elseif ($nonS1Cadence.label) { $nonS1Cadence.label } else { '' }
            $infoTxt = if ($cadTxt) { "$cadTxt ? $ageStr" } else { $ageStr }
            # Missed snapshot markers - gap-based, cadence computed per snap type (S1 vs others)
            # Only flag missing if we have at least 3+ snaps to establish reliable cadence
            $missedSnapHtml = ''
            foreach ($snapTypeGroup in @(@{snaps=$s1Snaps;medH=$s1MedH;label='S1';cadObj=$null},@{snaps=$nonS1Snaps;medH=$nonS1MedH;label='native';cadObj=$nonS1Cadence})) {
                $tgSnaps = @($snapTypeGroup.snaps | Where-Object { $_.ts.AddHours($axisShiftH) -ge $winStart })
                $tgMedH  = $snapTypeGroup.medH
                $tgLabel = $snapTypeGroup.label
                # Skip entirely if fewer than 3 snapshots (can't establish reliable cadence from 1-2 data points)
                if ($tgSnaps.Count -lt 3) { continue }
                if ($tgMedH -le 0) { continue }
                $tgSorted = @($tgSnaps | Sort-Object { $_.ts })
                $suppressEnd = $winEnd.AddHours(-$tgMedH)   # don't flag within last cadence of now

                # Case 1: gaps between consecutive snaps within the window
                # dotTop: center the 8px dot within the lane (lane top + 1px padding)
                $dotTop = if ($tgLabel -eq 'S1') { 3 } else { 13 }
                $msCls  = if ($tgLabel -eq 'S1') { 'tl-mss1' } else { 'tl-msna' }
                for ($mi = 0; $mi -lt $tgSorted.Count - 1; $mi++) {
                    $gapH = ($tgSorted[$mi+1].ts - $tgSorted[$mi].ts).TotalHours
                    if ($gapH -gt $tgMedH * 1.5) {
                        $nMissed = [int][math]::Floor($gapH / $tgMedH) - 1
                        for ($mk = 1; $mk -le $nMissed; $mk++) {
                            $exp = $tgSorted[$mi].ts.AddHours($tgMedH * $mk)
                            $expMach = $exp.AddHours($axisShiftH)
                            if ($expMach -ge $winStart -and $expMach -le $suppressEnd) {
                                $mp  = [math]::Round(($expMach - $winStart).TotalSeconds / $winSec * $barWidthPx, 1)
                                $cadFmt = TlCadTxt $tgMedH
                                $tip = "Missing $tgLabel snapshot | Expected: $($exp.ToString('M/d HH:mm')) | Cadence: $cadFmt | Gap: $([math]::Round($gapH,1))h between $($tgSorted[$mi].ts.ToString('HH:mm')) and $($tgSorted[$mi+1].ts.ToString('HH:mm'))"
                                $missedSnapHtml += "<div title='$tip' class='$msCls' style='left:${mp}px'></div>"
                            }
                        }
                    }
                }

                # Case 2: trailing gap (and daily pattern walk)
                $allTgSorted = @($snapTypeGroup.snaps | Sort-Object { $_.ts })
                $lastTgSnap  = if ($tgSorted.Count -gt 0) { $tgSorted[-1].ts }
                               elseif ($allTgSorted.Count -gt 0) { $allTgSorted[-1].ts }
                               else { $null }
                $cadObj = $snapTypeGroup.cadObj
                $isDailyToD = $cadObj -and $cadObj.type -eq 'daily' -and $cadObj.dailyToD
                if ($lastTgSnap) {
                    if ($isDailyToD) {
                        # Time-of-day pattern: walk day by day at the detected daily time.
                        # Only flag days where this weekday has EVER had a snap - native snaps
                        # are schedule-based (not frequency-based) so we can't infer a miss
                        # on a weekday the schedule has never historically run.
                        $tod = $cadObj.dailyToD
                        $tolMin = [math]::Max(30, $tod.sdMin * 2)
                        $activeDOW = @($allTgSorted | ForEach-Object { [int]$_.ts.DayOfWeek } | Sort-Object -Unique)
                        $day = $winStart.Date
                        while ($day -le $winEnd) {
                            if ([int]$day.DayOfWeek -in $activeDOW) {
                                $exp = $day.AddHours($tod.h).AddMinutes($tod.m)
                                $expMach = $exp.AddHours($axisShiftH)
                                if ($expMach -ge $winStart -and $expMach -le $suppressEnd) {
                                    $hasSnap = [bool]($allTgSorted | Where-Object { [math]::Abs(($_.ts - $exp).TotalMinutes) -le $tolMin })
                                    if (-not $hasSnap) {
                                        $mp   = [math]::Round(($expMach - $winStart).TotalSeconds / $winSec * $barWidthPx, 1)
                                        $agoH = [math]::Round(($winEnd - $exp).TotalHours, 1)
                                        $tip  = "Missing $tgLabel snapshot | Expected: $($exp.ToString('M/d HH:mm')) ($($tod.h.ToString('D2')):$($tod.m.ToString('D2')) daily schedule, $(([System.DayOfWeek]$day.DayOfWeek).ToString().Substring(0,3))) | $(Format-HM $agoH) ago"
                                        $missedSnapHtml += "<div title='$tip' class='$msCls' style='left:${mp}px'></div>"
                                    }
                                }
                            }
                            $day = $day.AddDays(1)
                        }
                    } elseif (($winEnd - $lastTgSnap.AddHours($axisShiftH)).TotalHours -gt $tgMedH * 1.5) {
                        # Interval pattern: project forward from last snap
                        $exp = $lastTgSnap.AddHours($tgMedH)
                        while ($exp.AddHours($axisShiftH) -le $suppressEnd) {
                            $expMach = $exp.AddHours($axisShiftH)
                            if ($expMach -ge $winStart) {
                                $mp     = [math]::Round(($expMach - $winStart).TotalSeconds / $winSec * $barWidthPx, 1)
                                $cadFmt = TlCadTxt $tgMedH
                                $agoH   = [math]::Round(($winEnd - $lastTgSnap).TotalHours, 1)
                                $tip    = "Missing $tgLabel snapshot | Expected: $($exp.ToString('M/d HH:mm')) | Cadence: $cadFmt | Last seen: $($lastTgSnap.ToString('M/d HH:mm')) ($(Format-HM $agoH) ago)"
                                $missedSnapHtml += "<div title='$tip' class='$msCls' style='left:${mp}px'></div>"
                            }
                            $exp = $exp.AddHours($tgMedH)
                        }
                    }
                }
            }
            $lbl     = "snaps $($drv.TrimEnd('\')) ($($localSnaps.Count))"
            $lblHtml = "<div style='overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#888'>$lbl</div><div style='font-size:0.90em;color:$infoCol;white-space:nowrap'>$infoTxt</div>"
            $labelItems += "<div class='tl-lrow'>$lblHtml</div>"
            $barItems   += "<div class='tl-brow' style='width:${barWidthPx}px'>$r$missedSnapHtml</div>"
        }
    }
    # Backup rows per plugin
    foreach ($plugId in ($D.last_sessions.Keys | Sort-Object)) {
        $sf = @($D.last_sessions[$plugId].sessions_full)
        if ($sf.Count -eq 0) { continue }
        $name    = $PLUGIN_NAMES_TL[$plugId] ?? "Plugin $plugId"
        $r       = Get-TlBackupRowHtml $sf $name $winStart $winSec $barWidthPx
        $history = @($D.last_sessions[$plugId].completed)
        $cadH    = if ($history.Count -ge 2) { TlCadH $history } else { 0 }
        $lastTs  = if ($history.Count -gt 0) { $history[0] } else { 0 }
        $ageStr  = Format-RelativeTime $lastTs
        $cadTxt  = if ($cadH -gt 0) { TlCadTxt $cadH } else { '' }
        $infoTxt = if ($cadTxt) { "$cadTxt avg ? $ageStr" } else { $ageStr }
        $ageH    = if ($lastTs -gt 0) { ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - $lastTs) / 3600.0 } else { 999 }
        $overdue = $cadH -gt 0 -and $ageH -gt 1.5 * $cadH
        $infoCol = if ($overdue) { '#e74c3c' } elseif ($cadH -gt 0 -and $ageH -gt 1.2 * $cadH) { '#e67e22' } else { '#555' }
        # Missed backup markers - gap-based; suppress if another plugin was running during the gap
        $missedBackHtml = ''
        if ($cadH -gt 0 -and $sf.Count -ge 2) {
            # Pre-build other-plugin session ranges (LocalDateTime) for overlap check
            $otherRanges = [System.Collections.Generic.List[object]]::new()
            foreach ($opId in $D.last_sessions.Keys) {
                if ($opId -eq $plugId) { continue }
                foreach ($os in $D.last_sessions[$opId].sessions_full) {
                    $otherRanges.Add(@{
                        s = [DateTimeOffset]::FromUnixTimeSeconds($os.start).LocalDateTime
                        e = [DateTimeOffset]::FromUnixTimeSeconds($os.end).LocalDateTime
                    }) | Out-Null
                }
            }
            $sortedSf = @($sf | Sort-Object { $_.start })
            for ($mi = 0; $mi -lt $sortedSf.Count - 1; $mi++) {
                $gapH = ($sortedSf[$mi+1].start - $sortedSf[$mi].start) / 3600.0
                if ($gapH -gt $cadH * 1.5) {
                    $nMissed = [int][math]::Floor($gapH / $cadH) - 1
                    for ($mk = 1; $mk -le $nMissed; $mk++) {
                        $exp = [DateTimeOffset]::FromUnixTimeSeconds($sortedSf[$mi].start).LocalDateTime.AddHours($cadH * $mk)
                        if ($exp -lt $winStart -or $exp -gt $winEnd.AddHours(-$cadH)) { continue }
                        # Suppress if another plugin was running at $exp (plugins are sequential - the gap is just waiting)
                        $blocked = [bool]($otherRanges | Where-Object { $exp -ge $_.s -and $exp -le $_.e })
                        if (-not $blocked) {
                            $mp = [math]::Round(($exp - $winStart).TotalSeconds / $winSec * $barWidthPx, 1)
                            $missedBackHtml += "<div title='Expected $name ~$($exp.ToString('M/d HH:mm')) - no attempt recorded' class='tl-msb' style='left:${mp}px'></div>"
                        }
                    }
                }
            }
        }
        # Success % + RPO pills
        # Clamp window to device creation date so new devices aren't penalised for history they don't have
        $nowU2       = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $winStartU   = [long]([DateTimeOffset]::new($winStart).ToUnixTimeSeconds())
        $createdAt   = if ($D.created_at) { [long]$D.created_at } else { 0L }
        $effectiveStartU = if ($createdAt -gt $winStartU) { $createdAt } else { $winStartU }
        $effectiveDays   = [math]::Round(($nowU2 - $effectiveStartU) / 86400.0, 1)
        $windowClamped   = $createdAt -gt $winStartU
        $sfWin     = @($sf | Where-Object { $_.end -ge $effectiveStartU })
        $nComp     = @($sfWin | Where-Object { $_.status -eq 'Completed' }).Count
        $nClean    = @($sfWin | Where-Object { $_.cleaned -and $_.status -ne 'Completed' }).Count
        $nOk       = $nComp + $nClean
        $nWithErr  = @($sfWin | Where-Object { $_.status -eq 'CompletedWithErrors' -and -not $_.cleaned }).Count
        $nFail     = @($sfWin | Where-Object { $_.status -in @('Failed','Aborted') }).Count
        $nSkip     = @($sfWin | Where-Object { $_.status -eq 'Skipped' }).Count
        $nRan      = $nOk + $nWithErr + $nFail
        $expSlots  = if ($cadH -gt 0) { [int][math]::Floor($effectiveDays * 24 / $cadH) } else { $nRan }
        $nMissed   = [math]::Max(0, $expSlots - $nRan)
        $nExtra    = [math]::Max(0, $nRan - $expSlots)
        $denom     = $nRan + $nMissed
        $sPct      = if ($denom -gt 0) { [math]::Round($nOk / $denom * 100, 0) } else { $null }
        $durs      = @($sfWin | Where-Object { $_.status -eq 'Completed' -and $_.start -gt 0 } |
                       ForEach-Object { ($_.end - $_.start) / 3600.0 } |
                       Where-Object { $_ -gt 0 -and ($cadH -le 0 -or $_ -lt ($cadH * 3)) })
        $medDurH   = if ($durs.Count -gt 0) { @($durs | Sort-Object)[[int]($durs.Count / 2)] } else { 0.0 }
        $rpoRat    = if ($cadH -gt 0 -and $medDurH -gt 0) { $medDurH / $cadH } else { 0.0 }
        $pillHtml  = ''
        if ($null -ne $sPct) {
            $sBg = if ($sPct -ge 95) { '#1a3a1a' } elseif ($sPct -ge 80) { '#3a2e0a' } else { '#3a1010' }
            $sFg = if ($sPct -ge 95) { '#6ecf6e' } elseif ($sPct -ge 80) { '#f0a030' } else { '#e05050' }
            $sBd = if ($sPct -ge 95) { '#2a5a2a' } elseif ($sPct -ge 80) { '#5a4a10' } else { '#5a1010' }
            $clampNote = if ($windowClamped) { "&#10;&#9432; Window clamped to device age ($($effectiveDays)d) - device created $([math]::Round(($nowU2 - $createdAt)/86400.0,0))d ago" } else { '' }
            $sTip = "$name success - $($effectiveDays)d window&#10;Completed: $nOk  Failed/errors: $($nWithErr+$nFail)&#10;Missed: $nMissed  Skipped: $nSkip (excluded)&#10;Extra/manual: $nExtra  Expected: $expSlots$clampNote"
            $pillHtml += "<span title='$sTip' style='display:inline-block;padding:0 6px;border-radius:10px;font-size:0.70em;font-weight:bold;background:$sBg;color:$sFg;border:1px solid $sBd;text-align:center;cursor:help;white-space:nowrap'>${sPct}%$(if($windowClamped){"*"})</span>"
        }
        if ($rpoRat -gt 0) {
            $rBg = if ($rpoRat -lt 0.5) { '#122030' } elseif ($rpoRat -lt 0.8) { '#302818' } else { '#301010' }
            $rFg = if ($rpoRat -lt 0.5) { '#4a9eda' } elseif ($rpoRat -lt 0.8) { '#c8943a' } else { '#e04040' }
            $rBd = if ($rpoRat -lt 0.5) { '#1a3a5a' } elseif ($rpoRat -lt 0.8) { '#5a4010' } else { '#5a1010' }
            $rTxt = if ($rpoRat -lt 0.5) { 'RPO ok' } elseif ($rpoRat -lt 0.8) { 'RPO ~' } else { 'RPO !' }
            $hmHead = Format-HM ([math]::Max(0.0, $cadH - $medDurH))
            $rTip = "$name RPO - $($winDays)d window&#10;Median duration: $(Format-HM $medDurH)  Cadence: $(Format-HM $cadH)&#10;Window consumed: $([math]::Round($rpoRat*100,0))%  Headroom: $hmHead"
            $pillHtml += "<span title='$rTip' style='display:inline-block;padding:0 6px;border-radius:10px;font-size:0.68em;font-weight:bold;background:$rBg;color:$rFg;border:1px solid $rBd;text-align:center;cursor:help;white-space:nowrap'>$rTxt</span>"
        }
        if ($pillHtml) {
            # 2-line layout: name + pills on line 1 (flex row), cadence/age on line 2
            $nameLine  = "<div style='display:flex;align-items:center;justify-content:flex-end;gap:3px;min-width:0'><div style='overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#7fb3d3;flex:1;text-align:right'>$name</div>$pillHtml</div>"
            $infoLine  = "<div style='font-size:0.88em;color:$infoCol;white-space:nowrap;text-align:right;padding-right:2px'>$infoTxt</div>"
            $lblHtml   = "$nameLine$infoLine"
            $lRowStyle = "style='height:44px;flex-direction:column;justify-content:center;gap:3px;padding-right:6px'"
            $barH      = "height:44px;"
        } else {
            $lblHtml   = "<div style='overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#7fb3d3'>$name</div><div style='font-size:0.90em;color:$infoCol;white-space:nowrap'>$infoTxt</div>"
            $lRowStyle = "style='height:34px'"
            $barH      = ""
        }
        $labelItems += "<div class='tl-lrow' $lRowStyle>$lblHtml</div>"
        $barItems   += "<div class='tl-brow-b' style='${barH}width:${barWidthPx}px'>$r$missedBackHtml</div>"
    }
    # VSS snapshot creation events row (from ESENT event log, device-local timestamps)
    if ($D.vss_events -and $D.vss_events.Count -gt 0) {
        $vssRowHtml = ''
        foreach ($ev in $D.vss_events) {
            $tS = $ev.start
            if ($tS -lt $winStart -or $tS -gt $winEnd) { continue }
            $lPx = [math]::Max(0, [math]::Round(($tS - $winStart).TotalSeconds / $winSec * $barWidthPx, 1))
            if ($ev.end) {
                $tE  = if ($ev.end -gt $winEnd) { $winEnd } else { $ev.end }
                $wPx = [math]::Max(3, [math]::Round(($tE - $tS).TotalSeconds / $winSec * $barWidthPx, 1))
                $dur = [math]::Round(($ev.end - $tS).TotalSeconds, 0)
                $tip = "$($ev.label) | $($tS.ToString('M/d HH:mm:ss'))-$($ev.end.ToString('HH:mm:ss')) (${dur}s)"
                $vssRowHtml += "<div title='$tip' class='tl-vss-ev' style='left:${lPx}px;width:${wPx}px'></div>"
            } else {
                $tip = "$($ev.label) | $($tS.ToString('M/d HH:mm:ss')) - no completion recorded"
                $vssRowHtml += "<div title='$tip' class='tl-vss-pt' style='left:${lPx}px'></div>"
            }
        }
        if ($vssRowHtml) {
            $inWinCount = ($D.vss_events | Where-Object { $_.start -ge $winStart -and $_.start -le $winEnd }).Count
            $lblHtml = "<div style='overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#5dade2'>VSS ops ($inWinCount)</div><div style='font-size:0.90em;color:#555;white-space:nowrap'>from event log</div>"
            $labelItems += "<div class='tl-lrow'>$lblHtml</div>"
            $barItems   += "<div class='tl-brow-b' style='width:${barWidthPx}px'>$vssRowHtml</div>"
        }
    }

    if (-not $labelItems) { return '' }

    $axLbl  = "<div class='tl-ax-lc'>time</div>"
    $axBar  = "<div class='tl-ax-row' style='width:${barWidthPx}px'>$(Get-TlAxisHtml $winStart $winEnd $winDays $axisShiftH $barWidthPx $gridEveryH)</div>"
    $nowLn  = "<div title='Now' class='tl-now'></div>"
    $legend = "<div style='font-size:0.78em;color:#555;margin-top:6px'>" +
        "<span><span style='display:inline-block;width:10px;height:10px;background:#27ae60;vertical-align:middle;margin-right:3px'></span>Completed</span>&ensp;" +
        "<span><span style='display:inline-block;width:10px;height:10px;background:#e67e22;vertical-align:middle;margin-right:3px'></span>w/Errors</span>&ensp;" +
        "<span><span style='display:inline-block;width:10px;height:10px;background:#888;vertical-align:middle;margin-right:3px'></span>Skipped</span>&ensp;" +
        "<span><span style='display:inline-block;width:10px;height:10px;background:#e74c3c;vertical-align:middle;margin-right:3px'></span>Failed</span>&ensp;" +
        "<span style='position:relative;display:inline-flex;align-items:center;margin-right:3px'><span style='display:inline-block;width:10px;height:10px;background:#27ae60;border-radius:1px;vertical-align:middle'></span><span style='position:absolute;top:0;left:0;right:0;height:3px;background:rgba(255,255,255,0.65);border-radius:1px 1px 0 0'></span></span>Accelerated&ensp;" +
        "<span style='position:relative;display:inline-flex;align-items:center;margin-right:3px'><span style='display:inline-block;width:10px;height:10px;background:#27ae60;border-radius:1px;vertical-align:middle;background-image:repeating-linear-gradient(45deg,rgba(0,0,0,0) 0px,rgba(0,0,0,0) 3px,rgba(0,0,0,0.38) 3px,rgba(0,0,0,0.38) 5px)'></span></span>Cleaned&ensp;" +
        "<span><span style='display:inline-block;width:3px;height:10px;background:#4ecdc4;vertical-align:middle;margin-right:3px'></span>S1 snap</span>&ensp;" +
        "<span><span style='display:inline-block;width:3px;height:10px;background:#fd79a8;vertical-align:middle;margin-right:3px'></span>Native snap</span>&ensp;" +
        "<span><span style='display:inline-block;width:3px;height:10px;background:#52be80;vertical-align:middle;margin-right:3px'></span>Cove snap (active job)</span>&ensp;" +
        "<span><span style='display:inline-block;border-left:2px dashed rgba(243,156,18,0.70);height:10px;vertical-align:middle;margin-right:3px'></span>Missed slot</span>" +
        "</div>"
    $tlId     = 'tl-' + [System.Guid]::NewGuid().ToString('N').Substring(0,8)
    $arrowL   = "<button class='tl-btn tl-lbtn' onclick='tlPage(`"$tlId`",-1)' title='Earlier'>&#9664;</button>"
    $arrowR   = "<button class='tl-btn tl-rbtn' onclick='tlPage(`"$tlId`",1)' title='Later'>&#9654;</button>"
    $labelCol = "<div class='tl-lc'>$axLbl$labelItems</div>"
    # Debug: capture timeline calc rows before returning
    if ($Script:TlDebug) {
        foreach ($plugId in ($D.last_sessions.Keys | Sort-Object)) {
            $sfD = @($D.last_sessions[$plugId].sessions_full)
            $nameD = $PLUGIN_NAMES_TL[$plugId] ?? "Plugin $plugId"
            foreach ($sess in ($sfD | Sort-Object { $_.start })) {
                $tSo = [DateTimeOffset]::FromUnixTimeSeconds($sess.start).LocalDateTime
                $tEo = [DateTimeOffset]::FromUnixTimeSeconds($sess.end).LocalDateTime
                $tSc = $tSo; $tEc = $tEo; $clipS = $false; $clipE = $false
                if ($tSc -lt $winStart) { $tSc = $winStart; $clipS = $true }
                if ($tEc -gt $winEnd)   { $tEc = $winEnd;   $clipE = $true }
                $lp = TlPct $tSc $winStart $winSec
                $rp = TlPct $tEc $winStart $winSec
                $dp = [math]::Round($rp - $lp, 4)
                $Script:TlDebugRows.Add([PSCustomObject]@{
                    Device       = $D.machine_name
                    AccountId    = $D.account_id
                    Plugin       = $nameD
                    StartUnix    = $sess.start
                    EndUnix      = $sess.end
                    StartLocal   = $tSo.ToString('yyyy-MM-dd HH:mm:ss')
                    EndLocal     = $tEo.ToString('yyyy-MM-dd HH:mm:ss')
                    DurActualMin = [int](($tEo - $tSo).TotalSeconds / 60)
                    WinStart     = $winStart.ToString('yyyy-MM-dd HH:mm:ss')
                    WinEnd       = $winEnd.ToString('yyyy-MM-dd HH:mm:ss')
                    WinSec       = [int]$winSec
                    BarWidthPx   = $barWidthPx
                    PxPerHour    = $pxPerHour
                    LPct         = $lp
                    RPct         = $rp
                    DurPct       = $dp
                    VisWidthPx   = [math]::Round($dp / 100.0 * $barWidthPx, 1)
                    Status       = $sess.status
                    InWindow     = ($tEo -ge $winStart -and $tSo -le $winEnd)
                    ClippedStart = $clipS
                    ClippedEnd   = $clipE
                }) | Out-Null
            }
        }
    }
    # Build explicit gridline divs at clock-hour boundaries using same pixel math as axis ticks
    $gridLinesDivs = ''
    $gTime = $winStart.AddHours($gridEveryH - ($winStart.Hour % $gridEveryH)).AddMinutes(-$winStart.Minute).AddSeconds(-$winStart.Second)
    while ($gTime -le $winEnd) {
        $gPx = [math]::Round(($gTime - $winStart).TotalSeconds / $winSec * $barWidthPx, 1)
        $gridLinesDivs += "<div class='tl-gl' style='left:${gPx}px'></div>"
        $gTime = $gTime.AddHours($gridEveryH)
    }
    $vpInner  = "<div class='tl-vi' style='width:${barWidthPx}px'>$gridLinesDivs$axBar$barItems$nowLn</div>"
    $vpCol    = "<div class='tl-vc'>$arrowL$arrowR<div class='tl-vp' id='$tlId'>$vpInner</div></div>"
    return "<div class='tl-outer'>$labelCol$vpCol</div>$legend"
}

function Format-RelativeTime {    param([long]$UnixTs)
    if ($UnixTs -le 0) { return 'never' }
    $age = (Get-Date).ToUniversalTime() - [DateTimeOffset]::FromUnixTimeSeconds($UnixTs).UtcDateTime
    if ($age.TotalMinutes -lt 60)  { return "$([int]$age.TotalMinutes)m ago" }
    if ($age.TotalHours   -lt 48)  { return "$([int]$age.TotalHours)h ago" }
    return "$([int]$age.TotalDays)d ago"
}

function Get-SessionAgeColor {
    param([long]$UnixTs)
    if ($UnixTs -le 0) { return '#e74c3c' }   # RED  - no data
    $ageH = ((Get-Date).ToUniversalTime() - [DateTimeOffset]::FromUnixTimeSeconds($UnixTs).UtcDateTime).TotalHours
    if ($ageH -lt 24) { return '#27ae60' }     # GREEN  - < 24 h
    if ($ageH -lt 48) { return '#e67e22' }     # ORANGE - 24-48 h
    return '#e74c3c'                           # RED    - > 48 h
}

function Get-DataSourcesHtml {
    param([hashtable]$LastSessions)
    if (-not $LastSessions -or $LastSessions.Count -eq 0) {
        return "<p style='color:#888;font-size:0.88em'>No completed session data available</p>"
    }

    $nowU = [DateTimeOffset]::new((Get-Date).ToUniversalTime(), [TimeSpan]::Zero).ToUnixTimeSeconds()

    $rows = ''
    foreach ($plugId in ($LastSessions.Keys | Sort-Object)) {
        $plugData = $LastSessions[$plugId]
        $history  = @($plugData.completed)   # completed sessions, sorted DESC
        $failed   = @($plugData.failed)      # failed/error sessions, sorted DESC
        if ($history.Count -eq 0 -and $failed.Count -eq 0) { continue }

        $name  = $PLUGIN_NAMES[$plugId] ?? "Plugin $plugId"
        $last  = if ($history.Count -gt 0) { $history[0] } else { 0 }
        $color = Get-SessionAgeColor $last
        $rel   = Format-RelativeTime  $last

        # Compute median interval between consecutive sessions
        $medianH    = 24.0
        $cadenceTxt = ''
        if ($history.Count -ge 2) {
            $intervals = @()
            for ($i = 0; $i -lt [Math]::Min($history.Count - 1, 20); $i++) {
                $iv = ($history[$i] - $history[$i+1]) / 3600.0
                if ($iv -gt 0) { $intervals += $iv }
            }
            if ($intervals.Count -gt 0) {
                $sorted  = @($intervals | Sort-Object)
                $medianH = $sorted[[int]($sorted.Count / 2)]
                $totalMin   = [int][Math]::Round($medianH * 60)
                $cadenceTxt = if     ($medianH -lt 0.75) { "~${totalMin}m avg" }
                              elseif ($medianH -lt 1.5)  { '~1h avg'  }
                              elseif ($medianH -lt 3)    { '~2h avg'  }
                              elseif ($medianH -lt 5)    { '~4h avg'  }
                              elseif ($medianH -lt 7)    { '~6h avg'  }
                              elseif ($medianH -lt 10)   { '~8h avg'  }
                              elseif ($medianH -lt 18)   { '~12h avg' }
                              elseif ($medianH -lt 36)   { '~24h avg' }
                              else                       { "~$([int]($medianH/24))d avg" }
            }
        }

        # Adaptive window & bucket size
        if     ($medianH -le 4)  { $bucketH = 1;  $windowH = 24;  $label = '24h' }
        elseif ($medianH -le 20) { $bucketH = 4;  $windowH = 72;  $label = '3d'  }
        else                     { $bucketH = 24; $windowH = 336; $label = '14d' }
        $cellCount  = [int]($windowH / $bucketH)
        $bucketSec  = [long]($bucketH * 3600)
        # Align strip end to the next clean bucket boundary so tooltips always show :00
        $stripEnd   = if ($nowU % $bucketSec -eq 0) { $nowU } else { $nowU + ($bucketSec - ($nowU % $bucketSec)) }

        # Success % + RPO pills — filter sessions_full to adaptive window
        $windowStart  = $nowU - [long]($windowH * 3600)
        $sfWindow     = @($plugData.sessions_full | Where-Object { $_.end -ge $windowStart })
        $nCompleted   = @($sfWindow | Where-Object { $_.status -eq 'Completed' }).Count
        $nCleaned     = @($sfWindow | Where-Object { $_.cleaned -and $_.status -ne 'Completed' }).Count
        $nSuccess     = $nCompleted + $nCleaned
        $nWithErrors  = @($sfWindow | Where-Object { $_.status -eq 'CompletedWithErrors' -and -not $_.cleaned }).Count
        $nFailed      = @($sfWindow | Where-Object { $_.status -in @('Failed','Aborted') }).Count
        $nSkipped     = @($sfWindow | Where-Object { $_.status -eq 'Skipped' }).Count
        $nRan         = $nSuccess + $nWithErrors + $nFailed
        $expectedSlots = [int][math]::Floor($windowH / $medianH)
        $missed        = [math]::Max(0, $expectedSlots - $nRan)
        $extra         = [math]::Max(0, $nRan - $expectedSlots)
        $denominator   = $nRan + $missed
        $successPct    = if ($denominator -gt 0) { [math]::Round($nSuccess / $denominator * 100, 0) } else { $null }

        # RPO: median duration of Completed sessions in window
        $durations  = @($sfWindow | Where-Object { $_.status -eq 'Completed' -and $_.start -gt 0 } |
                        ForEach-Object { ($_.end - $_.start) / 3600.0 } |
                        Where-Object { $_ -gt 0 -and $_ -lt ($medianH * 3) })
        $medianDurH = if ($durations.Count -gt 0) { @($durations | Sort-Object)[[int]($durations.Count / 2)] } else { 0.0 }
        $rpoRatio   = if ($medianH -gt 0 -and $medianDurH -gt 0) { $medianDurH / $medianH } else { 0.0 }

        # Success pill
        $pillSuccessHtml = ''
        if ($null -ne $successPct) {
            $sBg = if ($successPct -ge 95) { '#1a3a1a' } elseif ($successPct -ge 80) { '#3a2e0a' } else { '#3a1010' }
            $sFg = if ($successPct -ge 95) { '#6ecf6e' } elseif ($successPct -ge 80) { '#f0a030' } else { '#e05050' }
            $sBd = if ($successPct -ge 95) { '#2a5a2a' } elseif ($successPct -ge 80) { '#5a4a10' } else { '#5a1010' }
            $sTip = "$name - $label window&#10;Completed: $nSuccess  Failed/errors: $($nWithErrors+$nFailed)&#10;Missed slots: $missed  Skipped: $nSkipped (excluded)&#10;Extra/manual: $extra  Expected: $expectedSlots"
            $pillSuccessHtml = "<span title='$sTip' style='display:block;padding:2px 8px;border-radius:10px;font-size:0.85em;font-weight:bold;white-space:nowrap;background:$sBg;color:$sFg;border:1px solid $sBd;text-align:center;cursor:help;margin-bottom:4px'>${successPct}%</span>"
        }

        # RPO pill
        $pillRpoHtml = ''
        if ($rpoRatio -gt 0) {
            $rBg  = if ($rpoRatio -lt 0.5) { '#122030' } elseif ($rpoRatio -lt 0.8) { '#302818' } else { '#301010' }
            $rFg  = if ($rpoRatio -lt 0.5) { '#4a9eda' } elseif ($rpoRatio -lt 0.8) { '#c8943a' } else { '#e04040' }
            $rBd  = if ($rpoRatio -lt 0.5) { '#1a3a5a' } elseif ($rpoRatio -lt 0.8) { '#5a4010' } else { '#5a1010' }
            $rTxt = if ($rpoRatio -lt 0.5) { 'RPO ok' } elseif ($rpoRatio -lt 0.8) { 'RPO ~' } else { 'RPO !' }
            $headroom = Format-HM ([math]::Max(0.0, $medianH - $medianDurH))
            $rTip = "$name RPO - $label window&#10;Median duration: $(Format-HM $medianDurH)  Cadence: $(Format-HM $medianH)&#10;Window consumed: $([math]::Round($rpoRatio*100,0))%  Headroom: $headroom"
            $pillRpoHtml = "<span title='$rTip' style='display:block;padding:2px 8px;border-radius:10px;font-size:0.85em;font-weight:bold;white-space:nowrap;background:$rBg;color:$rFg;border:1px solid $rBd;text-align:center;cursor:help'>$rTxt</span>"
        }

        $pillsTd = "<td style='padding:0 6px;vertical-align:middle'>$pillSuccessHtml$pillRpoHtml</td>"

        # Build strip: prepend each cell so left=oldest, right=newest
        $strip = ''
        for ($c = 0; $c -lt $cellCount; $c++) {
            $bucketEnd   = $stripEnd - ($c * $bucketSec)
            $bucketStart = $bucketEnd - $bucketSec
            $cCount  = @($history | Where-Object { $_ -ge $bucketStart -and $_ -lt $bucketEnd }).Count
            $fCount  = @($failed   | Where-Object { $_ -ge $bucketStart -and $_ -lt $bucketEnd }).Count
            $cellBg  = if ($cCount -gt 0) { '#27ae60' } elseif ($fCount -gt 0) { '#e67e22' } else { '#2d2d2d' }
            $tsStart = [DateTimeOffset]::FromUnixTimeSeconds($bucketStart).UtcDateTime
            $tsEnd   = [DateTimeOffset]::FromUnixTimeSeconds($bucketEnd).UtcDateTime
            $tipTime = if ($bucketH -lt 24) {
                           if ($tsStart.Day -ne $tsEnd.Day) {
                               "$($tsStart.ToString('MMM dd HH:mm'))-$($tsEnd.ToString('HH:mm'))"
                           } else {
                               "$($tsStart.ToString('HH:mm'))-$($tsEnd.ToString('HH:mm'))"
                           }
                       } else {
                           $tsStart.ToString('MMM dd')
                       }
            $tip     = "$tipTime$(if ($cCount -gt 0) { " ($cCount ok)" } elseif ($fCount -gt 0) { " ($fCount failed)" } else { '' })"
            $strip   = "<span title='$tip' style='display:inline-block;width:9px;height:12px;background:$cellBg;border-radius:1px;margin:0 1px;vertical-align:middle'></span>$strip"
        }

        $labelSpan   = "<span style='color:#555;font-size:0.88em;margin-right:4px'>$label</span>"
        $cadenceSpan = if ($cadenceTxt) { "<span style='color:#666;font-size:0.90em;margin-left:6px'>$cadenceTxt</span>" } else { '' }

        $rows += "<tr>"
        $rows += "<td style='padding:6px 10px 6px 0;color:#ccc;font-size:1.0em;white-space:nowrap;vertical-align:middle;font-weight:500'>$name</td>"
        $rows += $pillsTd
        $rows += "<td style='padding-right:8px;white-space:nowrap;vertical-align:middle'>$labelSpan$strip</td>"
        $rows += "<td style='white-space:nowrap;vertical-align:middle'><span style='background:$color;color:#000;padding:2px 8px;border-radius:3px;font-size:0.95em;font-weight:bold'>$rel</span>$cadenceSpan</td>"
        $rows += "</tr>"
    }
    return "<table style='border-collapse:collapse'>$rows</table>"
}

# ============================================================================
# HTML Generation
# ============================================================================

$SEV_COLOR  = @{ RED = "#e74c3c"; YELLOW = "#f39c12"; GREEN = "#27ae60"; OFFLINE = "#7f8c8d" }
$SEV_BG     = @{ RED = "#2c1010"; YELLOW = "#2c2010"; GREEN = "#102c10"; OFFLINE = "#1a1a1a" }
$SEV_BORDER = @{ RED = "#e74c3c"; YELLOW = "#f39c12"; GREEN = "#27ae60"; OFFLINE = "#555" }

function Get-SevBadge {
    param([string]$Sev)
    $c = $SEV_COLOR[$Sev] ?? "#fff"
    return "<span style=`"background:$c;color:#000;padding:2px 8px;border-radius:3px;font-weight:bold;font-size:0.85em`">$Sev</span>"
}

function Get-StorageBar {
    param([double]$PctVal)
    $color = if ($PctVal -ge 90) { "#e74c3c" } elseif ($PctVal -ge 80) { "#f39c12" } else { "#27ae60" }
    $width = [math]::Min($PctVal, 100)
    $pctStr = $PctVal.ToString("F1")
    return @"
<div style="background:#333;border-radius:4px;height:14px;width:100%;position:relative">
  <div style="background:$color;width:$width%;height:14px;border-radius:4px"></div>
  <span style="position:absolute;top:-1px;left:4px;font-size:0.75em;color:#eee">$pctStr%</span>
</div>
"@
}

function Get-WriterTooltip {
    param([string]$Name)
    # Each entry: pattern -> @{purp; svc; start; rst (array); imp}
    # Checked in order; first match wins.
    $tips = [ordered]@{
        'ASR Writer' = @{
            purp  = 'Automated System Recovery - backs up disk geometry, partition layout, and boot configuration (BCD store, MBR/GPT).'
            svc   = 'vssvc.exe (via cryptsvc - Cryptographic Services)'
            start = 'Automatic'
            rst   = @('net stop vss','net start vss')
            imp   = 'Bare-metal and system state restores cannot reconstruct disk layout.'
        }
        'BITS Writer' = @{
            purp  = 'Background Intelligent Transfer Service - backs up BITS job queue and pending transfer state.'
            svc   = 'BITS'
            start = 'Manual (demand-start)'
            rst   = @('net stop BITS','net start BITS')
            imp   = 'Pending Windows Update transfers may be lost on restore; BITS will re-queue them automatically.'
        }
        'COM+ REGDB Writer' = @{
            purp  = 'COM+ Registration Database - backs up COM+ application catalog and activation data in %SystemRoot%\Registration.'
            svc   = 'EventSystem (COM+ Event System)'
            start = 'Automatic'
            rst   = @('net stop EventSystem','net start EventSystem')
            imp   = 'COM+ applications may fail to activate or register after a restore without this writer.'
        }
        'IIS Config Writer' = @{
            purp  = 'IIS configuration - backs up applicationHost.config, site bindings, application pools, and SSL certificate assignments.'
            svc   = 'iisadmin (IIS Admin Service)'
            start = 'Automatic on IIS servers'
            rst   = @('net stop iisadmin /y','net start iisadmin','net start w3svc')
            imp   = 'IIS configuration (sites, bindings, app pools) cannot be restored from backup - only web content survives.'
        }
        'IIS Metabase Writer' = @{
            purp  = 'Legacy IIS 6 metabase (MetaBase.xml) - present on IIS 6 or IIS 7+ in backward-compatibility mode.'
            svc   = 'iisadmin (IIS Admin Service)'
            start = 'Automatic'
            rst   = @('net stop iisadmin /y','net start iisadmin')
            imp   = 'IIS 6 / legacy metabase settings cannot be restored consistently.'
        }
        'Microsoft Hyper-V VSS Writer' = @{
            purp  = 'Hyper-V integration - coordinates online snapshots of running VMs (saved-state or child-VM VSS pass-through). Required for host-level VM backup.'
            svc   = 'vmms (Hyper-V Virtual Machine Management)'
            start = 'Automatic'
            rst   = @('net stop vmms','net start vmms')
            imp   = 'VM backups fall back to crash-consistent (no app consistency) or fail entirely. VMs with no VSS integration fall back to saved-state (brief pause).'
        }
        'MSMQ Writer' = @{
            purp  = 'Microsoft Message Queuing - backs up MSMQ message store and queue configuration.'
            svc   = 'MSMQ (Message Queuing)'
            start = 'Automatic when MSMQ is installed'
            rst   = @('net stop MSMQ','net start MSMQ')
            imp   = 'In-flight messages in MSMQ queues may be lost or duplicated on restore.'
        }
        'Performance Counters Writer' = @{
            purp  = 'Performance counter configuration - backs up PerfLib registry keys and performance counter definitions.'
            svc   = 'vssvc.exe (VSS-internal, no separate service)'
            start = 'Automatic via VSS'
            rst   = @('net stop vss','net start vss')
            imp   = 'Performance Monitor and custom perf counters may be missing after restore (can be rebuilt with lodctr /r).'
        }
        'Registry Writer' = @{
            purp  = 'System Registry - backs up all registry hives: SYSTEM, SOFTWARE, SAM, SECURITY, DEFAULT, and user hives.'
            svc   = 'vssvc.exe (VSS-internal, no separate service)'
            start = 'Automatic via VSS'
            rst   = @('net stop vss','net start vss')
            imp   = 'CRITICAL: registry backup fails silently. System state and bare-metal restores will have an inconsistent or missing registry.'
        }
        'Sentinel Agent Database VSS Writer' = @{
            purp  = 'SentinelOne EDR - application-consistent snapshot of the S1 threat intelligence and detection database.'
            svc   = 'SentinelAgent'
            start = 'Automatic (managed by S1)'
            rst   = @('# Managed by SentinelOne - restart via S1 Management Console','# Do NOT restart SentinelAgent independently without S1 authorization')
            imp   = 'S1 rollback snapshots may be inconsistent; threat database may require re-sync after restore.'
        }
        'Sentinel Agent DFI' = @{
            purp  = 'SentinelOne EDR Deep File Inspection - backs up DFI telemetry and research data store.'
            svc   = 'SentinelAgent'
            start = 'Automatic (managed by S1)'
            rst   = @('# Managed by SentinelOne - restart via S1 Management Console')
            imp   = 'DFI telemetry data may be incomplete after restore; does not affect core S1 protection.'
        }
        'Sentinel Agent Log VSS Writer' = @{
            purp  = 'SentinelOne EDR - consistent snapshot of the S1 agent activity and audit log database.'
            svc   = 'SentinelAgent'
            start = 'Automatic (managed by S1)'
            rst   = @('# Managed by SentinelOne - restart via S1 Management Console')
            imp   = 'S1 rollback snapshots require log consistency for tamper-proof audit trail. Unhealthy state may block rollback.'
        }
        'Sentinel' = @{
            purp  = 'SentinelOne EDR VSS writer - creates application-consistent snapshots enabling tamper-proof rollback on ransomware events.'
            svc   = 'SentinelAgent'
            start = 'Automatic (managed by S1)'
            rst   = @('# Managed by SentinelOne - restart via S1 Management Console','# Do NOT restart independently without S1 console authorization')
            imp   = 'S1 ransomware rollback capability is degraded or unavailable while this writer is unhealthy.'
        }
        'Shadow Copy Optimization Writer' = @{
            purp  = 'Shadow Copy Optimization - excludes files that do not need to be differentially tracked (temp files, page file regions, etc.), reducing shadow storage overhead.'
            svc   = 'vssvc.exe (VSS-internal, no separate service)'
            start = 'Automatic via VSS'
            rst   = @('net stop vss','net start vss')
            imp   = 'Shadow copies remain functional but may be larger than necessary. Low priority issue.'
        }
        'SqlServerWriter' = @{
            purp  = 'SQL Server VSS Writer - enables online, application-consistent backups of all SQL Server databases without detaching or taking them offline.'
            svc   = 'SQLWriter (SQL Server VSS Writer)'
            start = 'Manual (demand-start, launched on snapshot request)'
            rst   = @('net stop sqlwriter','net stop vss','net start vss','net start sqlwriter','vssadmin list writers')
            imp   = 'CRITICAL: SQL backups fall back to crash-consistent (data may be corrupt on restore) or are skipped entirely depending on backup software settings. All databases on this instance are affected.'
        }
        'System Writer' = @{
            purp  = 'System files - backs up core OS binaries, WinSxS component store, %SystemRoot%\System32, signed catalog files, and PnP driver store.'
            svc   = 'cryptsvc (Cryptographic Services) - also hosts ASR Writer'
            start = 'Automatic'
            rst   = @('net stop cryptsvc','net start cryptsvc')
            imp   = 'CRITICAL: OS file backup is incomplete. System state and bare-metal restores may produce an unbootable system or require SFC /scannow after restore.'
        }
        'Task Scheduler Writer' = @{
            purp  = 'Task Scheduler - backs up all scheduled task XML definitions from %SystemRoot%\System32\Tasks.'
            svc   = 'Schedule (Task Scheduler)'
            start = 'Automatic'
            rst   = @('net stop Schedule','net start Schedule')
            imp   = 'Scheduled tasks (maintenance, monitoring, backup jobs) will be missing after a system state restore.'
        }
        'VSS Metadata Store Writer' = @{
            purp  = 'VSS Metadata Store - backs up VSS writer registration data (GUIDs, component topology, writer metadata) needed to interpret backup sets on restore.'
            svc   = 'vssvc.exe (VSS-internal, no separate service)'
            start = 'Automatic via VSS'
            rst   = @('net stop vss','net start vss')
            imp   = 'Restore tools may not be able to enumerate or restore individual components from the backup set.'
        }
        'WIDWriter' = @{
            purp  = 'Windows Internal Database (WID) - backs up the embedded SQL Server instance used by WSUS, AD FS, Windows Server Update Services, and Device Health.'
            svc   = 'MSSQL$MICROSOFT##WID'
            start = 'Automatic when WID is installed'
            rst   = @('net stop MSSQL$MICROSOFT##WID','net start MSSQL$MICROSOFT##WID')
            imp   = 'WSUS update database, AD FS configuration, or DHA data may be inconsistent on restore.'
        }
        'WMI Writer' = @{
            purp  = 'WMI Repository - backs up the WMI object repository at %SystemRoot%\System32\wbem\Repository.'
            svc   = 'winmgmt (Windows Management Instrumentation)'
            start = 'Automatic'
            rst   = @('net stop winmgmt','net start winmgmt')
            imp   = 'WMI-dependent tools (monitoring agents, PowerShell Get-WmiObject, SCCM client, etc.) may fail after restore until WMI is rebuilt.'
        }
    }
    foreach ($pat in $tips.Keys) {
        if ($Name -like "*$pat*") { return $tips[$pat] }
    }
    return @{
        purp  = "VSS writer: $Name"
        svc   = 'Unknown'
        start = 'Unknown'
        rst   = @('net stop vss','net start vss','vssadmin list writers')
        imp   = 'Impact unknown. Check vssadmin list writers for details after restart.'
    }
}

function Get-SystemEventsSection {
    param([array]$Events, [string]$DeviceTz = '')

    if (-not $Events -or $Events.Count -eq 0) {
        return "<p style='color:#555;font-size:0.85em'>No Error or Warning events found in the collected event log data.</p>"
    }

    $tzNote = if ($DeviceTz) { " <span style='color:#555'>(device time: $DeviceTz)</span>" } else { '' }
    $rows = ''
    foreach ($ev in $Events) {
        $lvlColor = if ($ev.Level -eq 'Error') { '#e74c3c' } else { '#f39c12' }
        $lvlIcon  = if ($ev.Level -eq 'Error') { '&#9888;' } else { '&#9432;' }
        $countBadge = if ($ev.Count -gt 1) {
            "<span style='margin-left:6px;padding:1px 5px;background:#2a2a2a;border:1px solid #444;border-radius:3px;font-size:0.78em;color:#888'>x$($ev.Count)</span>"
        } else { '' }

        # Friendly provider highlight for known Cove/TakeControl entries
        $provDisp = $ev.Provider
        if ($ev.Provider -match 'Cove|TakeControl|BeAnywhere|BackupFP') {
            $provDisp = "<span style='color:#4ecdc4'>$($ev.Provider)</span>"
        } elseif ($ev.Provider -match 'Service Control Manager') {
            $provDisp = "<span style='color:#888'>$($ev.Provider)</span>"
        }

        # Encode message for Google AI search URL (udm=50 forces AI mode)
        $encodedMsg = [System.Web.HttpUtility]::UrlEncode($ev.Message)
        $googleAiLink = "https://www.google.com/search?udm=50&q=$encodedMsg+troubleshoot"
        
        $rows += "<tr>
          <td style='color:#666;white-space:nowrap;padding:3px 8px;vertical-align:top;font-size:0.82em'>$($ev.Time)</td>
          <td style='padding:3px 8px;vertical-align:top'><span style='color:$lvlColor'>$lvlIcon $($ev.Level)</span></td>
          <td style='padding:3px 8px;vertical-align:top;font-size:0.85em'>$provDisp</td>
          <td style='padding:3px 8px;color:#ccc;font-size:0.85em'>$($ev.Message)<a href='$googleAiLink' target='_blank' style='margin-left:8px;color:#4ecdc4;text-decoration:none;cursor:pointer;font-weight:bold' title='Search with Google AI for troubleshooting'>🔍</a>$countBadge</td>
        </tr>"
    }

    return @"
<div style='font-size:0.8em;color:#555;margin-bottom:6px'>Most recent errors and warnings, deduplicated$tzNote</div>
<table style='width:100%;border-collapse:collapse;font-size:0.88em'>
  <thead><tr style='background:#2a2a2a'>
    <th style='text-align:left;padding:4px 8px;color:#888;width:130px'>Time</th>
    <th style='text-align:left;padding:4px 8px;color:#888;width:70px'>Level</th>
    <th style='text-align:left;padding:4px 8px;color:#888;width:180px'>Provider / Source</th>
    <th style='text-align:left;padding:4px 8px;color:#888'>Message</th>
  </tr></thead>
  <tbody>$rows</tbody>
</table>
"@
}

function Get-VssCorrelationSection {
    param(
        [array]$Writers,
        [hashtable]$ServicesMap,
        [bool]$BackupInProgress,
        [hashtable]$LastSessions,
        [array]$Volumes,
        [hashtable]$SnapsByDrive,
        [bool]$S1Installed
    )
    $ServicesMap  = if ($ServicesMap)  { $ServicesMap }  else { @{} }
    $LastSessions = if ($LastSessions) { $LastSessions } else { @{} }
    $SnapsByDrive = if ($SnapsByDrive) { $SnapsByDrive } else { @{} }
    $nowU = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    # -------------------------------------------------------------------------
    # Phase 3 context: last snapshot age, last backup age
    # -------------------------------------------------------------------------
    $lastSnapDt = $null
    foreach ($drv in $SnapsByDrive.Keys) {
        foreach ($sn in $SnapsByDrive[$drv]) {
            if ($null -ne $sn.ts -and ($null -eq $lastSnapDt -or $sn.ts -gt $lastSnapDt)) {
                $lastSnapDt = $sn.ts
            }
        }
    }
    $lastSnapAgeH = if ($lastSnapDt) { [math]::Round(($nowU - [DateTimeOffset]::new($lastSnapDt).ToUnixTimeSeconds()) / 3600.0, 1) } else { $null }

    $lastBkpU = $null
    foreach ($pd in $LastSessions.Values) {
        $comp = @($pd.completed | Where-Object { $_ })
        if ($comp.Count -gt 0 -and ($null -eq $lastBkpU -or $comp[0] -gt $lastBkpU)) { $lastBkpU = $comp[0] }
    }
    $lastBkpAgeH = if ($lastBkpU) { [math]::Round(($nowU - $lastBkpU) / 3600.0, 1) } else { $null }

    # -------------------------------------------------------------------------
    # Phase 1: VSS Infrastructure Health
    # Each entry: key (service name lower), displayName, startup, expectedRuntime
    #   expectedRuntime: 'running' | 'stopped' | 'trigger'
    # -------------------------------------------------------------------------
    $infraKB = [ordered]@{
        'vss'         = @{ disp='Volume Shadow Copy';                      startup='Manual';    expected='stopped'; note='Started on demand during snapshot creation. Running at rest = hung snapshot set.' }
        'swprv'       = @{ disp='MS Software Shadow Copy Provider';        startup='Manual';    expected='stopped'; note='Manages in-process software shadow copies. Should stop when snapshot lifecycle completes.' }
        'rpcss'       = @{ disp='Remote Procedure Call (RPC)';             startup='Automatic'; expected='running'; note='Core Windows IPC. VSS, WMI, COM+ all depend on this. Stopped = catastrophic.' }
        'dcomlaunch'  = @{ disp='DCOM Server Process Launcher';            startup='Automatic'; expected='running'; note='Launches COM/DCOM servers including VSS components. Must be running.' }
        'eventsystem' = @{ disp='COM+ Event System';                       startup='Automatic'; expected='running'; note='Required by COM+ REGDB Writer and VSS provider event notifications.' }
        'comsysapp'   = @{ disp='COM+ System Application';                 startup='Manual';    expected='stopped'; note='COM+ application host. Started on demand; not required at rest.' }
        'winmgmt'     = @{ disp='Windows Management Instrumentation';      startup='Automatic'; expected='running'; note='WMI Writer host. Monitoring agents, PowerShell Get-WmiObject, SCCM all depend on this.' }
        'cryptsvc'    = @{ disp='Cryptographic Services';                  startup='Automatic'; expected='running'; note='Hosts System Writer and ASR Writer. Required for system state and bare-metal backup.' }
    }

    $recommendations = [System.Collections.Generic.List[string]]::new()
    $infraRows = ''
    foreach ($svcKey in $infraKB.Keys) {
        $kb   = $infraKB[$svcKey]
        $info = if ($ServicesMap.ContainsKey($svcKey)) { $ServicesMap[$svcKey] } else { $null }
        $state   = if ($info) { $info.state } else { $null }
        # Startup: prefer live data; fall back to KB (InternalInfo rarely exposes startup type)
        $startup = if ($info -and $info.startup) { $info.startup } else { $kb.startup }

        # Determine health
        $healthColor = '#555'; $healthLabel = 'Not Detected'
        if ($info) {
            $isRunning = $state -match 'Running'
            $isStopped = -not $isRunning
            $expectRun = $kb.expected -eq 'running'

            if ($expectRun -and $isRunning) {
                $healthColor = '#27ae60'; $healthLabel = 'Healthy'
            } elseif ($expectRun -and $isStopped) {
                $healthColor = '#e74c3c'; $healthLabel = 'Critical'
                $recommendations.Add("CRITICAL: $($kb.disp) ($svcKey) is Stopped but should be Automatic/Running. Restart: net start $svcKey") | Out-Null
            } elseif (-not $expectRun -and $isStopped) {
                $healthColor = '#27ae60'; $healthLabel = 'Healthy'
            } elseif (-not $expectRun -and $isRunning) {
                # Manual service running - check context
                # Suspicious if: no active backup, last snap stale (>2h), and last backup completed a while ago
                $snapStale  = $null -ne $lastSnapAgeH  -and $lastSnapAgeH  -gt 2
                $bkpStale   = $null -eq $lastBkpAgeH   -or  $lastBkpAgeH   -gt 2
                $isSuspicious = -not $BackupInProgress -and $snapStale -and $bkpStale
                if ($BackupInProgress) {
                    $healthColor = '#27ae60'; $healthLabel = 'Expected (backup active)'
                } elseif ($null -ne $lastSnapAgeH -and $lastSnapAgeH -lt 2) {
                    $healthColor = '#4ecdc4'; $healthLabel = 'OK (recent snap)'
                } elseif ($isSuspicious) {
                    $healthColor = '#f39c12'; $healthLabel = 'Suspicious'
                    $snapNote = if ($null -ne $lastSnapAgeH) { "$([int]$lastSnapAgeH)h since last snapshot" } else { 'no recent snapshot' }
                    $recommendations.Add("WARNING: $($kb.disp) ($svcKey) is Manual-start but Running with no active backup ($snapNote). May indicate an unreleased snapshot set. Run: vssadmin list shadows") | Out-Null
                } else {
                    $healthColor = '#5dade2'; $healthLabel = 'Informational'
                }
            }
        } else {
            $recommendations.Add("INFO: $($kb.disp) ($svcKey) not found in service inventory - InternalInfo may be incomplete.") | Out-Null
        }

        $stateDisp   = if ($state)   { $state }   else { '<span style="color:#444">Not detected</span>' }
        $startupDisp = if ($startup) {
            $col = if ($startup -match 'Manual') { '#f39c12' } else { '#888' }
            "<span style='color:$col'>$startup</span>"
        } else { '<span style="color:#444">-</span>' }
        $expectedDisp = if ($kb.expected -eq 'running') { '<span style="color:#27ae60">Running</span>' } else { '<span style="color:#888">Stopped</span>' }
        $healthBadge  = "<span style='color:$healthColor;font-weight:bold'>$healthLabel</span>"
        $noteTip = ($kb.note -replace "'", '&apos;')

        $infraRows += "<tr title='$noteTip' style='cursor:help'>
          <td style='color:#ccc'>$svcKey</td>
          <td style='color:#aaa'>$($kb.disp)</td>
          <td>$startupDisp</td>
          <td style='color:#aaa'>$stateDisp</td>
          <td>$expectedDisp</td>
          <td>$healthBadge</td>
        </tr>"
    }

    $infraHtml = @"
<table style="width:100%;border-collapse:collapse;font-size:0.85em;margin-bottom:12px">
  <thead><tr style="background:#2a2a2a">
    <th style="text-align:left;padding:4px 8px;color:#888;width:100px">Service</th>
    <th style="text-align:left;padding:4px 8px;color:#888">Display Name</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:90px">Startup</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:110px">Current State</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:90px">Expected</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:160px">Health</th>
  </tr></thead>
  <tbody>$infraRows</tbody>
</table>
"@

    # -------------------------------------------------------------------------
    # Phase 2: Writer Correlation Matrix
    # -------------------------------------------------------------------------
    # Vendor writer patterns for special handling
    $vendorPatterns = [ordered]@{
        'sentinel\s+agent|sentinel[\s-]?one' = @{ vendor='SentinelOne'; svcKey='sentinelagent'; note='Managed by S1 agent. Do not restart independently.' }
        'veeam'                              = @{ vendor='Veeam';        svcKey='veeamguesthelper'; note='Veeam in-guest processing writer. Requires Veeam agent.' }
        'commvault'                          = @{ vendor='Commvault';    svcKey='gxclmgrS(2)'; note='Commvault Galaxy writer.' }
        'acronis'                            = @{ vendor='Acronis';      svcKey='acronisagent'; note='Acronis backup agent writer.' }
        'mabs|microsoft azure backup'        = @{ vendor='Azure Backup'; svcKey='cbengine'; note='Microsoft Azure Backup writer.' }
        'storagecraft'                       = @{ vendor='StorageCraft';  svcKey='imagingservice'; note='StorageCraft ShadowProtect writer.' }
        'rubrik'                             = @{ vendor='Rubrik';       svcKey='rubrikbackupservice'; note='Rubrik Backup Service writer.' }
    }

    # Static startup-type KB for services not exposed in InternalInfo servicesinfo
    $startupKB = @{
        'vss'                    = 'Manual'
        'swprv'                  = 'Manual'
        'bits'                   = 'Manual'
        'rpcss'                  = 'Automatic'
        'dcomlaunch'             = 'Automatic'
        'eventsystem'            = 'Automatic'
        'comsysapp'              = 'Manual'
        'winmgmt'                = 'Automatic'
        'cryptsvc'               = 'Automatic'
        'iisadmin'               = 'Automatic'
        'sqlwriter'              = 'Manual'
        'schedule'               = 'Automatic'
        'sentinelagent'          = 'Automatic'
        'sentinelstaticengine'   = 'Automatic'
        'vmms'                   = 'Automatic'
    }

    # Accumulators for consolidated writer recommendations (grouped by root action)
    $vssInternalFailed  = [System.Collections.Generic.List[string]]::new()  # writers fixed by VSS restart alone
    $svcRestartNeeded   = [ordered]@{}   # keyed by svcKey; value = @{writers=[]; svcState=''; startup=''}
    $svcStoppedFailed   = [ordered]@{}   # Auto-start svc stopped -> writer failed
    $vendorFailed       = [System.Collections.Generic.List[string]]::new()
    $unknownFailed      = [System.Collections.Generic.List[string]]::new()

    $writerRows = ''
    foreach ($w in ($Writers | Sort-Object { $_.name })) {
        $svcKey  = ($w.service_name ?? '').ToLower()
        $svcInfo = if ($svcKey -and $ServicesMap.ContainsKey($svcKey)) { $ServicesMap[$svcKey] } else { $null }
        $state   = if ($svcInfo) { $svcInfo.state }   else { $null }
        # Startup: prefer live data; fall back to static KB
        $startup = if ($svcInfo -and $svcInfo.startup) { $svcInfo.startup }
                   elseif ($startupKB.ContainsKey($svcKey)) { $startupKB[$svcKey] }
                   else { $null }

        # Determine if VSS-internal (no separate service needed)
        $isVssInternal = ($svcKey -eq 'vss' -or -not $svcKey)

        # Detect vendor writer
        $isVendor = $false; $vendorNote = ''
        foreach ($pat in $vendorPatterns.Keys) {
            if ($w.name -match $pat) { $isVendor = $true; $vendorNote = $vendorPatterns[$pat].note; break }
        }

        # Expected service state
        $expectedState = if ($isVssInternal) { 'N/A (VSS-internal)' }
                         elseif ($startup -match 'Auto') { 'Running' }
                         elseif ($startup -match 'Manual') { 'Stopped at rest' }
                         else { 'Unknown' }

        # Writer status classification
        $isWaiting   = $w.status -match '^Waiting for'
        $writerOk    = $w.status -eq 'Stable' -or $isWaiting
        $writerFailed= -not $writerOk

        # Assessment
        $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'No action required.'

        if ($isVssInternal -and $writerOk) {
            $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'No action required.'
        } elseif ($isVssInternal -and $writerFailed) {
            $assessment = 'Critical'; $assessColor = '#e74c3c'
            $action = "Restart VSS: net stop vss && net start vss"
            $vssInternalFailed.Add($w.name) | Out-Null
        } elseif ($isVendor -and $writerOk) {
            $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'No action required.'
        } elseif ($isVendor -and $writerFailed) {
            $assessment = 'Critical'; $assessColor = '#e74c3c'
            $action = "Check vendor agent status. $vendorNote"
            $vendorFailed.Add("$($w.name): $vendorNote") | Out-Null
        } elseif ($writerOk -and ($null -eq $state -or -not $svcKey)) {
            $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'No action required.'
        } elseif ($writerOk -and $state -match 'Running') {
            if ($startup -match 'Manual' -and -not $BackupInProgress -and $null -ne $lastSnapAgeH -and $lastSnapAgeH -gt 72) {
                $assessment = 'Warning'; $assessColor = '#f39c12'
                $action = "Manual-start service running with no recent activity ($([int]$lastSnapAgeH)h). Verify no stuck snapshot set."
            } else {
                $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'No action required.'
            }
        } elseif ($writerOk -and $state -match 'Stopped') {
            if ($startup -match 'Auto') {
                $assessment = 'Warning'; $assessColor = '#f39c12'
                $action = "Service stopped but startup is Automatic. Restart: net start $svcKey"
                if (-not $svcStoppedFailed.Contains($svcKey)) { $svcStoppedFailed[$svcKey] = [System.Collections.Generic.List[string]]::new() }
                $svcStoppedFailed[$svcKey].Add($w.name) | Out-Null
            } else {
                $assessment = 'Healthy'; $assessColor = '#27ae60'; $action = 'Manual-start service stopped at rest - normal.'
            }
        } elseif ($writerFailed -and $state -match 'Running') {
            $assessment = 'Critical'; $assessColor = '#e74c3c'
            $action = "Service is Running but writer failed - internal writer error. Restart service: net stop $svcKey && net start $svcKey, then: net stop vss && net start vss"
            if (-not $svcRestartNeeded.Contains($svcKey)) { $svcRestartNeeded[$svcKey] = [System.Collections.Generic.List[string]]::new() }
            $svcRestartNeeded[$svcKey].Add($w.name) | Out-Null
        } elseif ($writerFailed -and $state -match 'Stopped') {
            if ($startup -match 'Auto') {
                $assessment = 'Critical'; $assessColor = '#e74c3c'
                $action = "Service stopped (should be running) caused writer failure. Start: net start $svcKey"
                if (-not $svcStoppedFailed.Contains($svcKey)) { $svcStoppedFailed[$svcKey] = [System.Collections.Generic.List[string]]::new() }
                $svcStoppedFailed[$svcKey].Add($w.name) | Out-Null
            } else {
                $assessment = 'Warning'; $assessColor = '#f39c12'
                $action = "Start service then restart VSS: net start $svcKey && net stop vss && net start vss"
                if (-not $svcRestartNeeded.Contains($svcKey)) { $svcRestartNeeded[$svcKey] = [System.Collections.Generic.List[string]]::new() }
                $svcRestartNeeded[$svcKey].Add($w.name) | Out-Null
            }
        } elseif ($writerFailed) {
            $assessment = 'Warning'; $assessColor = '#f39c12'
            $action = "Writer failed. Service state unknown. Run: vssadmin list writers && net stop vss && net start vss"
            $unknownFailed.Add($w.name) | Out-Null
        } elseif ($isWaiting) {
            $assessment = 'Informational'; $assessColor = '#5dade2'; $action = 'Backup in progress - transient state, no action needed.'
        }

        $wrStatusColor = if ($isWaiting) { '#666' } elseif ($writerOk) { '#27ae60' } else { '#e74c3c' }
        $wrStatusWeight = if ($writerFailed -and -not $isWaiting) { 'bold' } else { 'normal' }
        $stateDisp   = if ($state)   { $state }   else { '<span style="color:#444">-</span>' }
        $startupDisp = if ($startup) {
            $col = if ($startup -match 'Manual') { '#f39c12' } else { '#888' }
            "<span style='color:$col'>$startup</span>"
        } else { '<span style="color:#444">-</span>' }
        $svcDisp     = if ($svcKey -and $svcKey -ne 'vss') { $svcKey } elseif ($isVssInternal) { '<span style="color:#555">VSS-internal</span>' } else { '<span style="color:#444">-</span>' }

        $usageDisp   = if ($w.usage) { "<span style='color:#666;font-size:0.9em'>$($w.usage)</span>" } else { '' }

        # Rich mouseover tooltip (same format as _wrShowTip)
        $tipData = Get-WriterTooltip $w.name
        $tipData['n'] = $w.name
        $tipData['assessment'] = $assessment
        $tipData['action'] = $action
        $tipJson = ($tipData | ConvertTo-Json -Compress) -replace "'", '&apos;'

        # Fold Last Error into Status cell — only shown when non-empty
        $errInline = if ($w.error -and $w.error -ne 'No error') {
            "<div style='color:#e74c3c;font-size:0.8em;margin-top:1px'>$($w.error)</div>"
        } else { '' }

        $writerRows += "<tr data-wtip='$tipJson' onmouseenter='_wrShowTip(event,this)' onmousemove='_tlMoveTip(event)' onmouseleave='_tlHideTip()' style='cursor:help'>
          <td style='color:#ccc;white-space:nowrap'>$($w.name)</td>
          <td style='white-space:nowrap'><span style='color:$wrStatusColor;font-weight:$wrStatusWeight'>$($w.status)</span>$errInline</td>
          <td style='color:#555;font-size:0.82em;white-space:nowrap'>$usageDisp</td>
          <td style='white-space:nowrap'>$svcDisp</td>
          <td style='white-space:nowrap'>$startupDisp</td>
          <td style='color:#aaa;white-space:nowrap'>$stateDisp</td>
          <td style='color:#666;font-size:0.82em;white-space:nowrap'>$expectedState</td>
          <td><span style='color:$assessColor;font-weight:bold'>$assessment</span></td>
        </tr>"
    }

    $writerMatrixHtml = @"
<table style="width:100%;border-collapse:collapse;font-size:0.85em;margin-bottom:12px;table-layout:auto">
  <thead><tr style="background:#2a2a2a">
    <th style="text-align:left;padding:4px 8px;color:#888">Writer</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:160px">Status / Error</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:110px">Usage</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:110px">Service</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:75px">Startup</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:80px">Svc State</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:130px">Expected</th>
    <th style="text-align:left;padding:4px 8px;color:#888;width:90px">Assessment</th>
  </tr></thead>
  <tbody>$writerRows</tbody>
</table>
"@

    # -------------------------------------------------------------------------
    # Phase 3: Activity Context Banner
    # -------------------------------------------------------------------------
    $ctxParts = [System.Collections.Generic.List[string]]::new()
    if ($BackupInProgress) {
        $ctxParts.Add("<span style='color:#4ecdc4'>&#9679; Backup in progress</span>") | Out-Null
    }
    if ($null -ne $lastSnapAgeH) {
        $snapCol = if ($lastSnapAgeH -lt 6) { '#27ae60' } elseif ($lastSnapAgeH -lt 24) { '#f39c12' } else { '#e74c3c' }
        $snapAge = if ($lastSnapAgeH -ge 48) { "$([int]($lastSnapAgeH/24))d ago" } else { "$([int]$lastSnapAgeH)h ago" }
        $ctxParts.Add("<span style='color:#888'>Last snapshot: <span style='color:$snapCol'>$snapAge</span></span>") | Out-Null
    } else {
        $ctxParts.Add("<span style='color:#888'>Last snapshot: <span style='color:#555'>none detected</span></span>") | Out-Null
        if ($S1Installed) {
            $recommendations.Add("WARNING: S1 is installed but no snapshots detected in the monitoring window. Verify SentinelOne snapshot policy is active.") | Out-Null
        }
    }
    if ($null -ne $lastBkpAgeH) {
        $bkpCol = if ($lastBkpAgeH -lt 25) { '#27ae60' } elseif ($lastBkpAgeH -lt 48) { '#f39c12' } else { '#e74c3c' }
        $bkpAge = if ($lastBkpAgeH -ge 48) { "$([int]($lastBkpAgeH/24))d ago" } else { "$([int]$lastBkpAgeH)h ago" }
        $ctxParts.Add("<span style='color:#888'>Last backup: <span style='color:$bkpCol'>$bkpAge</span></span>") | Out-Null
    } else {
        $ctxParts.Add("<span style='color:#888'>Last backup: <span style='color:#555'>none found</span></span>") | Out-Null
    }
    $ctxHtml = "<div style='font-size:0.83em;margin-bottom:10px;display:flex;gap:18px;flex-wrap:wrap;padding:6px 8px;background:#1a1a1a;border-radius:4px;border:1px solid #2a2a2a'>$($ctxParts -join '')</div>"

    # -------------------------------------------------------------------------
    # Phase 4: Consolidated Recommendations
    # Emit one bullet per root action, listing all affected writers under each.
    # Order: service-restart-needed (svc running, writer failed) first (most actionable),
    #        then stopped auto-start services, then VSS-internal-only restart, then vendor, then infra.
    # -------------------------------------------------------------------------
    $recItems = ''

    # Helper: format a writer name list as inline italic chips
    $fmtWriters = { param($list) ($list | ForEach-Object { "<em style='color:#aaa'>$_</em>" }) -join ', ' }

    # 1. All writer failures that require a VSS restart (svcRestartNeeded + vssInternalFailed)
    #    collapse into ONE bullet — you can't restart VSS multiple times; it's one sequence.
    $vssRestartWriters = [System.Collections.Generic.List[string]]::new()
    foreach ($sk in $svcRestartNeeded.Keys)  { foreach ($wn in $svcRestartNeeded[$sk])  { $vssRestartWriters.Add($wn) | Out-Null } }
    foreach ($wn in $vssInternalFailed)       { $vssRestartWriters.Add($wn) | Out-Null }
    if ($vssRestartWriters.Count -gt 0) {
        $wList = & $fmtWriters $vssRestartWriters
        $recItems += "<li style='margin-bottom:10px'><span style='color:#e74c3c;font-weight:bold;font-size:0.8em'>CRITICAL</span> <span style='color:#ccc'>VSS restart required &mdash; stop all dependent services, restart VSS, then start them back.</span><div style='margin-top:3px;color:#888;font-size:0.88em'>Fixes: $wList</div><div style='margin-top:3px;color:#555;font-size:0.82em'>See <strong style='color:#aaa'>Remediation Script</strong> below for the exact stop/start sequence.</div></li>"
    }

    # 2. Auto-start services stopped (writer failed because service is down — start service, may not need full VSS restart)
    if ($svcStoppedFailed.Count -gt 0) {
        $allStopped = [System.Collections.Generic.List[string]]::new()
        $allWriters = [System.Collections.Generic.List[string]]::new()
        foreach ($sk in $svcStoppedFailed.Keys) {
            $allStopped.Add("<code style='color:#4ecdc4'>$sk</code>") | Out-Null
            foreach ($wn in $svcStoppedFailed[$sk]) { $allWriters.Add($wn) | Out-Null }
        }
        $wList = & $fmtWriters $allWriters
        $recItems += "<li style='margin-bottom:10px'><span style='color:#e74c3c;font-weight:bold;font-size:0.8em'>CRITICAL</span> <span style='color:#ccc'>Start stopped Automatic-start service(s): $($allStopped -join ', ')</span><div style='margin-top:3px;color:#888;font-size:0.88em'>Fixes: $wList</div></li>"
    }

    # 3. Vendor/S1 writers — group by note text so identical messages collapse
    $vendorGroups = [ordered]@{}
    foreach ($vf in $vendorFailed) {
        # format: "WriterName: note text"
        $colon = $vf.IndexOf(':')
        $wName = if ($colon -gt 0) { $vf.Substring(0, $colon).Trim() } else { $vf }
        $note  = if ($colon -gt 0) { $vf.Substring($colon+1).Trim() } else { '' }
        if (-not $vendorGroups.Contains($note)) { $vendorGroups[$note] = [System.Collections.Generic.List[string]]::new() }
        $vendorGroups[$note].Add($wName) | Out-Null
    }
    foreach ($note in $vendorGroups.Keys) {
        $wList = & $fmtWriters $vendorGroups[$note]
        $recItems += "<li style='margin-bottom:10px'><span style='color:#e74c3c;font-weight:bold;font-size:0.8em'>CRITICAL</span> <span style='color:#ccc'>$wList &mdash; $note</span></li>"
    }

    # 4. Unknown state failures
    if ($unknownFailed.Count -gt 0) {
        $wList = & $fmtWriters $unknownFailed
        $recItems += "<li style='margin-bottom:10px'><span style='color:#f39c12;font-weight:bold;font-size:0.8em'>WARNING</span> <span style='color:#ccc'>Writers failed with unknown service state &mdash; $wList</span><div style='margin-top:3px;color:#555;font-size:0.82em'>Run <code style='color:#4ecdc4'>vssadmin list writers</code> then see Remediation Script.</div></li>"
    }

    # 5. Infra-level recommendations (vss/swprv suspicious, stopped auto-start infra services)
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in @($recommendations | Sort-Object { if ($_ -match '^CRITICAL') { 0 } elseif ($_ -match '^WARNING') { 1 } else { 2 } })) {
        if ($seen.Add($r)) {
            $col  = if ($r -match '^CRITICAL') { '#e74c3c' } elseif ($r -match '^WARNING') { '#f39c12' } else { '#5dade2' }
            $sev  = if ($r -match '^CRITICAL') { 'CRITICAL' } elseif ($r -match '^WARNING') { 'WARNING' } else { 'INFO' }
            $text = $r -replace '^(CRITICAL|WARNING|INFO):\s*', ''
            $recItems += "<li style='margin-bottom:8px'><span style='color:$col;font-weight:bold;font-size:0.8em'>$sev</span> <span style='color:#ccc'>$text</span></li>"
        }
    }

    $recsHtml = if ($recItems) {
        "<div style='margin-top:10px'><div style='font-size:0.78em;color:#555;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px'>Recommended Actions</div><ol style='margin:0;padding-left:18px;font-size:0.85em;line-height:1.8'>$recItems</ol></div>"
    } else {
        "<div style='font-size:0.85em;color:#27ae60;margin-top:8px'>&#10003; No recommended actions - VSS dependency health looks normal.</div>"
    }

    return @"
$ctxHtml
<div style="font-size:0.82em;color:#555;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">VSS Infrastructure Services</div>
$infraHtml
<div style="font-size:0.82em;color:#555;text-transform:uppercase;letter-spacing:0.06em;margin:10px 0 6px">VSS Writers <span style="color:#444;font-weight:normal;font-size:0.9em">&mdash; hover row for purpose, restart sequence &amp; impact</span></div>
$writerMatrixHtml
$recsHtml
"@
}

function Get-VolumesSection {
    param([array]$Volumes, [bool]$ShowS1 = $false, [hashtable]$GuidMap = @{})
    if (-not $Volumes -or $Volumes.Count -eq 0) {
        return "<p style='color:#888'>No volume data available</p>"
    }
    $html = ""
    foreach ($vs in $Volumes) {
        # Resolve drive label: shadow storage gives drive letter; unmatched GUIDs fall back to disk config GUID map
        $d = $vs.drive
        if ($d -eq '_' -or $d -eq '?' -or -not $d) {
            if ($vs.guid -match '\{([0-9a-fA-F\-]{36})\}') {
                $guidKey = $Matches[1].ToLower()
                if ($GuidMap.ContainsKey($guidKey)) {
                    $d = $GuidMap[$guidKey]
                } else {
                    $d = "Vol {$($Matches[1].Substring(0,8))...}"
                }
            } else { $d = 'Unknown volume' }
        }

        $hasS1Data = $vs.s1_count -gt 0

        $covColor = if ($vs.s1_count -le 1) { "#f39c12" }
                    elseif ($vs.s1_coverage_days -lt 7) { "#e74c3c" }
                    elseif ($vs.s1_coverage_days -lt 10) { "#f39c12" }
                    else { "#27ae60" }

        $behindHtml = ""
        if ($vs.behind_schedule) {
            $hrs = [math]::Round($vs.minutes_since_newest / 60, 1)
            $col = if ($hrs -gt 8) { "#e74c3c" } else { "#f39c12" }
            $behindHtml = "<span style=`"color:$col;margin-left:8px`">&#9888; ${hrs}h behind schedule</span>"
        }

        $covStr  = $vs.s1_coverage_days.ToString("F1")

        $cadenceHtml = if ($ShowS1 -and $hasS1Data -and $null -ne $vs.s1_median_interval_h) {
            $cadColor = if ($vs.s1_median_interval_h -gt 6) { "#f39c12" } else { "#27ae60" }
            " &nbsp;|&nbsp; cadence: <span style='color:$cadColor'>~$(Format-HM $vs.s1_median_interval_h) typical</span>"
        } else { "" }

        $snapCounts = "$(if($vs.s1_count -gt 0 -or $ShowS1){"S1: <span style='color:#4ecdc4'>$($vs.s1_count)</span> &nbsp;|&nbsp; "})$(if($vs.writer_count){"Writer: <span style='color:#a29bfe'>$($vs.writer_count)</span> &nbsp;|&nbsp; "})$(if($vs.cove_count){"Cove: <span style='color:#74b9ff'>$($vs.cove_count)</span> &nbsp;|&nbsp; "})$(if($vs.native_count){"Native: <span style='color:#fd79a8'>$($vs.native_count)</span> &nbsp;|&nbsp; "})$(if($vs.other_count){"Other: <span style='color:#636e72'>$($vs.other_count)</span> &nbsp;|&nbsp; "})$(if($vs.unknown_count){"Unknown: <span style='color:#fdcb6e'>$($vs.unknown_count)</span> &nbsp;|&nbsp; "})Orphans: <span style='color:$(if($vs.orphan_count){"#e74c3c"}else{"#888"})'>$($vs.orphan_count)</span>"

        $s1RowHtml = if ($ShowS1 -and $hasS1Data) { @"
  <div style="display:flex;justify-content:space-between;margin-top:8px;font-size:0.85em">
    <span>S1 rollback period: <strong style="color:$covColor">$covStr days</strong><span style="color:#888"> (target $RollbackTarget+ days)</span>$behindHtml$cadenceHtml</span>
    <span style="color:#888">Oldest: $($vs.s1_oldest) | Newest: $($vs.s1_newest)</span>
  </div>
"@ } else { "" }

        # --- Disk & shadow allocation bars ---
        $diskBarsHtml = ''
        $s1Hint = if ($ShowS1) { " <span style='color:#888;font-size:0.82em'>(S1 min 10%)</span>" } else { "" }

        if ($vs.disk_total_gb -gt 0) {
            # Row 1: disk bar showing shadow allocation footprint
            $allocPct     = $vs.alloc_pct
            $allocPctBar  = [math]::Min($allocPct, 100)
            $allocColor   = if ($vs.unbounded) { '#74b9ff' } elseif ($allocPct -lt 10) { '#e74c3c' } elseif ($allocPct -lt 15) { '#f39c12' } else { '#27ae60' }
            $diskBarFill  = if ($vs.unbounded) { '#1a3a5c' } elseif ($allocPct -lt 10) { '#5c1a1a' } elseif ($allocPct -lt 15) { '#5c3a1a' } else { '#1a3a1a' }
            $allocRightTd = if ($vs.unbounded) {
                "$(Format-GB $vs.disk_total_gb) &nbsp;&mdash;&nbsp; alloc: <span style='color:#74b9ff;font-weight:bold'>UNBOUNDED</span> <span style='color:#555'>(grows with demand)</span>"
            } else {
                "$(Format-GB $vs.disk_total_gb) &nbsp;&mdash;&nbsp; alloc: <span style='color:$allocColor'>$($allocPct.ToString('F1'))% ($(Format-GB $vs.max_gb) reserved)</span>$s1Hint"
            }

            # Row 2: shadow usage bar
            $usedPct    = $vs.pct_used
            $usedPctBar = [math]::Min($usedPct, 100)
            $usedColor  = if ($usedPct -ge 90) { '#e74c3c' } elseif ($usedPct -ge 75) { '#f39c12' } else { '#27ae60' }
            $usedRightTd = if ($vs.unbounded) {
                "$(Format-GB $vs.used_gb) used <span style='color:#555'>(no cap)</span>"
            } else {
                "$(Format-GB $vs.used_gb) of $(Format-GB $vs.max_gb) used ($($usedPct.ToString('F1'))%)"
            }

            $diskBarsHtml = @"
  <table style="width:100%;border-collapse:collapse;margin-top:8px;font-size:0.83em">
    <colgroup><col style="width:90px"><col><col style="width:auto"></colgroup>
    <tr>
      <td style="color:#666;padding:3px 8px 3px 0;white-space:nowrap;vertical-align:middle">Disk total</td>
      <td style="vertical-align:middle;min-width:120px">
        <div style="background:#2a2a2a;border-radius:3px;height:9px;overflow:hidden">
          <div style="width:$($allocPctBar)%;height:100%;background:$diskBarFill;border-right:2px solid $allocColor;box-sizing:border-box"></div>
        </div>
      </td>
      <td style="padding:0 0 0 10px;color:#aaa;white-space:nowrap;vertical-align:middle">$allocRightTd</td>
    </tr>
    <tr>
      <td style="color:#666;padding:3px 8px 3px 0;white-space:nowrap;vertical-align:middle">Shadow used</td>
      <td style="vertical-align:middle">
        <div style="background:#2a2a2a;border-radius:3px;height:9px;overflow:hidden">
          <div style="width:$($usedPctBar)%;height:100%;background:$usedColor;border-radius:3px"></div>
        </div>
      </td>
      <td style="padding:0 0 0 10px;color:#aaa;white-space:nowrap;vertical-align:middle">$usedRightTd</td>
    </tr>
  </table>
"@
        } else {
            # No disk size known - fallback to old single bar
            $capStr = if ($vs.unbounded) { "<span style='color:#74b9ff'>UNBOUNDED</span>" } else { "$(Format-GB $vs.max_gb)" }
            $diskBarsHtml = @"
  <div style="margin-top:6px;font-size:0.83em;color:#888">Shadow: $(Format-GB $vs.used_gb) / $capStr</div>
  $(Get-StorageBar $vs.pct_used)
"@
        }

        $html += @"
<div style="background:#1e1e1e;border:1px solid #444;border-radius:6px;padding:12px;margin-bottom:10px">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
    <strong style="font-size:1.05em">$d</strong>
    <span style="font-size:0.85em;color:#aaa">$($vs.count) total &nbsp;|&nbsp; $snapCounts</span>
  </div>
  $diskBarsHtml$s1RowHtml
</div>
"@
    }
    return $html
}

function Get-SoftwareDetectionSection {
    param([object]$Software)
    if (-not $Software -or $Software.Count -eq 0) { return '' }

    # Deduplicate: one row per mfg+product, merging via/name if both service and process matched
    $grouped = [ordered]@{}
    foreach ($sw in $Software) {
        $key = "$($sw.cat)|$($sw.mfg)|$($sw.product)"
        if (-not $grouped.Contains($key)) {
            $grouped[$key] = @{ cat=$sw.cat; mfg=$sw.mfg; product=$sw.product; via=[System.Collections.Generic.List[string]]::new(); names=[System.Collections.Generic.List[string]]::new(); state='' }
        }
        $grouped[$key].via.Add($sw.via)    | Out-Null
        $grouped[$key].names.Add($sw.name) | Out-Null
        if ($sw.state -and -not $grouped[$key].state) { $grouped[$key].state = $sw.state }
    }

    $catOrder = @('security','backup','rmm','remote')
    $catLabel = @{ security='Security'; backup='Backup'; rmm='RMM'; remote='Remote Access' }

    # Single table with category header rows so columns stay aligned across all groups
    $tbody = ''
    $first = $true
    foreach ($cat in $catOrder) {
        $entries = @($grouped.Values | Where-Object { $_.cat -eq $cat })
        if ($entries.Count -eq 0) { continue }
        $sepStyle = if ($first) { '' } else { 'border-top:2px solid #2a3a2a' }
        $first = $false
        $tbody += "<tr><td colspan='5' style='padding:8px 8px 2px;font-size:0.82em;font-weight:bold;color:#888;text-transform:uppercase;letter-spacing:0.05em;$sepStyle'>$($catLabel[$cat])</td></tr>"
        foreach ($entry in $entries) {
            $viaStr   = ($entry.via   | Select-Object -Unique) -join ' + '
            $nameStr  = ($entry.names | Select-Object -Unique) -join ', '
            $stateCol = if ($entry.state -match 'Running|Started') { '#6ecf6e' } else { '#888' }
            $stateTxt = if ($entry.state) { $entry.state } else { '-' }
            $tbody += "<tr style='border-bottom:1px solid #1e1e1e'>"
            $tbody += "<td style='padding:4px 8px;color:#aaa'>$([System.Web.HttpUtility]::HtmlEncode($entry.mfg))</td>"
            $tbody += "<td style='padding:4px 8px;color:#ccc'>$([System.Web.HttpUtility]::HtmlEncode($entry.product))</td>"
            $tbody += "<td style='padding:4px 8px;color:#777;font-size:0.88em'>$viaStr</td>"
            $tbody += "<td style='padding:4px 8px;color:#666;font-size:0.85em;font-family:Consolas,monospace'>$([System.Web.HttpUtility]::HtmlEncode($nameStr))</td>"
            $tbody += "<td style='padding:4px 8px;color:$stateCol;font-size:0.88em;white-space:nowrap'>$stateTxt</td>"
            $tbody += "</tr>"
        }
    }
    $innerHtml = "<table style='width:100%;border-collapse:collapse;font-size:0.85em'>"
    $innerHtml += "<colgroup><col style='width:15%'><col style='width:20%'><col style='width:8%'><col><col style='width:80px'></colgroup>"
    $innerHtml += "<thead><tr style='color:#555;border-bottom:1px solid #333'>"
    $innerHtml += "<th style='text-align:left;padding:4px 8px;font-weight:normal'>Manufacturer</th>"
    $innerHtml += "<th style='text-align:left;padding:4px 8px;font-weight:normal'>Product</th>"
    $innerHtml += "<th style='text-align:left;padding:4px 8px;font-weight:normal'>Via</th>"
    $innerHtml += "<th style='text-align:left;padding:4px 8px;font-weight:normal'>Service / Process</th>"
    $innerHtml += "<th style='text-align:left;padding:4px 8px;font-weight:normal'>State</th>"
    $innerHtml += "</tr></thead><tbody>$tbody</tbody></table>"

    $count = $grouped.Count
    return @"
<details style="margin-bottom:14px">
  <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">Software Detection ($count detected)</summary>
  $innerHtml
</details>
"@
}

function Get-RemediationBlock {
    param([System.Collections.Generic.List[hashtable]]$Remediations)
    if (-not $Remediations -or $Remediations.Count -eq 0) {
        return "<p style='color:#888'>No automated remediation available for this device.</p>"
    }
    $html = ""
    foreach ($r in $Remediations) {
        $color  = if ($r.sev -eq "CRITICAL") { "#e74c3c" } else { "#f39c12" }
        $cmdStr = $r.cmds -join "`n"
        $html  += @"
<div style="margin-bottom:12px">
  <div style="color:$color;font-weight:bold;margin-bottom:4px">[$($r.sev)] $($r.title)</div>
  <pre style="background:#111;border:1px solid #333;border-radius:4px;padding:10px;font-family:Consolas,monospace;font-size:0.85em;color:#a8ff78;overflow-x:auto;white-space:pre-wrap">$cmdStr</pre>
</div>
"@
    }
    return $html
}

function Get-DeviceSection {
    param([hashtable]$D, [int]$Idx, [string]$ConsoleBaseUrl, [string]$MaxBmVersion = '', [int]$RollbackTarget = 7)

    $bg  = $SEV_BG[$D.severity]  ?? "#1a1a1a"
    $bdr = $SEV_BORDER[$D.severity] ?? "#555"
    $t0col = if ($D.t0 -eq 2) { "#e74c3c" } else { "#f39c12" }

    $issuesHtml = ""
    $warnIssues = @($D.issues | Where-Object { $_ -notmatch '^\[i\]' })
    $infoIssues = @($D.issues | Where-Object { $_ -match '^\[i\]' })
    foreach ($iss in $warnIssues) {
        $isCrit = $iss -match 'CRITICAL|hung|severely|unreachable'
        $icon = if ($isCrit) { "&#128308;" } else { "&#128993;" }
        $issuesHtml += "<li style=`"margin-bottom:4px`">$icon $iss</li>"
    }
    foreach ($iss in $infoIssues) {
        $text = $iss -replace '^\[i\]', ''
        $issuesHtml += "<li style=`"margin-bottom:4px;color:#888`"><span style=`"color:#4a9eda;font-weight:bold;font-style:italic;margin-right:4px`">i</span>$text</li>"
    }
    if (-not $issuesHtml) {
        $issuesHtml = "<li style=`"color:#27ae60`">&#9989; No issues detected</li>"
    }

    $badge      = Get-SevBadge $D.severity
    $consoleUrl = "https://backup.management/#/backup/overview/view/$($D.account_id)(panel:device-properties/$($D.account_id)/summary)"
    $custHtml   = if ($D.customer_name) { "<span style=`"color:#aaa;font-size:0.75em;font-weight:normal`"> | $($D.customer_name)</span>" } else { "" }
    $backupHtml = if ($D.backup_in_progress) { "<span style=`"color:#4ecdc4;font-weight:bold`">IN PROGRESS</span>" } else { "<span style=`"color:#888`">Idle</span>" }
    $manualHtml = if ($D.manual_vss_actionable -and $D.manual_vss_running.Count -gt 0 -and -not $D.backup_in_progress) {
        "&nbsp;|&nbsp;<span style=`"color:#f39c12`">&#9888; Manual VSS: $($D.manual_vss_running -join ', ')</span>"
    } else { "" }

    $writerSummary = if ($D.unhealthy_writers.Count -gt 0) {
        "&#8212; <span style=`"color:#e74c3c`">$($D.unhealthy_writers.Count) UNHEALTHY</span>"
    } else { "&#8212; all Stable" }

    $vssCorrelation = Get-VssCorrelationSection `
        -Writers        $D.writers `
        -ServicesMap    ($D.services_map ?? @{}) `
        -BackupInProgress $D.backup_in_progress `
        -LastSessions   ($D.last_sessions ?? @{}) `
        -Volumes        $D.volumes `
        -SnapsByDrive   ($D.snaps_by_drive ?? @{}) `
        -S1Installed    ([bool]$D.s1_installed)
    $remBlock    = Get-RemediationBlock $D.remediations
    $sysEvtSection = Get-SystemEventsSection ($D.system_events ?? @()) ($D.device_tz ?? '')
    $sysEvtCount = ($D.system_events ?? @()).Count
    $swSection   = Get-SoftwareDetectionSection $D.detected_software

    # Dot badges
    # Determine whether to show any S1 badges:
    # - show if S1 snapshots exist (installed + working)
    # - show if S1 installed but no snapshots yet (writer or service detected)
    # - show if this device is missing S1 but sibling devices at this customer have it
    # - suppress entirely if S1 is simply not present and not expected
    $s1Present       = [bool]($D.volumes | Where-Object { $_.s1_count -gt 0 })
    $s1Installed     = $D.s1_installed -eq $true
    $s1Inconsistency = $D.s1_inconsistency -eq $true
    $showS1Badge     = $s1Present -or $s1Installed -or $s1Inconsistency

    $volsSection = Get-VolumesSection $D.volumes ([bool]$showS1Badge) ($D.volume_guid_map ?? @{})

    $s1Dot = ''; $s1Txt = ''; $s1Tip = ''
    if ($showS1Badge) {
        if ($s1Present) {
            $s1Dot = 'dok'; $s1Txt = 'S1 ✓'; $s1Tip = 'SentinelOne writer-based snapshots detected'
        } elseif ($s1Inconsistency) {
            $s1Dot = 'dct'; $s1Txt = 'S1 ✗'; $s1Tip = 'S1 not detected on this device - other devices at this customer have S1'
        } else {
            $s1Dot = 'dct'; $s1Txt = 'S1 ✗'; $s1Tip = 'S1 writer/service detected but no S1 snapshots found'
        }
    }
    # Rollback (worst volume) - only shown when S1 badge is shown
    $rbVols      = @($D.volumes | Where-Object { $_.s1_count -gt 1 })
    $minRollback = if ($rbVols.Count -gt 0) { ($rbVols | Measure-Object s1_coverage_days -Minimum).Minimum } else { $null }
    $maxRollback = if ($rbVols.Count -gt 0) { ($rbVols | Measure-Object s1_coverage_days -Maximum).Maximum } else { $null }
    $rbDot  = if ($null -eq $minRollback) { 'dgr' } elseif ($minRollback -lt $RollbackTarget) { 'dct' } else { 'dok' }
    $rbTxt  = if ($null -eq $minRollback) { 'Rollback n/a' } elseif ($maxRollback -gt $minRollback + 0.5) { "Rollback $([math]::Round($minRollback,0))d-$([math]::Round($maxRollback,0))d" } else { "Rollback $([math]::Round($minRollback,0))d" }
    $rbTip  = if ($null -eq $minRollback) { 'No S1 rollback data' } else { "S1 rollback window: $([math]::Round($minRollback,1))d min / $([math]::Round($maxRollback,1))d max (target $RollbackTarget+ days)" }
    # Shadow storage (worst volume)
    $maxShadow = if ($D.volumes.Count -gt 0) { ($D.volumes | Measure-Object pct_used -Maximum).Maximum } else { 0 }
    $shDot  = if ($maxShadow -ge 90) { 'dct' } elseif ($maxShadow -ge 80) { 'dwn' } else { 'dok' }
    $shTxt  = "Shadow $([math]::Round($maxShadow,0))%"
    $shTip  = "Highest shadow storage utilization: $([math]::Round($maxShadow,1))%"
    # Last backup age
    $nowU = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $lastBkpU = ($D.last_sessions.Values | ForEach-Object { $_.completed | Select-Object -First 1 } | Where-Object { $_ } | Sort-Object -Desc | Select-Object -First 1)
    $bkpAgeH = if ($lastBkpU) { [math]::Round(($nowU - $lastBkpU) / 3600.0, 1) } else { 999 }
    $bkpDot  = if ($bkpAgeH -lt 24) { 'dok' } elseif ($bkpAgeH -lt 48) { 'dwn' } else { 'dct' }
    $bkpTxt  = if ($bkpAgeH -ge 48) { "Bkp $([int]($bkpAgeH/24))d ago" } elseif ($bkpAgeH -ge 1) { "Bkp $([int]$bkpAgeH)h ago" } else { "Bkp $([int]($bkpAgeH*60))m ago" }
    $bkpTip  = if ($lastBkpU) { "Last completed backup: $([math]::Round($bkpAgeH,1)) hours ago" } else { 'No completed backup found' }
    # S1 snap cadence - only shown when S1 badge is shown
    $s1CadH = 0
    foreach ($drv in $D.snaps_by_drive.Keys) {
        $s1Ts = @($D.snaps_by_drive[$drv] | Where-Object { $_.type -eq 's1' } | Sort-Object { $_.ts } -Desc)
        if ($s1Ts.Count -ge 2) {
            $g2 = @(); for ($gi=0;$gi -lt [math]::Min($s1Ts.Count-1,10);$gi++){$g=$( ($s1Ts[$gi].ts-$s1Ts[$gi+1].ts).TotalHours);if($g-gt0){$g2+=$g}}
            if ($g2.Count -gt 0) { $med2=(@($g2|Sort-Object))[[int]($g2.Count/2)];if($s1CadH-eq0-or$med2-lt$s1CadH){$s1CadH=$med2} }
        }
    }
    $scDot  = if ($s1CadH -eq 0) { 'dgr' } elseif ($s1CadH -le 1.5) { 'dnf' } elseif ($s1CadH -le 6) { 'dnf' } else { 'dwn' }
    $scTxt  = if ($s1CadH -gt 0) { "S1 ~$([int][math]::Round($s1CadH))h" } else { 'S1 cad n/a' }
    $scTip  = if ($s1CadH -gt 0) { "S1 snapshot cadence: ~$([math]::Round($s1CadH,1))h" } else { 'S1 cadence unknown (insufficient data)' }
    # Writers
    $wrBad  = $D.unhealthy_writers.Count -gt 0
    $wrDot  = if ($wrBad) { 'dct' } else { 'dok' }
    $wrTxt  = if ($wrBad) { "Writers $($D.unhealthy_writers.Count)✗" } else { 'Writers ✓' }
    $wrTip  = if ($wrBad) { "$($D.unhealthy_writers.Count) unhealthy VSS writer(s)" } else { 'All VSS writers stable' }

    $s1BadgeHtml  = if ($showS1Badge) { "<span class='di' title='$s1Tip'><span class='dt $s1Dot'></span>$s1Txt</span><span style='color:#2a2a2a'>|</span>" } else { '' }
    $rbBadgeHtml  = if ($showS1Badge) { "<span class='di' title='$rbTip'><span class='dt $rbDot'></span>$rbTxt</span><span style='color:#2a2a2a'>|</span>" } else { '' }
    $scBadgeHtml  = if ($showS1Badge) { "<span class='di' title='$scTip'><span class='dt $scDot'></span>$scTxt</span><span style='color:#2a2a2a'>|</span>" } else { '' }
    $badgesHtml = "<div class='bdg'>$s1BadgeHtml$rbBadgeHtml<span class='di' title='$shTip'><span class='dt $shDot'></span>$shTxt</span><span style='color:#2a2a2a'>|</span><span class='di' title='$bkpTip'><span class='dt $bkpDot'></span>$bkpTxt</span><span style='color:#2a2a2a'>|</span>$scBadgeHtml<span class='di' title='$wrTip'><span class='dt $wrDot'></span>$wrTxt</span></div>"

    $provHtml = ""
    foreach ($p in $D.providers) {
        $pName = ($p.Name ?? $p.'Provider Name' ?? '?')
        $pVer  = ($p.Version ?? '?')
        $pType = ($p.Type ?? '?')
        $prefix = if ($p -in $D.third_party_providers) { "&#9888; " } else { "" }
        $provHtml += "<li>$prefix$pName (v$pVer, $pType)</li>"
    }

    $issueCount  = $warnIssues.Count
    $writerCount = $D.writers.Count
    $volCount    = $D.volumes.Count
    $provCount   = $D.providers.Count
    $remCount    = $D.remediations.Count
    $tpCount     = $D.third_party_providers.Count
    $tpHtml      = if ($tpCount -gt 0) { "&#8212; <span style=`"color:#f39c12`">$tpCount THIRD PARTY</span>" } else { "" }
    $hwMake      = (($D.hw_manufacturer, $D.hw_model) | Where-Object { $_ }) -join ' '
    $hwInfo      = (@($hwMake, $(if($D.hw_cpu){"CPU: $($D.hw_cpu)"}), $(if($D.hw_ram){"RAM: $($D.hw_ram)"})) | Where-Object { $_ }) -join ' | '
    $retInfo     = (@($D.product_name, $D.retention_units, $(if($D.profile_name){"Profile: $($D.profile_name)"})) | Where-Object { $_ }) -join ' | '
    $bmHtml      = ''
    if ($D.backup_manager_version) {
        $bmColor = '#666'
        if ($MaxBmVersion -and $D.backup_manager_version -ne $MaxBmVersion) {
            try {
                if ([System.Version]$D.backup_manager_version -lt [System.Version]$MaxBmVersion) { $bmColor = '#e67e22' }
            } catch {}
        }
        $bmHtml = " &nbsp;|&nbsp; BM: <span style='color:$bmColor' title='Newest: $MaxBmVersion'>$($D.backup_manager_version)</span>"
    }

    $devTypeAttr = switch ([int]($D.device_type ?? 0)) { 1 { 'workstation' } 2 { 'server' } default { 'other' } }

    # Data sources inline summary with per-source tooltip of recent sessions
    $dataSourcesHtml = ''
    if ($D.last_sessions -and $D.last_sessions.Count -gt 0) {
        $nowU3 = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $srcParts = [System.Collections.Generic.List[string]]::new()
        foreach ($plugId in ($D.last_sessions.Keys | Sort-Object)) {
            $pd   = $D.last_sessions[$plugId]
            $name = $PLUGIN_NAMES[$plugId] ?? "Plugin $plugId"
            $sf   = @($pd.sessions_full | Sort-Object { $_.end } -Descending | Select-Object -First 6)
            if ($sf.Count -eq 0) { continue }

            # Most recent session determines color
            $latest = $sf[0]
            $latestOk = $latest.status -eq 'Completed' -or ($latest.cleaned -and $latest.status -ne 'Failed')
            $nameColor = if ($latestOk) { '#888' } else { '#c0392b' }

            # Build tooltip: last N sessions
            $tipLines = @("$name — last $([math]::Min($sf.Count,5)) sessions:")
            foreach ($s in ($sf | Select-Object -First 5)) {
                $ageH   = [math]::Round(($nowU3 - $s.end) / 3600.0, 0)
                $ageStr = if ($ageH -lt 24) { "${ageH}h ago" } elseif ($ageH -lt 48) { "yesterday" } else { "$([int]($ageH/24))d ago" }
                $st     = if ($s.cleaned) { 'Cleaned' } else { $s.status }
                $durMin = if ($s.start -gt 0) { [int](($s.end - $s.start) / 60) } else { 0 }
                $durStr = if ($durMin -gt 60) { "$([int]($durMin/60))h$($durMin%60)m" } elseif ($durMin -gt 0) { "${durMin}m" } else { '' }
                $tipLines += "  $ageStr — $st$(if($durStr){" ($durStr)"})"
            }
            $tip = [System.Web.HttpUtility]::HtmlAttributeEncode($tipLines -join "`n")

            $srcParts.Add("<span title='$tip' style='cursor:help;color:$nameColor;white-space:nowrap'>$([System.Web.HttpUtility]::HtmlEncode($name))</span>")
        }
        if ($srcParts.Count -gt 0) {
            $dataSourcesHtml = "<div style='color:#666;font-size:0.80em;margin-top:2px'>Sources: $($srcParts -join "<span style='color:#333'> &nbsp;|&nbsp; </span>")</div>"
        }
    }

    return @"
<div id="device-$Idx" data-severity="$($D.severity)" data-backup-status="$($D.backup_status)" data-device-type="$devTypeAttr" data-search="$(($D.machine_name ?? '').ToLower()) $(($D.device_name ?? '').ToLower()) $(($D.customer_name ?? '').ToLower())" data-accountid="$($D.account_id)" data-machine="$($D.machine_name -replace '"','&quot;' -replace '&','&amp;')" data-customer="$($D.customer_name -replace '"','&quot;' -replace '&','&amp;')" data-console="$consoleUrl" style="background:$bg;border:1px solid $bdr;border-radius:8px;padding:20px;margin-bottom:24px">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px">
    <div>
      <h2 style="margin:0;font-size:1.3em">
        $($D.machine_name)
        <span style="color:#888;font-size:0.75em;font-weight:normal">($($D.device_name))</span>$custHtml
        <a href="$consoleUrl" target="_blank" style="color:#5dade2;font-size:0.65em;font-weight:normal;text-decoration:none">(Open in Cove)</a>
      </h2>
      <div style="color:#aaa;font-size:0.85em;margin-top:2px">$($D.os) &nbsp;|&nbsp; AccountId: $($D.account_id) &nbsp;|&nbsp; Backup: <span style="color:$t0col;font-weight:bold">$($D.t0_label)</span></div>
            <div style="color:#888;font-size:0.82em;margin-top:2px">VSS: $($D.vss_service) &nbsp;|&nbsp; $backupHtml$manualHtml$bmHtml$(if($D.device_now){" &nbsp;|&nbsp; Device time: <span style='color:#666'>$($D.device_now.ToString('yyyy-MM-dd HH:mm')) $($D.device_tz)</span>"})</div>
            <div style="color:#888;font-size:0.82em;margin-top:2px">IP: $(if($D.ip_address){$D.ip_address}else{'N/A'})$(if($D.last_timestamp){" &nbsp;|&nbsp; Last seen: <span style='color:#$(if($D.last_timestamp -gt ((Get-Date).AddDays(-7)).ToUniversalTime().ToFileTime()){' 6ecf6e'}else{' e67e22'})'> $([System.DateTimeOffset]::FromUnixTimeSeconds($D.last_timestamp).LocalDateTime.ToString('yyyy-MM-dd HH:mm'))</span>"})$(if($hwInfo){" &nbsp;|&nbsp; $hwInfo"})$(if($retInfo){" &nbsp;|&nbsp; $retInfo"})</div>
            $dataSourcesHtml
    </div>
    <div style="display:flex;align-items:center">$badgesHtml$badge</div>
  </div>

  <details open style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">$(if($issueCount -gt 0){"Issues ($issueCount found)$(if($infoIssues.Count -gt 0){" <span style='color:#555;font-weight:normal;font-size:0.88em'>+ $($infoIssues.Count) note$(if($infoIssues.Count -ne 1){"s"})</span>"})"}else{if($infoIssues.Count -gt 0){"<span style='color:#27ae60'>No warnings</span> <span style='color:#555;font-weight:normal;font-size:0.88em'>— $($infoIssues.Count) note$(if($infoIssues.Count -ne 1){"s"})</span>"}else{"Issues (0 found)"}})</summary>
    <ul style="margin:6px 0;padding-left:20px;color:#ddd;font-size:0.9em">$issuesHtml</ul>
  </details>

  <details open style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">Snapshot &amp; Backup Timeline</summary>
    <div style="padding:6px 0">
      $(Get-TimelineHtml $D)
    </div>
  </details>

  <details style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">Shadow Copy Volumes ($volCount volumes)</summary>
    $volsSection
  </details>

  <details open style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">VSS Writers &amp; Dependency Analysis ($writerCount writers $writerSummary)</summary>
    <div style="padding-top:8px">$vssCorrelation</div>
  </details>

  <details style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">VSS Providers ($provCount registered $tpHtml)</summary>
    <ul style="margin:4px 0;padding-left:20px;font-size:0.88em;color:#ddd">$provHtml</ul>
  </details>

  $swSection

  <details style="margin-bottom:14px">
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">System Events (errors &amp; warnings) - $sysEvtCount event$(if ($sysEvtCount -ne 1) { 's' })</summary>
    <div style="padding-top:6px">$sysEvtSection</div>
  </details>

  <details$(if ($remCount -gt 0) { ' open' } else { '' })>
    <summary style="cursor:pointer;font-weight:bold;color:#ddd;margin-bottom:8px">Remediation Scripts ($remCount actions)</summary>
    $remBlock
  </details>
  <div style="margin-top:16px;border-top:1px solid #333;padding-top:12px">
    <div style="font-size:0.82em;color:#888;margin-bottom:4px">Notes <span id="nc-$($D.account_id)" style="color:#555">(0/5000)</span></div>
    <textarea id="notes-$($D.account_id)" maxlength="5000" rows="2" placeholder="Add notes..." oninput="updateNoteCount('$($D.account_id)',this.value)" style="width:100%;background:#111;border:1px solid #333;color:#ccc;padding:6px 8px;border-radius:4px;font-size:0.85em;resize:vertical;box-sizing:border-box;font-family:inherit"></textarea>
    <button onclick="saveNotes('$($D.account_id)')" style="margin-top:4px;padding:4px 12px;background:#1e2e1e;border:1px solid #3a5a3a;color:#8ecf8e;border-radius:4px;cursor:pointer;font-size:0.82em">Save notes</button>
  </div>
</div>
"@
}

function Build-HtmlReport {
    param([array]$Devices, [string]$PartnerLabel, [int]$PartnerId, [string]$ReportDate,
          [string]$ScriptName = "", [string]$RunDuration = "", [string]$EmailHref = "", [string]$ReportFolderHref = "", [string]$ReportFilePath = "")

    $red     = @($Devices | Where-Object { $_.severity -eq "RED" })
    $yellow  = @($Devices | Where-Object { $_.severity -eq "YELLOW" })
    $green   = @($Devices | Where-Object { $_.severity -eq "GREEN" })
    $offline = @($Devices | Where-Object { $_.severity -eq "OFFLINE" })

    $consoleBase = "https://backup.management/#/backup/overview/view/$PartnerId"

    # Sort by customer name, then machine name
    $Devices = @($Devices | Sort-Object { $_.customer_name }, { $_.machine_name })

    $navItems = ""
    $lastCust = ""
    for ($i = 0; $i -lt $Devices.Count; $i++) {
        $d = $Devices[$i]
        $c = $SEV_COLOR[$d.severity] ?? "#fff"
        $icon = switch ($d.severity) {
            "RED"     { "&#128308;" } "YELLOW" { "&#128993;" }
            "OFFLINE" { "&#9899;" }   default  { "&#128994;" }
        }
        $cust = if ($d.customer_name) { $d.customer_name } else { "(Unknown)" }
        $custKey = [System.Web.HttpUtility]::HtmlAttributeEncode($cust)
        if ($cust -ne $lastCust) {
            $navItems += "<div data-nav-group=`"$custKey`" style=`"padding:10px 10px 6px 10px;color:#888;font-size:0.72em;font-weight:bold;text-transform:uppercase;letter-spacing:0.05em;background:#161616;border-bottom:1px solid #2a2a2a;margin-top:0`">$cust</div>"
            $lastCust = $cust
        }
        $navDevType = switch ([int]($d.device_type ?? 0)) { 1 { 'workstation' } 2 { 'server' } default { 'other' } }
        $navItems += "<a href=`"#device-$i`" data-nav-sev=`"$($d.severity)`" data-nav-backup-status=`"$($d.backup_status)`" data-nav-type=`"$navDevType`" data-nav-group=`"$custKey`" data-nav-search=`"$(($d.machine_name ?? '').ToLower()) $(($d.device_name ?? '').ToLower()) $(($d.customer_name ?? '').ToLower())`" style=`"display:flex;flex-direction:column;gap:2px;padding:6px 10px 6px 18px;color:$c;text-decoration:none;border-bottom:1px solid #222;font-size:0.88em`"><span style=`"display:inline`">$icon $($d.machine_name)</span><span style=`"color:#666;font-size:0.8em`">$($d.t0_label)</span></a>"
    }

    $maxBmParsed = $null
    foreach ($d in $Devices) {
        $v = $d.backup_manager_version
        if (-not $v) { continue }
        try { $p = [System.Version]$v; if ($null -eq $maxBmParsed -or $p -gt $maxBmParsed) { $maxBmParsed = $p } } catch {}
    }
    $maxBmStr = if ($maxBmParsed) { $maxBmParsed.ToString() } else { '' }

    $devSections = ""
    for ($i = 0; $i -lt $Devices.Count; $i++) {
        $devSections += Get-DeviceSection $Devices[$i] $i $consoleBase $maxBmStr $RollbackTarget
    }

    # Top issue patterns (normalize titles to combine similar issues)
    $topIssues = @{}
    $issueTracking = @{}  # Track original titles for stats
    foreach ($d in $Devices) {
        foreach ($r in $d.remediations) {
            $t = $r.title
            $normalized = $t
            
            # Normalize: remove specific counts in parentheses for aggregation
            # e.g., "VSS writer failure (4 writer(s))" -> "VSS writer failure(s)"
            $normalized = $normalized -replace '\s*\(\d+\s+writer\(s\)\)', ''
            if ($normalized -match 'writer failure') { $normalized = $normalized -replace '$', '(s)' }
            
            # Normalize shadow storage patterns across drives
            # e.g., "Shadow storage near-full on C:\: 96.6% ..." -> "Shadow storage near-full"
            $normalized = $normalized -replace '\s+on\s+[A-Za-z]:\\.*', ''
            
            # Normalize snapshot lifecycle descriptions
            # Keep the base message, remove the detailed state descriptions
            $normalized = $normalized -replace '\s+\(snapshot\s+lifecycle.*\)', ''
            
            $topIssues[$normalized] = ($topIssues[$normalized] ?? 0) + 1
            # Track original for stats
            if (-not $issueTracking[$normalized]) { $issueTracking[$normalized] = @() }
            $issueTracking[$normalized] += $t
        }
    }
    $patternHtml = if ($topIssues.Count -gt 0) {
        ($topIssues.GetEnumerator() | Sort-Object Value -Descending |
         ForEach-Object { "<div style=`"margin-bottom:4px`">&#128993; $($_.Key) ($($_.Value) device$(if($_.Value -gt 1){'s'}))</div>" }) -join ""
    } else { "<div style='color:#888'>No common patterns identified</div>" }

    # Fleet stats
    $failedCount   = @($Devices | Where-Object { $_.t0 -eq 2  }).Count
    $errorsCount   = @($Devices | Where-Object { $_.t0 -eq 8  }).Count
    $unhWriters    = @($Devices | Where-Object { $_.unhealthy_writers.Count -gt 0 }).Count
    $lowCoverage   = @($Devices | Where-Object {
        @($_.volumes | Where-Object { $_.s1_count -gt 1 -and $_.s1_coverage_days -lt 7 }).Count -gt 0
    }).Count
    $behindSched   = @($Devices | Where-Object {
        @($_.volumes | Where-Object { $_.behind_schedule }).Count -gt 0
    }).Count
    $s1MissCount   = @($Devices | Where-Object { $_.s1_inconsistency -eq $true }).Count

    # ===== Fleet Software & Tools Analysis =====
    $osStats      = @{}
    $rmmsDetected = @{}  # Track devices (not instances): { product => [set of device IDs] }
    $securityVendors = @{}
    $backupTools   = @{}
    
    foreach ($d in $Devices) {
        # Aggregate OS
        if ($d.os) {
            $osKey = $d.os -replace '.*?((Windows|Server|CentOS|Ubuntu|Debian|RHEL|Fedora|Alma|Rocky).*)', '$1'
            $osKey = $osKey -replace ',.*', ''
            $osStats[$osKey] = ($osStats[$osKey] ?? 0) + 1
        }
        
        # Aggregate software by category - track UNIQUE devices per tool
        if ($d.detected_software -and $d.detected_software.Count -gt 0) {
            # Use a set to track which products we've seen on this device
            $productsSeenOnDevice = @{}
            foreach ($sw in $d.detected_software) {
                $cat = $sw.cat ?? 'other'
                $prod = $sw.product ?? $sw.mfg
                
                # Only count once per device per product
                if ($productsSeenOnDevice.ContainsKey($prod)) { continue }
                $productsSeenOnDevice[$prod] = $true
                
                if ($cat -eq 'rmm') {
                    if (-not $rmmsDetected.ContainsKey($prod)) {
                        $rmmsDetected[$prod] = @{}
                    }
                    $rmmsDetected[$prod][$d.account_id] = $true
                } elseif ($cat -eq 'security') {
                    if (-not $securityVendors.ContainsKey($prod)) {
                        $securityVendors[$prod] = @{}
                    }
                    $securityVendors[$prod][$d.account_id] = $true
                } elseif ($cat -eq 'backup') {
                    if (-not $backupTools.ContainsKey($prod)) {
                        $backupTools[$prod] = @{}
                    }
                    $backupTools[$prod][$d.account_id] = $true
                }
            }
        }
    }
    
    # Convert device sets to counts
    $rmmsCount = @{}; $rmmsDetected.GetEnumerator() | ForEach-Object { $rmmsCount[$_.Key] = $_.Value.Count }
    $securityCount = @{}; $securityVendors.GetEnumerator() | ForEach-Object { $securityCount[$_.Key] = $_.Value.Count }
    $backupCount = @{}; $backupTools.GetEnumerator() | ForEach-Object { $backupCount[$_.Key] = $_.Value.Count }
    
    # Build fleet software HTML
    $fleetSoftwareHtml = ""
    if ($osStats.Count -gt 0) {
        # Categorize OSes into Windows, Linux, Mac
        $windows = @{}; $linux = @{}; $mac = @{}
        foreach ($os in $osStats.GetEnumerator()) {
            if ($os.Key -match 'Windows|Server') {
                $windows[$os.Key] = $os.Value
            } elseif ($os.Key -match 'Ubuntu|Debian|CentOS|RHEL|Fedora|Alma|Rocky|Linux') {
                $linux[$os.Key] = $os.Value
            } elseif ($os.Key -match 'macOS|Darwin|Mac') {
                $mac[$os.Key] = $os.Value
            }
        }
        
        $fleetSoftwareHtml += "<div style='margin-bottom:12px'><strong>Operating Systems:</strong><div style='font-size:0.85em;color:#ddd;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px'>"
        
        # Windows column
        $fleetSoftwareHtml += "<div><strong style='color:#4ecdc4'>Windows</strong><div style='margin-top:4px'>"
        if ($windows.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $windows.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                $pct = [math]::Round(($_.Value / $Devices.Count) * 100, 1)
                $fleetSoftwareHtml += "<div style='margin-bottom:2px'>$($_.Key)<br/><span style='color:#888;font-size:0.9em'> $($_.Value) ($pct%)</span></div>"
            }
        }
        $fleetSoftwareHtml += "</div></div>"
        
        # Linux column
        $fleetSoftwareHtml += "<div><strong style='color:#f0a030'>Linux</strong><div style='margin-top:4px'>"
        if ($linux.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $linux.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                $pct = [math]::Round(($_.Value / $Devices.Count) * 100, 1)
                $fleetSoftwareHtml += "<div style='margin-bottom:2px'>$($_.Key)<br/><span style='color:#888;font-size:0.9em'> $($_.Value) ($pct%)</span></div>"
            }
        }
        $fleetSoftwareHtml += "</div></div>"
        
        # Mac column
        $fleetSoftwareHtml += "<div><strong style='color:#b8ff79'>macOS</strong><div style='margin-top:4px'>"
        if ($mac.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $mac.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                $pct = [math]::Round(($_.Value / $Devices.Count) * 100, 1)
                $fleetSoftwareHtml += "<div style='margin-bottom:2px'>$($_.Key)<br/><span style='color:#888;font-size:0.9em'> $($_.Value) ($pct%)</span></div>"
            }
        }
        $fleetSoftwareHtml += "</div></div></div></div>"
    }
    
    if ($rmmsCount.Count -gt 0 -or $backupCount.Count -gt 0 -or $securityCount.Count -gt 0) {
        $fleetSoftwareHtml += "<div style='margin-bottom:12px'><strong>Tools:</strong><div style='font-size:0.85em;color:#ddd;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px'>"
        $fleetSoftwareHtml += "<div><strong style='color:#b8ff79'>RMM</strong><div>"
        if ($rmmsCount.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $rmmsCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                $fleetSoftwareHtml += "<div>$($_.Key) &#8226; $($_.Value)</div>"
            }
        }
        $fleetSoftwareHtml += "</div></div><div><strong style='color:#ff9c42'>Backup</strong><div>"
        if ($backupCount.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $backupCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                $fleetSoftwareHtml += "<div>$($_.Key) &#8226; $($_.Value)</div>"
            }
        }
        $fleetSoftwareHtml += "</div></div><div><strong style='color:#4ecdc4'>Security</strong><div>"
        if ($securityCount.Count -eq 0) {
            $fleetSoftwareHtml += "<div style='color:#888'>None</div>"
        } else {
            $securityCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object {
                $fleetSoftwareHtml += "<div>$($_.Key) &#8226; $($_.Value)</div>"
            }
            if ($securityCount.Count -gt 5) {
                $fleetSoftwareHtml += "<div style='color:#888'>+$($securityCount.Count - 5) more</div>"
            }
        }
        $fleetSoftwareHtml += "</div></div></div></div>"
    }

    $servers      = @($Devices | Where-Object { $_.device_type -eq 2 })
    $workstations = @($Devices | Where-Object { $_.device_type -eq 1 })
    $srvTotal  = $servers.Count
    $wsTotal   = $workstations.Count

    $srvRed    = @($servers      | Where-Object { $_.severity -eq "RED" }).Count
    $srvYellow = @($servers      | Where-Object { $_.severity -eq "YELLOW" }).Count
    $srvGreen  = @($servers      | Where-Object { $_.severity -eq "GREEN" }).Count
    $srvOffline= @($servers      | Where-Object { $_.severity -eq "OFFLINE" }).Count
    $wsRed     = @($workstations | Where-Object { $_.severity -eq "RED" }).Count
    $wsYellow  = @($workstations | Where-Object { $_.severity -eq "YELLOW" }).Count
    $wsGreen   = @($workstations | Where-Object { $_.severity -eq "GREEN" }).Count
    $wsOffline = @($workstations | Where-Object { $_.severity -eq "OFFLINE" }).Count

    $genLocal = [DateTimeOffset]::Now.ToString("yyyy-MM-dd HH:mm zzz")
    $genUtc   = [DateTime]::UtcNow.ToString("HH:mm") + " UTC"
    $genTime  = "$genLocal ($genUtc)"
    $totalCount  = $Devices.Count
    $redCount    = $red.Count; $yellowCount = $yellow.Count
    $greenCount  = $green.Count; $offlineCount = $offline.Count

    $emailAttr = ($EmailHref ?? '') -replace '&', '&amp;' -replace '"', '&quot;'
    $reportPathAttr = ($ReportFilePath ?? '') -replace '&', '&amp;' -replace '"', '&quot;'

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>VSS Health Report - $PartnerLabel - $ReportDate</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Segoe UI, Arial, sans-serif; background: #121212; color: #e0e0e0; display: flex; min-height: 100vh; }
  #sidebar { width: 220px; min-width: 220px; background: #1a1a1a; border-right: 1px solid #333; position: sticky; top: 0; height: 100vh; overflow-y: auto; }
  #sidebar-header { padding: 14px; background: #111; border-bottom: 1px solid #333; }
  #main { flex: 1; padding: 24px; overflow-y: auto; max-width: 1800px; }
  h1 { font-size: 1.6em; margin-bottom: 4px; }
  h2 { font-size: 1.2em; }
  details summary { outline: none; }
  details summary::-webkit-details-marker { color: #888; }
  a { color: #4ecdc4; }
  pre { tab-size: 4; }
    #toast { position: fixed; right: 18px; top: 18px; background:#1f3a1f; color:#d7ffd7; border:1px solid #2f6b2f; border-radius:8px; padding:10px 12px; font-size:0.85em; opacity:0; pointer-events:none; transition:opacity .18s ease; z-index:9999; }
    #toast.show { opacity:1; }
  .tl-vp  { width:100%; overflow-x:scroll; scrollbar-width:none; -ms-overflow-style:none; }
  .tl-vp::-webkit-scrollbar { display:none; }
  .tl-btn { background:rgba(20,20,20,0.85); border:1px solid #3a3a3a; color:#666; border-radius:3px; cursor:pointer; padding:0; line-height:1; transition:color .12s,background .12s; display:flex; align-items:center; justify-content:center; }
  .tl-btn:hover  { color:#ccc; background:rgba(55,55,55,0.95); border-color:#555; }
  .tl-btn:disabled { opacity:0.18; cursor:default; }
  .tl-lbtn { position:absolute; left:0; top:2px; width:20px; height:18px; z-index:10; font-size:0.7em; }
  .tl-rbtn { position:absolute; right:0; top:2px; width:20px; height:18px; z-index:10; font-size:0.7em; }
  .tl-outer { background:#161616; border-radius:4px; padding:10px; font-family:Consolas,monospace; display:flex; align-items:stretch; }
  .tl-lc { width:160px; min-width:160px; font-size:0.74em; text-align:right; padding-right:8px; flex-shrink:0; }
  .tl-vc { flex:1; min-width:0; position:relative; }
  .tl-vi { position:relative; }
  .tl-ax-lc { height:38px; display:flex; align-items:flex-end; justify-content:flex-end; padding-right:8px; color:#444; font-size:0.70em; margin-bottom:2px; }
  .tl-ax-row { flex-shrink:0; position:relative; height:38px; margin-bottom:2px; }
  .tl-lrow { height:34px; display:flex; flex-direction:column; justify-content:center; align-items:flex-end; padding-right:8px; margin-bottom:3px; }
  .tl-brow { flex-shrink:0; position:relative; height:40px; overflow:hidden; margin-bottom:3px; }
  .tl-brow-b { flex-shrink:0; position:relative; height:34px; margin-bottom:3px; }
  .tl-now { position:absolute; right:0; top:0; bottom:0; width:1px; background:rgba(255,255,255,0.08); pointer-events:none; }
  .tl-gl { position:absolute; top:0; width:1px; height:100%; background:rgba(255,255,255,0.055); pointer-events:none; }
  .tl-pbg { position:absolute; left:0; top:0; height:100%; background:repeating-linear-gradient(90deg,#1e1e1e 0px,#1e1e1e 3px,#222 3px,#222 6px); }
  .tl-lbg { position:absolute; top:0; right:0; height:100%; background:rgba(78,205,196,0.04); }
  .tl-cm  { position:absolute; top:0; width:1px; height:100%; background:#4ecdc466; }
  .tl-ax-e1 { position:absolute; left:0; top:0; width:1px; height:6px; background:#3a3a3a; }
  .tl-ax-ds { position:absolute; left:0; top:8px; font-size:10px; color:#555; text-align:left; }
  .tl-ax-hl { position:absolute; top:0; transform:translateX(-50%); font-size:9px; color:#555; white-space:nowrap; }
  .tl-ax-t  { position:absolute; top:0; width:1px; height:10px; background:#4a4a4a; }
  .tl-ax-dl { position:absolute; top:12px; transform:translateX(3px); text-align:left; font-size:10px; color:#999; line-height:1.3; }
  .tl-ax-er { position:absolute; right:0; top:0; width:1px; height:6px; background:#555; }
  .tl-ax-nw { position:absolute; right:0; top:8px; font-size:10px; color:#666; transform:translateX(-50%); white-space:nowrap; }
  .tl-s-s1  { position:absolute; top:0;   width:12px; height:13px; background:linear-gradient(to right,transparent 5px,#4ecdc4 5px,#4ecdc4 7px,transparent 7px); transform:translateX(-6px); cursor:default; }
  .tl-s-nat { position:absolute; top:10px; width:12px; height:14px; background:linear-gradient(to right,transparent 5px,#fd79a8 5px,#fd79a8 7px,transparent 7px); transform:translateX(-6px); cursor:default; }
  .tl-s-wri { position:absolute; top:10px; width:12px; height:14px; background:linear-gradient(to right,transparent 5px,#a29bfe 5px,#a29bfe 7px,transparent 7px); transform:translateX(-6px); cursor:default; }
  .tl-s-vsc { position:absolute; top:10px; width:12px; height:14px; background:linear-gradient(to right,transparent 5px,#7fb3d3 5px,#7fb3d3 7px,transparent 7px); transform:translateX(-6px); cursor:default; }
  .tl-s-oth { position:absolute; top:21px; width:12px; height:13px; background:linear-gradient(to right,transparent 5px,#787878 5px,#787878 7px,transparent 7px); transform:translateX(-6px); cursor:default; }
  .tl-s-cov { position:absolute; top:0; width:4px; height:100%; background:#52be80; opacity:0.9; transform:translateX(-2px); cursor:default; box-shadow:0 0 6px 1px #52be8088; border-radius:1px; }
  .tl-s-unk { position:absolute; top:0; width:4px; height:100%; background:#fdcb6e; opacity:0.9; transform:translateX(-2px); cursor:default; box-shadow:0 0 6px 1px #fdcb6e88; border-radius:1px; }
  .tl-bk-ok  { position:absolute; top:3px; height:14px; background:#27ae60; border-radius:1px; opacity:0.9; cursor:default; }
  .tl-bk-err { position:absolute; top:3px; height:14px; background:#e67e22; border-radius:1px; opacity:0.9; cursor:default; }
  .tl-bk-skp { position:absolute; top:3px; height:14px; background:#888888; border-radius:1px; opacity:0.9; cursor:default; }
  .tl-bk-fai { position:absolute; top:3px; height:14px; background:#e74c3c; border-radius:1px; opacity:0.9; cursor:default; }
  .tl-bk-cl::after  { content:''; position:absolute; inset:0; border-radius:1px; background-image:repeating-linear-gradient(45deg,rgba(0,0,0,0) 0px,rgba(0,0,0,0) 3px,rgba(0,0,0,0.38) 3px,rgba(0,0,0,0.38) 5px); pointer-events:none; }
  .tl-bk-ac::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; border-radius:1px 1px 0 0; background:rgba(255,255,255,0.65); pointer-events:none; }
  #tl-tip { display:none; position:fixed; background:#1e1e1e; border:1px solid #3a3a3a; border-radius:6px; padding:10px 12px; font-size:0.78em; line-height:1.65; color:#ccc; z-index:9999; min-width:210px; pointer-events:none; box-shadow:0 6px 20px rgba(0,0,0,0.6); }
  .jt-head { color:#4ecdc4; font-weight:bold; margin-bottom:5px; border-bottom:1px solid #2a2a2a; padding-bottom:4px; }
  .jt-row  { display:flex; justify-content:space-between; gap:12px; }
  .jt-lbl  { color:#666; white-space:nowrap; }
  .jt-val  { color:#ddd; text-align:right; }
  .jt-ok   { color:#27ae60; font-weight:bold; }
  .jt-err  { color:#e67e22; font-weight:bold; }
  .jt-fai  { color:#e74c3c; font-weight:bold; }
  .jt-skp  { color:#888;    font-weight:bold; }
  .jt-flag    { display:inline-block; padding:1px 6px; border-radius:3px; font-size:0.88em; margin-right:3px; background:rgba(255,255,255,0.07); color:#aaa; }
  .jt-sep     { border:none; border-top:1px solid #2a2a2a; margin:5px 0; }
  .jt-sec-lbl { font-size:0.72em; color:#3a3a3a; text-transform:uppercase; letter-spacing:0.06em; margin:7px 0 4px; border-top:1px solid #1e1e1e; padding-top:5px; }
  .jt-sub     { display:flex; justify-content:space-between; gap:12px; padding-left:12px; margin-bottom:2px; }
  .jt-slbl    { color:#3e3e3e; font-size:0.85em; white-space:nowrap; }
  .jt-sval    { color:#666; font-size:0.85em; text-align:right; }
  .tl-mss1 { position:absolute; top:3px;  height:8px; border-left:1px dashed rgba(243,156,18,0.80); transform:translateX(-3px); cursor:help; width:6px; }
  .tl-msna { position:absolute; top:13px; height:8px; border-left:1px dashed rgba(243,156,18,0.80); transform:translateX(-3px); cursor:help; width:6px; }
  .tl-msb  { position:absolute; top:0; height:100%; border-left:2px dashed rgba(243,156,18,0.70); pointer-events:none; }
  .tl-vss-ev { position:absolute; top:4px; height:12px; background:#5dade2; border-radius:1px; opacity:0.85; cursor:default; }
  .tl-vss-pt { position:absolute; top:2px; width:4px; height:16px; background:#5dade2; border-radius:1px; opacity:0.7; cursor:default; }
  .bdg{display:inline-flex;gap:7px;flex-wrap:nowrap;align-items:center;padding:3px 8px;background:#1a1a1a;border-radius:5px;border:1px solid #222;margin-right:8px}
  .di{display:inline-flex;align-items:center;gap:3px;font-size:0.73em;white-space:nowrap;color:#888}
  .dt{width:6px;height:6px;border-radius:50%;flex-shrink:0}
  .dok{background:#27ae60}.dwn{background:#f39c12}.dct{background:#e74c3c}.dnf{background:#5dade2}.dgr{background:#555}
</style>
<script>
    function showToast(msg) {
        var t = document.getElementById('toast');
        if (!t) return;
        t.textContent = msg;
        t.classList.add('show');
        setTimeout(function(){ t.classList.remove('show'); }, 1400);
    }
    function _tlFmt(b) {
        if (!b || b <= 0) return '0 B';
        if (b < 1024) return b + ' B';
        if (b < 1048576) return (b/1024).toFixed(1) + ' KB';
        if (b < 1073741824) return (b/1048576).toFixed(1) + ' MB';
        return (b/1073741824).toFixed(2) + ' GB';
    }
    function _tlShowTip(e, el) {
        var d = JSON.parse(el.dataset.tip), t = document.getElementById('tl-tip');
        var fl = d.fl && d.fl.length ? d.fl.map(function(f){ return '<span class="jt-flag">'+f+'</span>'; }).join('') : '<span style="color:#444">none</span>';
        var bv = d.pb ? '<span style="color:#4ecdc4">'+d.bv+'</span> <span style="color:#555">(was '+d.pb+')</span>' : (d.bv || '&mdash;');
        var durTxt = d.dur < 60 ? d.dur+' min' : Math.floor(d.dur/60)+'h '+(d.dur%60)+'m';
        // Data section
        var procCount = (d.nc||0) + (d.cc||0);
        var procBytes = (d.nb||0) + (d.cb||0);
        var hasSel = d.sb > 0, hasProc = procBytes > 0, hasSent = d.snt > 0;
        var dataHtml = '';
        if (hasSel)  dataHtml += '<div class="jt-row"><span class="jt-lbl">Selected</span><span class="jt-val">'+d.sc2.toLocaleString()+' files &nbsp;<span style="color:#555">('+_tlFmt(d.sb)+')</span></span></div>';
        if (hasProc) dataHtml += '<div class="jt-row"><span class="jt-lbl">Processed</span><span class="jt-val">'+procCount.toLocaleString()+' files &nbsp;<span style="color:#555">('+_tlFmt(procBytes)+')</span></span></div>';
        if (hasProc && (d.nc > 0 || d.cb > 0)) {
            if (d.nc > 0) dataHtml += '<div class="jt-sub"><span class="jt-slbl">New</span><span class="jt-sval">'+d.nc.toLocaleString()+' files &nbsp;('+_tlFmt(d.nb)+')</span></div>';
            if (d.cc > 0) dataHtml += '<div class="jt-sub"><span class="jt-slbl">Changed</span><span class="jt-sval">'+d.cc.toLocaleString()+' files &nbsp;('+_tlFmt(d.cb)+')</span></div>';
        } else if (!hasProc) {
            dataHtml += '<div class="jt-row"><span class="jt-lbl">Changed</span><span class="jt-val">'+d.cc.toLocaleString()+' files &nbsp;<span style="color:#555">('+_tlFmt(d.cb)+')</span></span></div>';
        }
        if (hasSent) dataHtml += '<div class="jt-row" style="margin-top:2px"><span class="jt-lbl">Sent</span><span class="jt-val" style="color:#6ecf6e">'+_tlFmt(d.snt)+'</span></div>';
        if (d.dc > 0) dataHtml += '<div class="jt-row"><span class="jt-lbl">Deleted</span><span class="jt-val">'+d.dc.toLocaleString()+' files</span></div>';
        // Dedup section
        var ddHtml = '';
        if (hasSent && hasProc) {
            var ddProc = procBytes > 0 ? Math.round((1 - d.snt/procBytes)*100) : null;
            var ddProcCol = ddProc >= 80 ? '#27ae60' : ddProc >= 50 ? '#f0a030' : '#e74c3c';
            if (ddProc !== null) ddHtml +=
                '<div class="jt-row"><span class="jt-lbl">vs Processed</span>'+
                '<span class="jt-val" style="color:'+ddProcCol+'">'+ddProc+'% deduped</span></div>'+
                '<div style="height:3px;background:#151515;border-radius:2px;overflow:hidden;margin-bottom:4px">'+
                '<div style="height:100%;width:'+ddProc+'%;background:'+ddProcCol+';border-radius:2px"></div></div>';
        }
        if (hasSent && hasSel) {
            var ddSel = d.sb > 0 ? (d.snt/d.sb*100) : null;
            var ddSelCol = ddSel < 5 ? '#27ae60' : ddSel < 30 ? '#f0a030' : '#e74c3c';
            var ddSelPct = ddSel < 0.1 ? ddSel.toFixed(3) : ddSel < 1 ? ddSel.toFixed(2) : ddSel.toFixed(1);
            if (ddSel !== null) ddHtml +=
                '<div class="jt-row"><span class="jt-lbl">vs Selected</span>'+
                '<span class="jt-val" style="color:'+ddSelCol+'">'+ddSelPct+'% of total sent</span></div>'+
                '<div style="height:3px;background:#151515;border-radius:2px;overflow:hidden;margin-bottom:4px">'+
                '<div style="height:100%;width:'+Math.min(ddSel,100)+'%;background:#4a9eda;border-radius:2px;min-width:3px"></div></div>';
        }
        t.innerHTML =
            '<div class="jt-head">'+d.p+'</div>'+
            '<div class="jt-row"><span class="jt-lbl">Time</span><span class="jt-val">'+d.d+'&nbsp;&nbsp;'+d.s+'&ndash;'+d.e+'</span></div>'+
            '<div class="jt-row"><span class="jt-lbl">Duration</span><span class="jt-val">'+durTxt+'</span></div>'+
            '<div class="jt-row"><span class="jt-lbl">Status</span><span class="jt-val '+d.sc+'">'+d.st+'</span></div>'+
            (d.err > 0 ? '<div class="jt-row"><span class="jt-lbl">Errors</span><span class="jt-val jt-err">'+d.err+'</span></div>' : '')+
            '<div class="jt-row"><span class="jt-lbl">Flags</span><span class="jt-val">'+fl+'</span></div>'+
            (dataHtml ? '<div class="jt-sec-lbl">Data</div>'+dataHtml : '')+
            (ddHtml   ? '<div class="jt-sec-lbl">Deduplication</div>'+ddHtml : '')+
            '<hr class="jt-sep">'+
            '<div class="jt-row"><span class="jt-lbl">Build</span><span class="jt-val">'+bv+'</span></div>';
        t.style.display = 'block';
        _tlMoveTip(e);
    }
    function _tlMoveTip(e) {
        var t = document.getElementById('tl-tip');
        var x = e.clientX + 14, y = e.clientY - 10;
        if (x + 240 > window.innerWidth)  { x = e.clientX - 250; }
        if (y + t.offsetHeight + 10 > window.innerHeight) { y = e.clientY - t.offsetHeight - 4; }
        t.style.left = x + 'px'; t.style.top = y + 'px';
    }
    function _tlHideTip() { document.getElementById('tl-tip').style.display = 'none'; }
    function _wrShowTip(e, el) {
        var d = JSON.parse(el.dataset.wtip), t = document.getElementById('tl-tip');
        var rstHtml = (d.rst || []).map(function(r) {
            var isComment = r.indexOf('#') === 0;
            return '<div style="font-family:Consolas,monospace;font-size:0.9em;padding:1px 0;color:' +
                   (isComment ? '#666' : '#4ecdc4') + '">' + r.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div>';
        }).join('');
        var assessHtml = '';
        if (d.assessment) {
            var aCol = d.assessment === 'Healthy' ? '#27ae60' : d.assessment === 'Critical' ? '#e74c3c' : d.assessment === 'Warning' ? '#f39c12' : '#5dade2';
            assessHtml = '<div class="jt-sec-lbl">Live Assessment</div>' +
                '<div style="margin-bottom:2px"><span style="color:' + aCol + ';font-weight:bold">' + d.assessment + '</span></div>' +
                (d.action ? '<div style="color:#aaa;font-size:0.9em;line-height:1.5;margin-top:2px">' + d.action + '</div>' : '');
        }
        t.innerHTML =
            '<div class="jt-head">' + d.n + '</div>' +
            '<div class="jt-sec-lbl">Purpose</div>' +
            '<div style="color:#ccc;margin-bottom:4px;line-height:1.5">' + d.purp + '</div>' +
            '<div class="jt-row" style="margin-bottom:2px"><span class="jt-lbl">Service</span><span class="jt-val" style="color:#aaa;text-align:right">' + d.svc + '</span></div>' +
            '<div class="jt-row" style="margin-bottom:6px"><span class="jt-lbl">Startup</span><span class="jt-val" style="color:' + (d.start.indexOf('Manual') >= 0 ? '#f39c12' : '#aaa') + '">' + d.start + '</span></div>' +
            '<div class="jt-sec-lbl">Restart sequence</div>' +
            '<div style="background:#151515;border-radius:3px;padding:4px 6px;margin-bottom:6px">' + rstHtml + '</div>' +
            '<div class="jt-sec-lbl">If unhealthy</div>' +
            '<div style="color:' + (d.imp.indexOf('CRITICAL') >= 0 ? '#e74c3c' : '#f39c12') + ';line-height:1.5;margin-bottom:' + (assessHtml ? '6px' : '0') + '">' + d.imp + '</div>' +
            assessHtml;
        t.style.display = 'block';
        t.style.maxWidth = '380px';
        _tlMoveTip(e);
    }
    function notesKey(id) { return 'vss-notes-' + id; }
    function updateNoteCount(id, val) {
        var el = document.getElementById('nc-' + id);
        if (el) el.textContent = '(' + val.length + '/5000)';
    }
    function saveNotes(id) {
        var ta = document.getElementById('notes-' + id);
        if (!ta) return;
        var val = ta.value.trim();
        if (val) { localStorage.setItem(notesKey(id), val); } else { localStorage.removeItem(notesKey(id)); }
        updateNoteCount(id, ta.value);
        showToast('Notes saved');
    }
    function loadAllNotes() {
        document.querySelectorAll('[id^="notes-"]').forEach(function(ta) {
            var id = ta.id.slice(6);
            var val = localStorage.getItem(notesKey(id)) || '';
            ta.value = val;
            updateNoteCount(id, val);
        });
    }
    function buildNotesMailto(href) {
        var blocks = [];
        document.querySelectorAll('[data-accountid]').forEach(function(card) {
            var id = card.dataset.accountid;
            var note = (localStorage.getItem(notesKey(id)) || '').trim();
            if (!note) return;
            var hdr = (card.dataset.machine || id) + (card.dataset.customer ? ' (' + card.dataset.customer + ')' : '');
            var url = card.dataset.console || '';
            blocks.push(hdr + (url ? '\r\n' + url : '') + '\r\n' + note);
        });
        if (!blocks.length) return href;
        var to = (href.match(/^mailto:([^?]*)/) || ['',''])[1];
        var subj = (href.match(/[?&]subject=([^&]*)/) || ['',''])[1];
        var body = 'VSS Notes\r\n========================================\r\n\r\n' + blocks.join('\r\n\r\n');
        return 'mailto:' + to + '?subject=' + subj + '&body=' + encodeURIComponent(body);
    }
    var _hiddenFilters = new Set();
    var _hiddenBackupStatus = new Set();
    var _allSevs = ['RED','YELLOW','GREEN','OFFLINE'];
    var _allBackupStatus = ['Failed','CompletedWithErrors','Completed','Skipped','Unknown'];
    var _typeFilter = 'all';
    var _sevCounts = {
        all:         { RED: $redCount, YELLOW: $yellowCount, GREEN: $greenCount, OFFLINE: $offlineCount },
        server:      { RED: $srvRed,   YELLOW: $srvYellow,   GREEN: $srvGreen,   OFFLINE: $srvOffline   },
        workstation: { RED: $wsRed,    YELLOW: $wsYellow,    GREEN: $wsGreen,    OFFLINE: $wsOffline    }
    };

    function applyFilters() {
        var inp = document.getElementById('dev-search');
        var ql  = inp ? inp.value.toLowerCase().trim() : '';
        document.querySelectorAll('[data-search]').forEach(function(el) {
            var sevHidden   = _hiddenFilters.has(el.dataset.severity);
            var bsHidden    = _hiddenBackupStatus.has(el.dataset.backupStatus);
            var typeHidden  = _typeFilter !== 'all' && el.dataset.deviceType !== _typeFilter;
            var searchMiss  = ql && el.dataset.search.toLowerCase().indexOf(ql) < 0;
            el.style.display = (sevHidden || bsHidden || typeHidden || searchMiss) ? 'none' : '';
        });
        document.querySelectorAll('[data-nav-sev]').forEach(function(el) {
            var sevHidden  = _hiddenFilters.has(el.dataset.navSev);
            var bsHidden   = _hiddenBackupStatus.has(el.dataset.navBackupStatus);
            var typeHidden = _typeFilter !== 'all' && el.dataset.navType !== _typeFilter;
            var searchMiss = ql && (el.dataset.navSearch || '').toLowerCase().indexOf(ql) < 0;
            el.style.display = (sevHidden || bsHidden || typeHidden || searchMiss) ? 'none' : 'flex';
        });
        // Hide customer group headers when all their device links are filtered out
        document.querySelectorAll('[data-nav-group]:not([data-nav-sev])').forEach(function(hdr) {
            var group = hdr.dataset.navGroup;
            var links = document.querySelectorAll('[data-nav-sev][data-nav-group="' + group + '"]');
            var anyVisible = false;
            links.forEach(function(l) { if (l.style.display !== 'none') anyVisible = true; });
            hdr.style.display = anyVisible ? '' : 'none';
        });
        // Update severity counts for active device type
        var counts = _sevCounts[_typeFilter] || _sevCounts.all;
        _allSevs.forEach(function(s) {
            var el = document.getElementById('sev-count-' + s);
            if (el) el.textContent = counts[s] || 0;
        });
        // Update device type card active states
        ['all','server','workstation'].forEach(function(t) {
            var card = document.querySelector('[data-filter-type="' + t + '"]');
            if (!card) return;
            var active = _typeFilter === t;
            card.style.opacity     = active ? '1' : '0.5';
            card.style.borderColor = active ? '#4ecdc4' : '#3a3a3a';
            var countEl = document.getElementById('dt-count-' + t);
            if (countEl) countEl.style.color = active ? '#4ecdc4' : '#888';
        });
        _allSevs.forEach(function(s) {
            var card = document.querySelector('[data-filter-sev="' + s + '"]');
            var icon = document.getElementById('fi-' + s);
            if (!card || !icon) return;
            var hidden = _hiddenFilters.has(s);
            card.style.opacity = hidden ? '0.35' : '1';
            icon.textContent = hidden ? '\uD83D\uDE48' : '\uD83D\uDC41';
            icon.title = hidden ? 'Hidden \u2014 click to add' : (_hiddenFilters.size === 0 ? 'Click to focus' : 'Click to hide');
        });
        // Update backup status filter states
        _allBackupStatus.forEach(function(bs) {
            var card = document.querySelector('[data-filter-bs="' + bs + '"]');
            var icon = document.getElementById('fi-bs-' + bs);
            if (!card || !icon) return;
            var hidden = _hiddenBackupStatus.has(bs);
            card.style.opacity = hidden ? '0.35' : '1';
            icon.textContent = hidden ? '\uD83D\uDE48' : '\uD83D\uDC41';
            icon.title = hidden ? 'Hidden \u2014 click to add' : (_hiddenBackupStatus.size === 0 ? 'Click to focus' : 'Click to hide');
        });
    }

    function filterSearch(q) { applyFilters(); }

    function toggleType(type) { _typeFilter = type; applyFilters(); }

    function toggleFilter(sev) {
        var allVisible = _hiddenFilters.size === 0;
        if (allVisible) {
            _allSevs.forEach(function(s) { if (s !== sev) _hiddenFilters.add(s); });
        } else if (_hiddenFilters.has(sev)) {
            _hiddenFilters.delete(sev);
        } else {
            _hiddenFilters.add(sev);
            if (_hiddenFilters.size === _allSevs.length) _hiddenFilters.clear();
        }
        applyFilters();
    }

    function toggleBackupStatusFilter(bs) {
        var allVisible = _hiddenBackupStatus.size === 0;
        if (allVisible) {
            _allBackupStatus.forEach(function(s) { if (s !== bs) _hiddenBackupStatus.add(s); });
        } else if (_hiddenBackupStatus.has(bs)) {
            _hiddenBackupStatus.delete(bs);
        } else {
            _hiddenBackupStatus.add(bs);
            if (_hiddenBackupStatus.size === _allBackupStatus.length) _hiddenBackupStatus.clear();
        }
        applyFilters();
    }

    function tlPage(id, dir) {
        var vp = document.getElementById(id);
        if (!vp) return;
        var step = Math.round((vp.clientWidth || 600) * 0.75);
        vp.scrollLeft += dir * step;
        tlArrows(id);
    }
    function tlArrows(id) {
        var vp = document.getElementById(id);
        if (!vp) return;
        var par = vp.parentElement;
        var lb = par ? par.querySelector('.tl-lbtn') : null;
        var rb = par ? par.querySelector('.tl-rbtn') : null;
        if (lb) lb.disabled = vp.scrollLeft <= 0;
        if (rb) rb.disabled = vp.scrollLeft >= vp.scrollWidth - vp.clientWidth - 2;
    }
    function tlScrollToNow(vp) {
        [0, 50, 150, 400, 900].forEach(function(ms) {
            setTimeout(function() {
                if (vp.scrollWidth > vp.clientWidth + 10) {
                    vp.scrollLeft = vp.scrollWidth;
                    tlArrows(vp.id);
                }
            }, ms);
        });
    }
    document.addEventListener('DOMContentLoaded', function() {
        loadAllNotes();
        document.querySelectorAll('.tl-vp').forEach(function(vp) {
            tlScrollToNow(vp);
            vp.addEventListener('scroll', function(){ tlArrows(vp.id); });
        });
        document.querySelectorAll('details').forEach(function(det) {
            det.addEventListener('toggle', function() {
                if (!det.open) return;
                det.querySelectorAll('.tl-vp').forEach(function(vp) { tlScrollToNow(vp); });
            });
        });
    });

    function copyThenEmail(reportPath, mailtoHref) {
        mailtoHref = buildNotesMailto(mailtoHref);
        var launch = function(){ if (mailtoHref) { window.location.href = mailtoHref; } };
        if (!reportPath) { launch(); return false; }
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(reportPath).then(function(){
                showToast('Report copied to clipboard');
                setTimeout(launch, 180);
            }).catch(function(){
                window.prompt('Copy report path:', reportPath);
                launch();
            });
            return false;
        }
        window.prompt('Copy report path:', reportPath);
        launch();
        return false;
    }
</script>
</head>
<body>
<div id="toast"></div>
<div id="tl-tip"></div>

<div id="sidebar">
  <div id="sidebar-header">
    <div style="font-weight:bold;font-size:0.95em;color:#4ecdc4">$PartnerLabel</div>
    <div style="font-size:0.78em;color:#888">VSS Health Report</div>
    <div style="font-size:0.75em;color:#666;margin-top:2px">$ReportDate</div>
  </div>
  <a href="#exec-summary" style="display:block;padding:8px 10px;color:#f0f0f0;text-decoration:none;border-bottom:2px solid #333;font-size:0.9em;background:#222">&#128202; Executive Summary</a>
  <div style="padding:6px 8px;border-bottom:1px solid #222">
    <input id="dev-search" type="text" placeholder="&#128269; device / company..." oninput="filterSearch(this.value)" style="width:100%;background:#1a1a1a;border:1px solid #2a2a2a;color:#ccc;padding:5px 8px;border-radius:4px;font-size:0.82em;outline:none;box-sizing:border-box">
  </div>
  $navItems
</div>

<div id="main">
    <div style="display:flex;justify-content:space-between;align-items:center;gap:12px">
        <h1 style="margin-bottom:0">VSS Health Report</h1>
        <div style="display:flex;align-items:center;gap:8px">
            $(if($ReportFolderHref){"<a href='$ReportFolderHref' title='Open report folder' style='display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border:1px solid #3a3a3a;border-radius:6px;background:#1e1e1e;color:#d0d0d0;text-decoration:none;font-size:1.0em'>&#128194;</a>"})
            $(if($EmailHref){"<a href='$EmailHref' data-mailto=`"$emailAttr`" data-report=`"$reportPathAttr`" onclick='return copyThenEmail(this.dataset.report, this.dataset.mailto);' title='Copy report and email' style='display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border:1px solid #3a3a3a;border-radius:6px;background:#1e1e1e;color:#d0d0d0;text-decoration:none;font-size:1.05em'>&#9993;</a>"})
        </div>
    </div>
  <div style="color:#888;margin-bottom:4px">Partner: $PartnerLabel (ID: $PartnerId) &nbsp;|&nbsp; Generated: $genTime$(if($ScriptName){" &nbsp;|&nbsp; $ScriptName"})$(if($RunDuration){" &nbsp;|&nbsp; Duration: $RunDuration"})</div>

  <div id="exec-summary" style="background:#1a1a1a;border:1px solid #444;border-radius:8px;padding:20px;margin-bottom:28px;margin-top:16px">
    <h2 style="margin-bottom:14px">Executive Summary</h2>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px">
      <div data-filter-type="all" onclick="toggleType('all')" title="Show all devices" style="background:#141c1c;border:1px solid #4ecdc4;border-radius:6px;padding:14px;text-align:center;cursor:pointer;user-select:none;transition:opacity .15s">
        <div id="dt-count-all" style="font-size:2.2em;font-weight:bold;color:#4ecdc4">$totalCount</div>
        <div style="font-size:0.85em;color:#aaa">All Devices</div>
      </div>
      <div data-filter-type="server" onclick="toggleType('server')" title="Show servers only" style="background:#0e1a1a;border:1px solid #3a3a3a;border-radius:6px;padding:14px;text-align:center;cursor:pointer;user-select:none;transition:opacity .15s;opacity:0.5">
        <div id="dt-count-server" style="font-size:2.2em;font-weight:bold;color:#888">$srvTotal</div>
        <div style="font-size:0.85em;color:#aaa">Servers</div>
      </div>
      <div data-filter-type="workstation" onclick="toggleType('workstation')" title="Show workstations only" style="background:#0e1a1a;border:1px solid #3a3a3a;border-radius:6px;padding:14px;text-align:center;cursor:pointer;user-select:none;transition:opacity .15s;opacity:0.5">
        <div id="dt-count-workstation" style="font-size:2.2em;font-weight:bold;color:#888">$wsTotal</div>
        <div style="font-size:0.85em;color:#aaa">Workstations</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px">
      <div data-filter-sev="RED" onclick="toggleFilter('RED')" title="Click to show/hide Critical devices" style="background:#2c1010;border:1px solid #e74c3c;border-radius:6px;padding:14px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-RED" title="Visible &mdash; click to hide" style="position:absolute;top:6px;right:8px;font-size:0.9em">&#128065;</span>
        <div id="sev-count-RED" style="font-size:2.2em;font-weight:bold;color:#e74c3c">$redCount</div>
        <div style="font-size:0.85em;color:#aaa">Critical (RED)</div>
      </div>
      <div data-filter-sev="YELLOW" onclick="toggleFilter('YELLOW')" title="Click to show/hide Warning devices" style="background:#2c2010;border:1px solid #f39c12;border-radius:6px;padding:14px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-YELLOW" title="Visible &mdash; click to hide" style="position:absolute;top:6px;right:8px;font-size:0.9em">&#128065;</span>
        <div id="sev-count-YELLOW" style="font-size:2.2em;font-weight:bold;color:#f39c12">$yellowCount</div>
        <div style="font-size:0.85em;color:#aaa">Warning (YELLOW)</div>
      </div>
      <div data-filter-sev="GREEN" onclick="toggleFilter('GREEN')" title="Click to show/hide OK devices" style="background:#102c10;border:1px solid #27ae60;border-radius:6px;padding:14px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-GREEN" title="Visible &mdash; click to hide" style="position:absolute;top:6px;right:8px;font-size:0.9em">&#128065;</span>
        <div id="sev-count-GREEN" style="font-size:2.2em;font-weight:bold;color:#27ae60">$greenCount</div>
        <div style="font-size:0.85em;color:#aaa">OK (GREEN)</div>
      </div>
      <div data-filter-sev="OFFLINE" onclick="toggleFilter('OFFLINE')" title="Click to show/hide Offline devices" style="background:#1a1a1a;border:1px solid #555;border-radius:6px;padding:14px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-OFFLINE" title="Visible &mdash; click to hide" style="position:absolute;top:6px;right:8px;font-size:0.9em">&#128065;</span>
        <div id="sev-count-OFFLINE" style="font-size:2.2em;font-weight:bold;color:#7f8c8d">$offlineCount</div>
        <div style="font-size:0.85em;color:#aaa">Offline / No Data</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px">
      <div data-filter-bs="Failed" onclick="toggleBackupStatusFilter('Failed')" title="Click to show/hide Failed devices" style="background:#2c1010;border:1px solid #e74c3c;border-radius:6px;padding:12px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-bs-Failed" title="Visible &mdash; click to hide" style="position:absolute;top:4px;right:6px;font-size:0.85em">&#128065;</span>
        <div style="font-size:1.8em;font-weight:bold;color:#e74c3c">✕</div>
        <div style="font-size:0.8em;color:#aaa">Failed</div>
      </div>
      <div data-filter-bs="CompletedWithErrors" onclick="toggleBackupStatusFilter('CompletedWithErrors')" title="Click to show/hide CompletedWithErrors devices" style="background:#2c2010;border:1px solid #f39c12;border-radius:6px;padding:12px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-bs-CompletedWithErrors" title="Visible &mdash; click to hide" style="position:absolute;top:4px;right:6px;font-size:0.85em">&#128065;</span>
        <div style="font-size:1.8em;font-weight:bold;color:#f39c12">⚠</div>
        <div style="font-size:0.8em;color:#aaa">With Errors</div>
      </div>
      <div data-filter-bs="Completed" onclick="toggleBackupStatusFilter('Completed')" title="Click to show/hide Completed devices" style="background:#102c10;border:1px solid #27ae60;border-radius:6px;padding:12px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-bs-Completed" title="Visible &mdash; click to hide" style="position:absolute;top:4px;right:6px;font-size:0.85em">&#128065;</span>
        <div style="font-size:1.8em;font-weight:bold;color:#27ae60">✓</div>
        <div style="font-size:0.8em;color:#aaa">Completed</div>
      </div>
      <div data-filter-bs="Skipped" onclick="toggleBackupStatusFilter('Skipped')" title="Click to show/hide Skipped devices" style="background:#1a1a1a;border:1px solid #666;border-radius:6px;padding:12px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-bs-Skipped" title="Visible &mdash; click to hide" style="position:absolute;top:4px;right:6px;font-size:0.85em">&#128065;</span>
        <div style="font-size:1.8em;font-weight:bold;color:#888">⊘</div>
        <div style="font-size:0.8em;color:#aaa">Skipped</div>
      </div>
      <div data-filter-bs="Unknown" onclick="toggleBackupStatusFilter('Unknown')" title="Click to show/hide Unknown status devices" style="background:#1a1a1a;border:1px solid #555;border-radius:6px;padding:12px;text-align:center;cursor:pointer;position:relative;user-select:none;transition:opacity .15s">
        <span id="fi-bs-Unknown" title="Visible &mdash; click to hide" style="position:absolute;top:4px;right:6px;font-size:0.85em">&#128065;</span>
        <div style="font-size:1.8em;font-weight:bold;color:#aaa">?</div>
        <div style="font-size:0.8em;color:#aaa">Unknown</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
      <div>
        <div style="font-weight:bold;color:#ccc;margin-bottom:8px">Common Patterns</div>
        <div style="font-size:0.88em;color:#ddd">$patternHtml</div>
      </div>
      <div>
        <div style="font-weight:bold;color:#ccc;margin-bottom:8px">Fleet Status</div>
        <div style="font-size:0.88em;color:#ddd">
          <div>&#8226; Total devices analyzed: <strong>$totalCount</strong></div>
          <div>&#8226; Backup status: <span style="color:#e74c3c">$failedCount Failed</span>, <span style="color:#f39c12">$errorsCount CompletedWithErrors</span></div>
          <div>&#8226; Devices with unhealthy writers: <strong style="color:#e74c3c">$unhWriters</strong></div>
          <div>&#8226; Devices with S1 coverage &lt;7 days: <strong style="color:#e74c3c">$lowCoverage</strong></div>
          <div>&#8226; Devices behind S1 schedule: <strong style="color:#f39c12">$behindSched</strong></div>
          $(if($s1MissCount -gt 0){"<div>&#8226; Devices missing S1 (customer inconsistency): <strong style='color:#f39c12'>$s1MissCount</strong></div>"})
          <div>&#8226; Devices offline/unreachable: <strong style="color:#7f8c8d">$offlineCount</strong></div>
        </div>
      </div>
    </div>
    
    <!-- Fleet Software & Tools Section -->
    <div style="border-top:2px solid #333;margin-top:16px;padding-top:14px">
      <div style="font-weight:bold;color:#ccc;margin-bottom:10px;cursor:pointer;user-select:none" onclick="document.getElementById('fleet-sw-details').style.display = document.getElementById('fleet-sw-details').style.display === 'none' ? 'block' : 'none'">▼ Fleet Software & Tools</div>
      <div id="fleet-sw-details" style="font-size:0.88em;color:#ddd;display:none">
        $fleetSoftwareHtml
      </div>
    </div>
  </div>

  $devSections
</div>
</body>
</html>
"@
}

# ============================================================================
# Main
# ============================================================================

Write-Host "`nCove VSS Health Report Generator" -ForegroundColor Cyan
Write-Host "Partner: $PartnerLabel (ID: $PartnerId)`n" -ForegroundColor Cyan
$host.UI.RawUI.WindowTitle = "Cove VSS Report - $PartnerLabel (ID: $PartnerId)"
$ScriptStartTime = [datetime]::Now
$Script:TlDebugRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$Script:TlDebug     = $DebugTimeline.IsPresent
$ScriptName      = Split-Path $PSCommandPath -Leaf

# 1. Credentials + Login
$creds = Get-CoveCredentials -CredFile $CredentialFile -Clear:$ClearCredentials
$visa  = Invoke-CoveLogin -Creds $creds

# 2. Device list
$devicePool = Get-CoveDevices -Visa $visa -PartnerId $PartnerId
if ($TestAccountIds -and $TestAccountIds.Count -gt 0) {
    Write-Host "Filtering to test AccountIds: $($TestAccountIds -join ', ')" -ForegroundColor Yellow
    $devicePool = @($devicePool | Where-Object { $_.AccountId -in $TestAccountIds })
    if ($devicePool.Count -eq 0) { Write-Host "No devices found matching AccountIds: $($TestAccountIds -join ', ')" -ForegroundColor Red; exit 1 }
}

# Apply T0 filter unless -AllDevices is set or specific test devices were requested
if ($AllDevices -or $T0Filter.Count -eq 0 -or ($TestAccountIds -and $TestAccountIds.Count -gt 0)) {
    $devices = $devicePool | Select-Object -First $MaxDevices
    Write-Host "Scanning all $($devices.Count) device(s) (no T0 filter)." -ForegroundColor Cyan
} else {
    $devices = @($devicePool | Where-Object { $_.T0 -in $T0Filter } | Select-Object -First $MaxDevices)
    Write-Host "Filtered to $($devices.Count) device(s) with T0 in ($($T0Filter -join ',')) - Failed/CompletedWithErrors." -ForegroundColor Cyan
}
if ($devices.Count -eq 0) { Write-Host "No devices match the filter. Exiting." -ForegroundColor Yellow; exit 0 }

# 3. Per-device: fetch InternalInfo, parse, analyze (parallel, PS7)
$total = $devices.Count

# Capture helper function definitions for parallel runspaces
$_fnNames = @(
    'Invoke-CoveAPI', 'Get-InternalInfoUrl', 'Get-DeviceRepserv',
    'Get-DeviceLastSessions', 'Invoke-InternalInfoPage', 'Strip-HtmlTags',
    'Split-InternalInfoSections', 'Get-HtmlTableData', 'Parse-DlSection',
    'Parse-TableSection', 'Parse-EventLogs', 'Get-GBValue', 'Parse-UtcDateTime', 'Format-HM', 'Analyze-Device',
    'Format-RelativeTime', 'Get-SessionAgeColor', 'Get-DataSourcesHtml',
    'Save-SessionsToCache', 'Load-SessionsFromCache'
)
$_initSb = [scriptblock]::Create(($_fnNames | ForEach-Object {
    $fn   = $_
    $body = (Get-Item "Function:$fn").ScriptBlock
    "function $fn {`n$body`n}"
}) -join "`n`n")
$_initStr  = $_initSb.ToString()   # must pass as string - $using: blocks scriptblocks
$_pVisa    = $visa
$_pApiUrl  = $COVE_API_URL
$_pTimeout = $DiagTimeoutSec
# Resolve cache directory and ensure it exists
$_pCachePath = if ($CachePath) { $CachePath } else { Join-Path $PSScriptRoot 'cache\internalinfo' }
if (-not (Test-Path $_pCachePath)) { New-Item -ItemType Directory -Path $_pCachePath -Force | Out-Null }

Write-Host "Processing $total device(s) with $ParallelCount parallel workers..." -ForegroundColor Cyan

$results = @($devices | ForEach-Object -Parallel {
    . ([scriptblock]::Create($using:_initStr))
    $COVE_API_URL = $using:_pApiUrl
    $visa         = $using:_pVisa
    $dev          = $_
    $devTypeTag   = switch ([int]([int]::TryParse(($dev.OT ?? '0'), [ref]$null) ? ($dev.OT -as [int]) ?? 0 : 0)) { 1 { 'WRK' } 2 { 'SVR' } default { '' } }
    $devTypeStr   = if ($devTypeTag) { " [$devTypeTag]" } else { '' }
    $label        = "$($dev.MN ?? $dev.DeviceName) ($($dev.DeviceName), AccountId=$($dev.AccountId))$devTypeStr"

    try {
        $infoUrl  = Get-InternalInfoUrl -Visa $visa -AccountId $dev.AccountId
        $sections = @{}
        
        # Cache directory - available for both InternalInfo and sessions
        $cacheDir    = $using:_pCachePath
        $safeName    = ($dev.DeviceName ?? 'unknown') -replace '[\\/:*?"<>|]', '-'
        $cachePrefix = "$safeName-$($dev.AccountId)-"
        $cacheTtlMin = 5

        if ($infoUrl) {
            # Cache: look for a file younger than 15 min for this device
            $cachedFile  = Get-ChildItem -Path $cacheDir -Filter "$cachePrefix*.html" -ErrorAction SilentlyContinue |
                           Sort-Object LastWriteTime -Descending | Select-Object -First 1
            $html = $null
            if ($cachedFile -and ((Get-Date) - $cachedFile.LastWriteTime).TotalMinutes -lt $cacheTtlMin) {
                $html = [IO.File]::ReadAllText($cachedFile.FullName)
                Write-Host "  $label [cache hit: $($cachedFile.Name)]" -ForegroundColor DarkCyan
            } else {
                $html = Invoke-InternalInfoPage -Url $infoUrl -TimeoutSec $using:_pTimeout
                if ($html) {
                    # Save to cache with self-describing filename: name-accountId-yyyyMMdd-HHmmss.html
                    $stamp     = (Get-Date).ToString('yyyyMMdd-HHmmss')
                    $cacheFile = Join-Path $cacheDir "$cachePrefix$stamp.html"
                    [IO.File]::WriteAllText($cacheFile, $html)
                }
            }
            if ($html) {
                $sections = Split-InternalInfoSections -Html $html
                Write-Host "  $label [$(($sections.Keys | Measure-Object).Count) sections]" -ForegroundColor DarkGray
            } else {
                Write-Host "  $label [fetch failed]" -ForegroundColor DarkYellow
            }
        } else {
            $html = $null
            Write-Host "  $label [no endpoint]" -ForegroundColor DarkYellow
        }

        # Session data - cloud repserv call, works even when device is offline
        $repserv      = Get-DeviceRepserv -Visa $visa -AccountId $dev.AccountId
        $lastSessions = if ($repserv) { Get-DeviceLastSessions -Visa $visa -Repserv $repserv -AccountId $dev.AccountId } else { @{} }
        $sessionFailed = (-not $lastSessions -or $lastSessions.Count -eq 0)

        # Cache sessions to CSV for analysis
        if (-not $sessionFailed) {
            Save-SessionsToCache -CacheDir $cacheDir -DeviceName ($dev.DeviceName ?? 'unknown') -AccountId $dev.AccountId -LastSessions $lastSessions
        }

        $result = Analyze-Device -Device $dev -Sections $sections -LastSessions $lastSessions `
            -RollbackWarnBelow $using:RollbackWarnBelow -RollbackTarget $using:RollbackTarget `
            -RollbackWarnAbove $using:RollbackWarnAbove -ShadowAllocMinPct $using:ShadowAllocMinPct `
            -MaxSnapPerVol $using:MaxSnapPerVol

        $result._retry_info    = ($infoUrl -and -not $html)
        $result._retry_session = $sessionFailed

        $icon = switch ($result.severity) { 'RED' { 'RED' }; 'YELLOW' { 'WARN' }; 'GREEN' { 'OK' }; 'OFFLINE' { 'OFFLINE' }; default { '?' } }
        $col  = switch ($result.severity) { 'RED' { 'Red' }; 'YELLOW' { 'Yellow' }; 'GREEN' { 'Green' }; default { 'Gray' } }
        Write-Host "  -> $label $icon ($($result.issues.Count) issues)$(if ($result._retry_info -or $result._retry_session) { ' [RETRY queued]' })" -ForegroundColor $col

        $result
    } catch {
        Write-Host "  [ERR] $label : $_" -ForegroundColor Red
        $null
    }
} -ThrottleLimit $ParallelCount | Where-Object { $_ })

# 3a. Retry: devices where InternalInfo or sessions failed on first attempt (1 retry each)
$retryDevices = @($devices | Where-Object {
    $aid = $_.AccountId
    $r   = $results | Where-Object { $_.account_id -eq $aid } | Select-Object -First 1
    $r -and ($r._retry_info -or $r._retry_session)
})
if ($retryDevices.Count -gt 0) {
    Write-Host "`nRetrying $($retryDevices.Count) device(s) with missing data..." -ForegroundColor Yellow
    $retryResults = @($retryDevices | ForEach-Object -Parallel {
        . ([scriptblock]::Create($using:_initStr))
        $visa        = $using:_pVisa
        $dev         = $_
        $cacheDir    = $using:_pCachePath
        $safeName    = ($dev.DeviceName ?? 'unknown') -replace '[\\/:*?"<>|]', '-'
        $cachePrefix = "$safeName-$($dev.AccountId)-"
        $label       = "$($dev.MN ?? $dev.DeviceName) ($($dev.DeviceName), AccountId=$($dev.AccountId)) [RETRY]"

        # Find which fetches failed
        $aid  = $dev.AccountId
        $prev = $using:results | Where-Object { $_.account_id -eq $aid } | Select-Object -First 1

        try {
            $sections     = @{}
            $lastSessions = if ($prev) { $prev.last_sessions } else { @{} }

            # Retry InternalInfo if it failed
            if ($prev._retry_info) {
                $infoUrl = Get-InternalInfoUrl -Visa $visa -AccountId $dev.AccountId
                if ($infoUrl) {
                    $html = Invoke-InternalInfoPage -Url $infoUrl -TimeoutSec $using:_pTimeout
                    if ($html) {
                        $stamp     = (Get-Date).ToString('yyyyMMdd-HHmmss')
                        $cacheFile = Join-Path $cacheDir "$cachePrefix$stamp.html"
                        [IO.File]::WriteAllText($cacheFile, $html)
                        $sections = Split-InternalInfoSections -Html $html
                        Write-Host "  $label [info retry OK: $(($sections.Keys | Measure-Object).Count) sections]" -ForegroundColor Cyan
                    } else {
                        Write-Host "  $label [info retry failed again]" -ForegroundColor DarkYellow
                    }
                }
            } else {
                # Re-use whatever sections were parsed originally (not available across runspace boundary)
                # so just re-parse from cache if present
                $cachedFile = Get-ChildItem -Path $cacheDir -Filter "$cachePrefix*.html" -ErrorAction SilentlyContinue |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($cachedFile) {
                    $html = [IO.File]::ReadAllText($cachedFile.FullName)
                    $sections = Split-InternalInfoSections -Html $html
                }
            }

            # Retry sessions if they failed
            if ($prev._retry_session) {
                $repserv      = Get-DeviceRepserv -Visa $visa -AccountId $dev.AccountId
                $lastSessions = if ($repserv) { Get-DeviceLastSessions -Visa $visa -Repserv $repserv -AccountId $dev.AccountId } else { @{} }
                if ($lastSessions -and $lastSessions.Count -gt 0) {
                    Save-SessionsToCache -CacheDir $cacheDir -DeviceName ($dev.DeviceName ?? 'unknown') -AccountId $dev.AccountId -LastSessions $lastSessions
                    Write-Host "  $label [session retry OK: $($lastSessions.Count) DS]" -ForegroundColor Cyan
                } else {
                    Write-Host "  $label [session retry failed again]" -ForegroundColor DarkYellow
                }
            }

            $result = Analyze-Device -Device $dev -Sections $sections -LastSessions $lastSessions `
                -RollbackWarnBelow $using:RollbackWarnBelow -RollbackTarget $using:RollbackTarget `
                -RollbackWarnAbove $using:RollbackWarnAbove -ShadowAllocMinPct $using:ShadowAllocMinPct `
                -MaxSnapPerVol $using:MaxSnapPerVol
            $result._retry_info    = $false
            $result._retry_session = $false
            $result
        } catch {
            Write-Host "  [ERR retry] $label : $_" -ForegroundColor Red
            $prev  # fall back to original result on retry error
        }
    } -ThrottleLimit $ParallelCount | Where-Object { $_ })

    # Replace original results with retry results (keyed by account_id)
    $retryMap = @{}
    foreach ($r in $retryResults) { $retryMap[$r.account_id] = $r }
    $results  = @($results | ForEach-Object {
        if ($retryMap.ContainsKey($_.account_id)) { $retryMap[$_.account_id] } else { $_ }
    })

    # Report retry outcomes
    $retryOk   = [System.Collections.Generic.List[string]]::new()
    $retryFail = [System.Collections.Generic.List[string]]::new()
    foreach ($orig in $retryDevices) {
        $aid  = $orig.AccountId
        $prev = $retryDevices | Where-Object { $_.AccountId -eq $aid } | Select-Object -First 1
        $origR = ($retryMap[$aid])
        # A retry succeeded if sessions or info improved vs what was queued
        $hadSessions = $origR -and $origR.last_sessions -and $origR.last_sessions.Count -gt 0
        $label2 = "$($orig.MN ?? $orig.DeviceName) (AccountId=$aid)"
        if ($hadSessions -or (-not $origR._retry_session)) {
            $retryOk.Add($label2) | Out-Null
        } else {
            $retryFail.Add($label2) | Out-Null
        }
    }
    if ($retryOk.Count -gt 0) {
        Write-Host "  Retry recovered ($($retryOk.Count)):" -ForegroundColor Green
        foreach ($n in $retryOk) { Write-Host "    + $n" -ForegroundColor Green }
    }
    if ($retryFail.Count -gt 0) {
        Write-Host "  Still missing data ($($retryFail.Count)):" -ForegroundColor DarkYellow
        foreach ($n in $retryFail) { Write-Host "    - $n" -ForegroundColor DarkYellow }
    }
    Write-Host "Retry complete. $($retryOk.Count) recovered, $($retryFail.Count) still missing." -ForegroundColor Cyan
}

# 4a. S1 cross-customer inconsistency check
# Flag Windows devices missing S1 when sibling devices at the same customer have it.
$custGroups = $results | Group-Object { $_.customer_name }
foreach ($grp in $custGroups) {
    $windows   = @($grp.Group | Where-Object { -not $_.is_posix })
    $withS1    = @($windows   | Where-Object { $_.s1_installed })
    $withoutS1 = @($windows   | Where-Object { -not $_.s1_installed })
    if ($withS1.Count -ge [math]::Ceiling($windows.Count / 2.0) -and $withoutS1.Count -gt 0) {
        foreach ($dev in $withoutS1) {
            $dev.s1_inconsistency = $true
            $dev.issues.Add("[s1-miss]S1 not detected on this device, but $($withS1.Count) other device(s) at this customer have S1 installed - possible missed deployment") | Out-Null
            if ($dev.severity -eq 'GREEN') { $dev.severity = 'YELLOW' }
        }
    }
}

# 4. Generate report
Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$outputFile = Join-Path $OutputPath "vss-$(($PartnerLabel -replace '\s+','-').ToLower())-$PartnerId-$(Get-Date -Format 'HHmmss').html"
$RunDuration = "$([math]::Round(([datetime]::Now - $ScriptStartTime).TotalSeconds, 1))s"

# Summary
$red    = @($results | Where-Object { $_.severity -eq "RED" }).Count
$yellow = @($results | Where-Object { $_.severity -eq "YELLOW" }).Count
$green  = @($results | Where-Object { $_.severity -eq "GREEN" }).Count
$off    = @($results | Where-Object { $_.severity -eq "OFFLINE" }).Count

$emailHref = ""
$reportFolderHref = ""
$mailtoHref = ""
$subjectText = ""
$bodyText = ""
if ($ShowEmailIcon -or $GenerateEmailAndCopyFile -or $OpenEmailDraftWithAttachment) {
    $tokens = @{
        PartnerLabel = $PartnerLabel
        PartnerId    = $PartnerId
        ReportDate   = $REPORT_DATE
        ReportPath   = $outputFile
        GeneratedUtc = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm') + ' UTC'
        RedCount     = $red
        YellowCount  = $yellow
        GreenCount   = $green
        OfflineCount = $off
    }
    $subjectText = Expand-TemplateText -Template $EmailSubjectTemplate -Tokens $tokens
    $bodyText    = Expand-TemplateText -Template $EmailBodyTemplate -Tokens $tokens
    $subjectEsc  = [uri]::EscapeDataString($subjectText)
    $bodyEsc     = [uri]::EscapeDataString($bodyText)
    $attachEsc   = [uri]::EscapeDataString($outputFile)
    $mailtoHref  = "mailto:${EmailTo}?subject=$subjectEsc&body=$bodyEsc&attachment=$attachEsc"
    $emailHref = $mailtoHref
}

if ($ShowOpenFolderIcon) {
    $reportFolderHref = ([uri](Split-Path $outputFile -Parent)).AbsoluteUri
}

$html = Build-HtmlReport -Devices $results -PartnerLabel $PartnerLabel -PartnerId $PartnerId -ReportDate $REPORT_DATE -ScriptName $ScriptName -RunDuration $RunDuration -EmailHref $emailHref -ReportFolderHref $reportFolderHref -ReportFilePath $outputFile
$html | Out-File -FilePath $outputFile -Encoding UTF8 -Force

if ($DebugTimeline -and $Script:TlDebugRows.Count -gt 0) {
    $debugCsv = [System.IO.Path]::ChangeExtension($outputFile, "$(Get-Date -Format 'HHmmss').timeline-debug.csv")
    $Script:TlDebugRows | Export-Csv -Path $debugCsv -NoTypeInformation -Force
    Write-Host "Debug CSV: $debugCsv ($($Script:TlDebugRows.Count) rows)" -ForegroundColor Yellow
    Start-Process $debugCsv
}

Write-Host "`nReport saved: $outputFile" -ForegroundColor Green
Write-Host "Summary: RED=$red  YELLOW=$yellow  GREEN=$green  OFFLINE=$off`n" -ForegroundColor Cyan

if ($GenerateEmailAndCopyFile) {
    $copied = Set-ReportFileClipboard -FilePath $outputFile
    if ($copied) {
        Write-Host "Report file copied to clipboard (file object)." -ForegroundColor Green
    } else {
        Write-Host "Could not copy report file to clipboard." -ForegroundColor Yellow
    }
    if ($mailtoHref) {
        Start-Process $mailtoHref
    }
}

if ($CopyReportFileToClipboard -and -not $GenerateEmailAndCopyFile) {
    $copied = Set-ReportFileClipboard -FilePath $outputFile
    if ($copied) {
        Write-Host "Report file copied to clipboard (file object)." -ForegroundColor Green
    } else {
        Write-Host "Could not copy report file to clipboard." -ForegroundColor Yellow
    }
}

if ($OpenEmailDraftWithAttachment -and $ShowEmailIcon) {
    $opened = Open-OutlookDraftWithAttachment -To $EmailTo -Subject $subjectText -Body $bodyText -AttachmentPath $outputFile
    if (-not $opened -and $mailtoHref) {
        Start-Process $mailtoHref
    }
}

if ($OpenReport) {
    Start-Process $outputFile
}


