<# ----- About: ----
    # Cove Data Protection - Combined Device Profile / LSV / Storage Status + Profile Summary Report
    # Revision v1.3 - 2026-09-17 (v05 - restore standard credential file)
    # Author: Eric Harless, Head Backup Nerd - N-able
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
    # Requires Microsoft Excel installed locally (uses the Excel COM object to write the .xlsx)
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials
    # Authenticate to https://api.backup.management/jsonapi
    #
    # Combines two prior standalone scripts (Get-DeviceProfileReport.ps1 + Get-AllProfilesSummary.ps1)
    # into a single .xlsx workbook with two worksheets:
    #
    #   Sheet "DeviceProfileReport" - one row per device under the resolved partner (and all sub-partners) via
    #       EnumerateAccountStatistics: AccountId, DeviceName, MachineName, CompanyName,
    #       OwningPartnerId, OS, DeviceType, ProfileId, ProfileName, ProfileVersion, LSVEnabled,
    #       LSVStatus, StorageStatus, LastSeen.
    #
    #   Sheet "ProfileSummary" - one row per UNIQUE backup profile found anywhere in the partner
    #       tree. EnumerateAccountProfiles(partnerId) only returns a partner's OWN profiles plus
    #       profiles INHERITED from ANCESTOR partners - it never returns profiles owned by CHILD/
    #       sub-partners. A profile can be created at ANY partner level, so to find every unique
    #       profile in the tree this script:
    #         1. Calls EnumerateAccountProfiles at -PartnerId itself (captures its own profiles +
    #            everything inherited from ancestors, in one call)
    #         2. Walks every DESCENDANT partner under the resolved partner (EnumeratePartners, recursive) and
    #            calls EnumerateAccountProfiles at each one (captures profiles owned deeper down;
    #            re-returned inherited profiles are simply de-duplicated)
    #         3. De-duplicates the combined results by the profile's unique Id
    #       Reports: ProfileId, ProfileName, ProfileVersion, OwningPartnerId, OwningPartnerName,
    #       DeviceCount (cross-referenced from the DeviceProfileReport sheet's ProfileId column -
    #       no extra API call needed), Local Speed Vault settings (Mode/Location/UserName, from
    #       ProfileData.LocalSpeedVaultSettings) when configured, and one column PER KNOWN DATA
    #       SOURCE (fixed list, so every row has the same columns). Each data-source column folds
    #       in Selection paths, backup frequency (ProfileData.HighFrequentBackupSchedule - the
    #       simple Every1Hour/Daily/etc. cadence, NOT the separate per-day-of-week/time-window
    #       "individual schedule" collection), Policy/SelectionModification settings, and the
    #       exclusion filter list (one wildcard per line - it is itself pipe-delimited in the raw
    #       API response).
    #
    # NOTE on LSVStatus (YV) / StorageStatus (YS): these are raw numeric codes returned by the Cove
    # API (ValueType=IntNumber). No public enum reference for their meaning was found in the Cove
    # developer docs at time of writing - the raw values are exported as-is.
    #
    # Use the -PartnerName parameter to set the target partner; if omitted, the authenticated partner is inspected
    # Restricted Root/Sub-root/Distributor users are prompted for a customer/partner name
    # Partner resolution uses GetPartnerTree + GetPartnerInfoById; deprecated GetPartnerInfo is not used
    # v05 restores the standard machine/user API credential file instead of the HEH MCP credential file
    # Use the -MaxDevices ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -ExportPath (?:\Folder) parameter to specify export file path
    # Use the -Launch switch parameter to open the .xlsx after completion (default: on; use -Launch:$false to suppress)
    # Use the -ClearCredentials switch parameter to remove stored API credentials at start of script
    # Automatically refresh the API visa when it is more than 10 minutes old
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)] [string]$PartnerName,                             ## Override root partner name
    [Parameter(Mandatory=$False)] [int]$MaxDevices        = 5000,                    ## Maximum number of devices to return
    [Parameter(Mandatory=$False)] [string]$ExportPath     = "$PSScriptRoot\Output",  ## Export Path
    [Parameter(Mandatory=$False)] [switch]$Launch          = $true,                  ## Open .xlsx file when done (use -Launch:$false to suppress)
    [Parameter(Mandatory=$False)] [switch]$ClearCredentials                         ## Remove stored API credentials at start of script
)

#region ----- Environment ----
    $ConsoleTitle = "Combined Device Profile + Profile Summary Report"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-Host "`n  $ConsoleTitle`n" -ForegroundColor Cyan

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $CurrentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $urlJSON     = "https://api.backup.management/jsonapi"
    $Script:strLineSeparator = "  ---------"
#endregion ----- Environment ----

#region ----- Authentication (standard repo pattern) ----
    $Script:CredentialFile = "C:\ProgramData\MXB\$($env:COMPUTERNAME)_$($env:USERNAME)_API_Credentials.xml"

    Function ConvertFrom-SecureString2 {
        param([System.Security.SecureString]$SecureString)
        $ptr  = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
        $text = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)
        return $text
    }

    Function Get-APICredentials {
        $credFile = $Script:CredentialFile

        if ($ClearCredentials -and (Test-Path $credFile)) {
            Remove-Item $credFile -Force
            Write-Host "  Credential file removed. Re-prompting..." -ForegroundColor Yellow
        }

        if (Test-Path $credFile) {
            Write-Host "  Loading credentials from $credFile" -ForegroundColor Cyan
            $stored = Import-Clixml -Path $credFile
            if ($stored.PSObject.Properties['PartnerName']) {
                $secPwd = $stored.Password | ConvertTo-SecureString
                $Script:cred0 = $stored.PartnerName
                $Script:cred1 = $stored.Username
                $Script:cred2 = ConvertFrom-SecureString2 -SecureString $secPwd
                Write-Host "  Partner  : $Script:cred0" -ForegroundColor DarkCyan
                Write-Host "  API User : $Script:cred1" -ForegroundColor DarkCyan
            } else {
                $Script:cred0 = ''
                $Script:cred1 = $stored.UserName
                $Script:cred2 = ConvertFrom-SecureString2 -SecureString $stored.Password
                Write-Host "  API User : $Script:cred1  (legacy format)" -ForegroundColor DarkCyan
            }
        } elseif ($env:COVE_USERNAME -and $env:COVE_PASSWORD) {
            Write-Host "  Using COVE_USERNAME / COVE_PASSWORD environment variables" -ForegroundColor Cyan
            $Script:cred0 = ''
            $Script:cred1 = $env:COVE_USERNAME
            $Script:cred2 = $env:COVE_PASSWORD
        } else {
            Write-Host "  No credential file found. Creating..." -ForegroundColor Yellow
            $partnerName = ''
            do { $partnerName = Read-Host "  Enter Backup.Management Partner Name (e.g. 'Acme, Inc')" }
            while ($partnerName.Length -eq 0)

            $cred = Get-Credential -Message 'Enter Backup.Management API User (email / username) and API Token (password)'
            if (-not $cred) { throw "No credentials provided." }

            $dir = Split-Path $credFile -Parent
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

            [PSCustomObject]@{
                PartnerName = $partnerName
                Username    = $cred.UserName
                Password    = ($cred.Password | ConvertFrom-SecureString)
            } | Export-Clixml -Path $credFile -Force

            Write-Host "  Credentials saved to $credFile" -ForegroundColor Green
            $Script:cred0 = $partnerName
            $Script:cred1 = $cred.UserName
            $Script:cred2 = ConvertFrom-SecureString2 -SecureString $cred.Password
        }
    }

    Function Send-APICredentialsCookie {
        Get-APICredentials

        $body = @{ jsonrpc='2.0'; id='2'; method='Login'
                   params=@{ username=$Script:cred1; password=$Script:cred2 } } | ConvertTo-Json

        $maxRetries = 4
        $resp = $null
        $lastError = $null
        for ($attempt = 0; $attempt -le $maxRetries; $attempt++) {
            try {
                $resp = Invoke-RestMethod -Uri $urlJSON -Method POST -ContentType 'application/json' -Body $body -TimeoutSec 120 -ErrorAction Stop
                break
            } catch {
                $lastError = $_
                $canRetry = $attempt -lt $maxRetries -and (Test-CoveTransientError -ErrorRecord $_)
                if (-not $canRetry) { break }

                $delayMs = ([Math]::Min(30000, [int](2000 * [Math]::Pow(2, $attempt)))) + (Get-Random -Minimum 0 -Maximum 1000)
                Write-Host "  [RETRY] Login transport failed; retrying in $([Math]::Round($delayMs / 1000, 1)) seconds ($($attempt + 1)/$maxRetries)" -ForegroundColor Yellow
                Start-Sleep -Milliseconds $delayMs
            }
        }

        if ($lastError -and -not $resp) {
            Write-Host "  Login request failed after $maxRetries retries: $($lastError.Exception.Message)" -ForegroundColor Red
            Write-Host "  This appears to be a network/API transport failure, not necessarily invalid credentials." -ForegroundColor Yellow
            exit 1
        }

        if ($resp.visa) {
            $Script:visa           = $resp.visa
            $Script:UserId         = $resp.result.result.id
            $Script:AuthPartnerId  = [int]$resp.result.result.PartnerId
            Write-Host "  Authenticated as: $Script:cred1" -ForegroundColor Green
            Write-Host "  Authenticated PartnerId (default search scope): $Script:AuthPartnerId" -ForegroundColor DarkCyan
        } else {
            Write-Host "  Authentication rejected by Cove - use -ClearCredentials to re-prompt" -ForegroundColor Red
            exit 1
        }
    }

    Function Convert-UnixTimeToDateTime {
        param([long]$InputUnixTime)

        if ($InputUnixTime -gt 0) {
            return [DateTimeOffset]::FromUnixTimeSeconds($InputUnixTime).UtcDateTime
        }

        return [DateTime]::MinValue
    }

    Function Get-VisaTime {
        if ($Script:visa) {
            $visaTime = Convert-UnixTimeToDateTime ([int]$Script:visa.Split('-')[3])
            if ($visaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)) {
                Write-Host "  API visa is older than 10 minutes. Refreshing authentication..." -ForegroundColor Yellow
                Send-APICredentialsCookie
            }
        }
    }
#endregion ----- Authentication ----

#region ----- Cove API ----
    Function Get-CovePartnerInfoById {
        param([int]$PartnerId)

        $resp = Invoke-CoveAPI -Method 'GetPartnerInfoById' -Params @{ partnerId = $PartnerId }
        if ($resp.error) { throw "GetPartnerInfoById($PartnerId) failed: $($resp.error.message)" }

        $info = $resp.result
        if ($info -and $info.PSObject.Properties['result']) { $info = $info.result }
        if (-not $info) { throw "GetPartnerInfoById($PartnerId) returned no partner details." }

        return [PSCustomObject]@{
            Uid   = [string]$info.Uid
            Id    = [int]$info.Id
            Name  = [string]$info.Name
            Level = [string]$info.Level
        }
    }

    Function Get-CovePartnerInfo {
        ## Combined replacement for deprecated GetPartnerInfo(name=...).
        ## Uses authenticated partner scope by default, prompts only for restricted users,
        ## searches descendants with GetPartnerTree, then gets full details by ID.
        param(
            [string]$RequestedPartnerName,
            [int]$SearchPartnerId = $Script:AuthPartnerId
        )

        $restrictedLevels = @('Root', 'Sub-root', 'Distributor')

        if (-not $RequestedPartnerName) {
            $self = Get-CovePartnerInfoById -PartnerId $SearchPartnerId
            if ($self.Level -notin $restrictedLevels) { return $self }

            do {
                $RequestedPartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'"
            } while ([string]::IsNullOrWhiteSpace($RequestedPartnerName))
        }

        do {
            Write-Host "  Searching partner tree for '$RequestedPartnerName' (scope $SearchPartnerId; API timeout 120 seconds)..." -ForegroundColor Cyan
            Write-Progress -Id 3 -Activity "Partner lookup" -Status "Searching GetPartnerTree..." -PercentComplete 15
            $treeResp = Invoke-CoveAPI -Method 'GetPartnerTree' -Params @{
                partnerId     = $SearchPartnerId
                fields        = @(0,1,3,4)
                childrenLimit = 5000
                partnerFilter = @{ NamePattern = [string]$RequestedPartnerName }
            }
            Write-Progress -Id 3 -Activity "Partner lookup" -Status "Processing partner tree results..." -PercentComplete 75

            if ($treeResp.error) {
                Write-Host "  GetPartnerTree failed: $($treeResp.error.message)" -ForegroundColor Yellow
                $partnerMatches = @()
            } else {
                $treeRoot = $treeResp.result
                if ($treeRoot -and $treeRoot.PSObject.Properties['result']) { $treeRoot = $treeRoot.result }

                $partnerMatches = [System.Collections.Generic.List[object]]::new()
                $partnerTreeInfoById = @{}
                function Add-PartnerTreeMatches {
                    param($Node, [string]$Needle)

                    if ($Node.Info -and $Node.Info.Id) {
                        $partnerTreeInfoById[[int]$Node.Info.Id] = $Node.Info
                    }

                    if ($Node.Info -and $Node.Info.Name -and ($Node.Info.Name -like "*$Needle*")) {
                        if ($Node.Info.Level -notin $restrictedLevels) {
                            $partnerMatches.Add([PSCustomObject]@{
                                Id       = ([int]$Node.Info.Id).ToString([System.Globalization.CultureInfo]::InvariantCulture)
                                Name     = [string]$Node.Info.Name
                                Level    = [string]$Node.Info.Level
                                ParentId = ([int]$Node.Info.ParentId).ToString([System.Globalization.CultureInfo]::InvariantCulture)
                                Path     = ''
                            }) | Out-Null
                        }
                    }

                    foreach ($child in $Node.Children) {
                        Add-PartnerTreeMatches -Node $child -Needle $Needle
                    }
                }

                if ($treeRoot) { Add-PartnerTreeMatches -Node $treeRoot -Needle $RequestedPartnerName }
                $partnerMatches = @($partnerMatches)

                foreach ($partnerMatch in $partnerMatches) {
                    $pathParts = [System.Collections.Generic.List[string]]::new()
                    $pathNode = $partnerTreeInfoById[[int]$partnerMatch.Id]
                    $pathGuard = [System.Collections.Generic.HashSet[int]]::new()
                    while ($pathNode -and $pathGuard.Add([int]$pathNode.Id)) {
                        $pathParts.Insert(0, [string]$pathNode.Name)
                        $parentId = [int]$pathNode.ParentId
                        $pathNode = $partnerTreeInfoById[$parentId]
                    }
                    $partnerMatch.Path = $pathParts -join ' > '
                }
            }

            if ($partnerMatches.Count -eq 0) {
                Write-Progress -Id 3 -Activity "Partner lookup" -Completed
                Write-Host "  No partner found matching '$RequestedPartnerName'." -ForegroundColor Yellow
                do {
                    $RequestedPartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup"
                } while ([string]::IsNullOrWhiteSpace($RequestedPartnerName))
                continue
            }

            if ($partnerMatches.Count -gt 1) {
                Write-Progress -Id 3 -Activity "Partner lookup" -Completed
                Write-Host "  At least $($partnerMatches.Count) partners match '$RequestedPartnerName'. Select one:" -ForegroundColor Cyan
                $selected = $partnerMatches |
                    Select-Object Id, Name, Level, ParentId, Path |
                    Out-GridView -Title "Select Partner matching '$RequestedPartnerName'" -OutputMode Single
                if (-not $selected) {
                    do {
                        $RequestedPartnerName = Read-Host "  No partner selected. Enter Customer/ Partner display name (or partial name) to lookup"
                    } while ([string]::IsNullOrWhiteSpace($RequestedPartnerName))
                    continue
                }
            } else {
                $selected = $partnerMatches[0]
            }

            Write-Progress -Id 3 -Activity "Partner lookup" -Completed
            return Get-CovePartnerInfoById -PartnerId ([int]$selected.Id)
        } while ($true)
    }

    Function Test-CoveTransientError {
        param([System.Management.Automation.ErrorRecord]$ErrorRecord)

        $statusCode = 0
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
            $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
        }

        if ($statusCode -eq 408 -or $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) {
            return $true
        }

        $message = $ErrorRecord.Exception.ToString()
        return $message -match '(?i)(timed out|timeout|connection.*(failed|closed|reset|refused)|SSL connection|did not properly respond|temporarily unavailable)'
    }

    Function Invoke-CoveAPI {
        param(
            [string]$Method,
            [hashtable]$Params,
            [int]$MaxRetries = 4
        )

        Get-VisaTime
        $payload = @{ jsonrpc='2.0'; id='1'; method=$Method; visa=$Script:visa; params=$Params } | ConvertTo-Json -Depth 10

        for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
            try {
                return Invoke-RestMethod -Uri $urlJSON -Method POST -ContentType 'application/json' -Body $payload -TimeoutSec 120 -ErrorAction Stop
            } catch {
                $lastError = $_
                $canRetry = $attempt -lt $MaxRetries -and (Test-CoveTransientError -ErrorRecord $_)
                if (-not $canRetry) { break }

                ## Exponential backoff with up to one second of jitter. This avoids repeatedly
                ## hitting the API at the same time when several requests fail together.
                $delayMs = ([Math]::Min(30000, [int](2000 * [Math]::Pow(2, $attempt)))) + (Get-Random -Minimum 0 -Maximum 1000)
                Write-Host "  [RETRY] $Method failed; retrying in $([Math]::Round($delayMs / 1000, 1)) seconds ($($attempt + 1)/$MaxRetries)" -ForegroundColor Yellow
                Start-Sleep -Milliseconds $delayMs
            }
        }

        if ($lastError) {
            return @{ error = @{ message = $lastError.Exception.Message }; result = $null }
        }
    }

    Function Get-DeviceProfileRows {
        param([int]$PartnerId, [int]$MaxDevices)

        Write-Host "  Enumerating devices for partner $PartnerId (incl. sub-partners)..." -ForegroundColor Cyan
        Write-Progress -Id 1 -Activity "Cove report" -Status "Enumerating devices..." -PercentComplete 5

        $resp = Invoke-CoveAPI -Method "EnumerateAccountStatistics" -Params @{
            query = @{
                PartnerId          = $PartnerId
                RecurseSubPartners = $true
                Columns            = @("AU","AN","AR","MN","OS","OT","OI","OP","OV","VE","YV","YS","TS")
                SelectionMode      = "Merged"
                StartRecordNumber  = 0
                RecordsCount       = $MaxDevices
                OrderBy            = "AR ASC"
            }
        }

        if ($resp.error) { throw "EnumerateAccountStatistics failed: $($resp.error.message)" }

        ## Double-wrapped result: result.result
        $raw = $resp.result
        if ($raw -and $raw.PSObject.Properties['result']) { $raw = $raw.result }
        if (-not $raw) { Write-Host "  No devices returned." -ForegroundColor Yellow; return @() }

        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $raw) {
            ## Settings deserializes as an array of single-key objects: [{"AU":"..."},{"AN":"..."},...]
            $flat = @{}
            foreach ($setting in $entry.Settings) {
                foreach ($key in $setting.PSObject.Properties.Name) { $flat[$key] = $setting.$key }
            }

            $deviceType = switch ([int]($flat['OT'] ?? 0)) { 1 { 'Workstation' } 2 { 'Server' } default { 'Other' } }
            $lastSeen   = if ($flat['TS']) { [DateTimeOffset]::FromUnixTimeSeconds([long]$flat['TS']).LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }

            $rows.Add([PSCustomObject]@{
                AccountId       = [int]($flat['AU'] ?? 0)
                DeviceName      = ($flat['AN'] ?? '').Trim()
                MachineName     = ($flat['MN'] ?? '').Trim()
                CompanyName     = ($flat['AR'] ?? '').Trim()
                OwningPartnerId = $entry.PartnerId
                OS              = ($flat['OS'] ?? '').Trim()
                DeviceType      = $deviceType
                ProfileId       = $flat['OI']
                ProfileName     = ($flat['OP'] ?? '').Trim()
                ProfileVersion  = $flat['OV']
                LSVEnabled      = $flat['VE']
                LSVStatus       = $flat['YV']
                StorageStatus   = $flat['YS']
                LastSeen        = $lastSeen
            }) | Out-Null
        }

        Write-Progress -Id 1 -Activity "Cove report" -Status "Device enumeration complete" -PercentComplete 20
        Write-Host "  Found $($rows.Count) device(s)." -ForegroundColor Green
        return $rows
    }

    Function Get-DescendantPartnerIds {
        param([int]$ParentPartnerId)

        Write-Host "  Enumerating sub-partners under $ParentPartnerId (recursive)..." -ForegroundColor Cyan
        Write-Progress -Id 1 -Activity "Cove report" -Status "Discovering sub-partners..." -PercentComplete 25
        $resp = Invoke-CoveAPI -Method "EnumeratePartners" -Params @{
            parentPartnerId = $ParentPartnerId
            fetchRecursively = $true
        }
        if ($resp.error) { throw "EnumeratePartners failed: $($resp.error.message)" }

        $raw = $resp.result
        if ($raw -and $raw.PSObject.Properties['result']) { $raw = $raw.result }
        if (-not $raw) { return @() }

        $ids = @($raw | ForEach-Object { [int]$_.Id })
        Write-Progress -Id 1 -Activity "Cove report" -Status "Sub-partner discovery complete" -PercentComplete 30
        Write-Host "  Found $($ids.Count) sub-partner(s)." -ForegroundColor Green
        return $ids
    }

    Function Get-PartnerNameMap {
        param([int[]]$PartnerIds)

        $map = @{}
        $total = $PartnerIds.Count
        $current = 0
        foreach ($partnerId in $PartnerIds) {
            $current++
            $percent = 60 + [int](($current / [Math]::Max(1, $total)) * 15)
            Write-Progress -Id 2 -Activity "Resolving partner names" `
                -Status "Partner $current of $total (ID $partnerId)" -PercentComplete $percent
            $resp = Invoke-CoveAPI -Method "GetPartnerInfoById" -Params @{ partnerId = $partnerId }
            if (-not $resp.error) {
                $info = $resp.result
                if ($info -and $info.PSObject.Properties['result']) { $info = $info.result }
                if ($info) { $map[$partnerId] = $info.Name }
            }
        }
        Write-Progress -Id 2 -Activity "Resolving partner names" -Completed
        return $map
    }

    Function Get-ProfilesAtPartner {
        param([int]$PartnerId)

        $resp = Invoke-CoveAPI -Method "EnumerateAccountProfiles" -Params @{ partnerId = $PartnerId }
        if ($resp.error) {
            Write-Host "  [WARN] EnumerateAccountProfiles($PartnerId) failed: $($resp.error.message)" -ForegroundColor Yellow
            return @()
        }
        $raw = $resp.result
        if ($raw -and $raw.PSObject.Properties['result']) { $raw = $raw.result }
        return @($raw)
    }

    ## Fixed, known data source list -> gives every profile row the SAME set of columns (required
    ## for a consistent worksheet) regardless of which data sources any individual profile configures.
    $Script:KnownDataSources = @(
        'WorkstationFileSystem','ServerFileSystem','SystemState','NetworkShares',
        'MsSql','Exchange','VMWare','SharePoint','Oracle','HyperV','MySql'
    )

    Function Format-DataSourceCell {
        ## One column per data source: folds selections, schedule, policy/mod settings, and
        ## exclusions (each pipe-delimited entry on its own line) into a single readable cell.
        param($DsSetting, $ScheduleByPlugin)

        if (-not $DsSetting) { return '' }

        $lines = [System.Collections.Generic.List[string]]::new()

        $selections = @($DsSetting.SelectionCollection | ForEach-Object { $_.Selection })
        if ($selections.Count -gt 0) { $lines.Add("Selection: $($selections -join '; ')") | Out-Null }

        $freq = $ScheduleByPlugin[$DsSetting.DataSource]
        if ($freq) { $lines.Add("Schedule: $freq") | Out-Null }

        if ($DsSetting.Policy -or $DsSetting.SelectionModification) {
            $lines.Add("Policy: $($DsSetting.Policy) ($($DsSetting.SelectionModification))") | Out-Null
        }

        if ($DsSetting.ExclusionFilter) {
            $lines.Add("Exclusions:") | Out-Null
            $lines.Add("| $($DsSetting.ExclusionFilter -replace '\|', "`r`n| ")") | Out-Null
        }

        return ($lines -join "`r`n")
    }

    Function Get-ProfileSummaryRows {
        param([int]$PartnerId, [hashtable]$DeviceCountsByProfileId)

        ## Root partner itself already returns its own profiles + everything inherited from ancestors
        $descendantIds  = Get-DescendantPartnerIds -ParentPartnerId $PartnerId
        $partnersToScan = @($PartnerId) + $descendantIds

        Write-Host "  Scanning $($partnersToScan.Count) partner node(s) for backup profiles..." -ForegroundColor Cyan

        ## Plain hashtable, not [ordered]: OrderedDictionary treats an int key as a POSITIONAL index
        ## on both get and set, throwing "Specified argument was out of range" for a profile Id that
        ## doesn't match an existing position. Final output order comes from Sort-Object at export time.
        $profileMap = @{}   ## key = profile Id (unique) -> profile object
        $totalPartners = $partnersToScan.Count
        $currentPartner = 0
        foreach ($partnerIdItem in $partnersToScan) {
            $currentPartner++
            $percent = 30 + [int](($currentPartner / [Math]::Max(1, $totalPartners)) * 30)
            Write-Progress -Id 1 -Activity "Scanning backup profiles" `
                -Status "Partner $currentPartner of $totalPartners (ID $partnerIdItem)" -PercentComplete $percent
            foreach ($profileItem in (Get-ProfilesAtPartner -PartnerId $partnerIdItem)) {
                if ($profileItem -and $profileItem.Id -and -not $profileMap.ContainsKey([int]$profileItem.Id)) {
                    $profileMap[[int]$profileItem.Id] = $profileItem
                }
            }
        }

        Write-Progress -Id 1 -Activity "Scanning backup profiles" -Status "Profile scan complete" -PercentComplete 60
        Write-Host "  Found $($profileMap.Count) unique profile(s) across the tree." -ForegroundColor Green
        if ($profileMap.Count -eq 0) {
            Write-Progress -Id 1 -Activity "Cove report" -Completed
            return @()
        }

        $owningPartnerIds = @($profileMap.Values | ForEach-Object { [int]$_.PartnerId } | Sort-Object -Unique)
        $partnerNames     = Get-PartnerNameMap -PartnerIds $owningPartnerIds

        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($profileItem in $profileMap.Values) {
            $pd = ($profileItem.ProfileData) ?? @{}

            $dsByName = @{}
            foreach ($ds in ($pd.BackupDataSourceSettings ?? @())) { $dsByName[$ds.DataSource] = $ds }

            ## HighFrequentBackupSchedule.BackupScheduleItems = one simple Frequency (e.g. Every1Hour,
            ## Every4Hours, Daily) per data source - the cadence schedule, NOT the separate
            ## BackupSchedule collection (per-day-of-week/time-window "individual schedule" entries).
            $scheduleByPlugin = @{}
            foreach ($item in ($pd.HighFrequentBackupSchedule.BackupScheduleItems ?? @())) {
                $scheduleByPlugin[$item.PluginId] = $item.Frequency
            }

            $lsv = $pd.LocalSpeedVaultSettings

            $row = [ordered]@{
                ProfileId         = [int]$profileItem.Id
                ProfileName       = $profileItem.Name
                ProfileVersion    = $profileItem.Version
                OwningPartnerId   = [int]$profileItem.PartnerId
                OwningPartnerName = $partnerNames[[int]$profileItem.PartnerId]
                DeviceCount       = ($DeviceCountsByProfileId[[int]$profileItem.Id] ?? 0)
                DataSourceCount   = $dsByName.Count
                LSVMode           = if ($lsv) { $lsv.LocalSpeedVaultMode } else { '' }
                LSVLocation       = if ($lsv) { $lsv.Location } else { '' }
                LSVUserName       = if ($lsv) { $lsv.UserName } else { '' }
            }
            foreach ($name in $Script:KnownDataSources) {
                $row[$name] = Format-DataSourceCell -DsSetting $dsByName[$name] -ScheduleByPlugin $scheduleByPlugin
            }

            $rows.Add([PSCustomObject]$row) | Out-Null
        }

        return @($rows | Sort-Object OwningPartnerName, ProfileName)
    }
#endregion ----- Cove API ----

#region ----- Excel Export ----
    Function ConvertTo-ExcelColumnWidthChars {
        ## Excel's Range.ColumnWidth is in character-width units (based on the workbook's Normal
        ## style font), NOT pixels - there is no direct "set width in pixels" COM property. Uses
        ## the standard ~7px-per-character + 5px padding approximation (accurate for the default
        ## Normal-style font; not pixel-exact across every font/DPI combination).
        param([double]$Pixels)
        return [Math]::Round(($Pixels - 5) / 7, 2)
    }

    ## ProfileSummary column widths, copied from the user's manually-formatted reference workbook
    ## (CombinedDeviceProfileReport_..._12-42-34.xlsx) so re-generated reports keep the same layout.
    ## Columns K-U (the wrapped data-source columns) use a fixed 260px width instead of the
    ## reference's bestFit values - AutoFit/bestFit does not work well for wrapped multi-line cells.
    $Script:ProfileSheetDataSourceColWidthChars = ConvertTo-ExcelColumnWidthChars -Pixels 260
    $Script:ProfileSheetColumnWidths = @(
        7.5703125, 36.5703125, 12.42578125, 14.7109375, 27.7109375, 11.28515625,
        12, 29.7109375, 25.140625, 15.7109375
    ) + (@($Script:ProfileSheetDataSourceColWidthChars) * 11)
    ## Columns K-U (11-21): the per-data-source cells hold multi-line text and need wrap text.
    ## Columns A-J (1-10): plain single-value cells - wrap text explicitly disabled.
    $Script:ProfileSheetWrapColumns = 11..21

    Function Add-DataSheetToWorkbook {
        param($Workbook, $Worksheet, [string]$SheetName, [array]$Rows, [double[]]$ColumnWidths, [switch]$BoldHeader, [int[]]$WrapTextColumns = @(), [double]$DataRowHeight = 0, [int[]]$AutoFitColumns = @())

        $safeName = ($SheetName -replace '[\\/\?\*\[\]:]', '_')
        if ($safeName.Length -gt 31) { $safeName = $safeName.Substring(0, 31) }
        $Worksheet.Name = $safeName

        if ($Rows.Count -eq 0) { return }

        $headers  = @($Rows[0].PSObject.Properties.Name)
        $rowCount = $Rows.Count
        $colCount = $headers.Count

        $data = New-Object 'object[,]' ($rowCount + 1), $colCount
        for ($c = 0; $c -lt $colCount; $c++) { $data[0,$c] = $headers[$c] }
        for ($r = 0; $r -lt $rowCount; $r++) {
            ## Precompute the row index: "$data[$r+1,$c] = ..." throws "does not contain a method
            ## named 'op_Addition'" - the assignment-form multi-dim indexer mishandles an inline
            ## arithmetic expression before the comma. Assigning to a variable first works fine.
            $rr = $r + 1
            for ($c = 0; $c -lt $colCount; $c++) {
                $val = $Rows[$r].($headers[$c])
                $data[$rr, $c] = if ($null -eq $val) { '' } else { [string]$val }
            }
        }

        $range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($rowCount+1, $colCount))
        $range.Value2 = $data
        $range.VerticalAlignment = -4160   ## xlTop - keeps wrapped multi-line cells top-aligned

        ## Wrap text only on the columns explicitly requested; everything else stays unwrapped.
        $wrapSet = [System.Collections.Generic.HashSet[int]]::new([int[]]$WrapTextColumns)
        for ($c = 1; $c -le $colCount; $c++) {
            $colRange = $Worksheet.Range($Worksheet.Cells.Item(1,$c), $Worksheet.Cells.Item($rowCount+1,$c))
            $colRange.WrapText = $wrapSet.Contains($c)
        }

        $headerRange = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item(1,$colCount))
        if ($BoldHeader) { $headerRange.Font.Bold = $true }
        $headerRange.AutoFilter() | Out-Null

        if ($ColumnWidths -and $ColumnWidths.Count -eq $colCount) {
            ## Excel's ColumnWidth (character-width units) hard-caps at 255 - throws
            ## "Unable to set the ColumnWidth property" above that, so clamp requested values.
            ## Explicit [double] cast avoids "Unable to cast ... Int32 ... to ... Double" from
            ## [Math]::Min() overload resolution when $ColumnWidths mixes int/double literals.
            for ($c = 1; $c -le $colCount; $c++) { $Worksheet.Columns.Item($c).ColumnWidth = [Math]::Min(255.0, [double]$ColumnWidths[$c-1]) }
        } else {
            $Worksheet.Columns.AutoFit() | Out-Null
        }

        ## Explicitly requested columns override the fixed width above with AutoFit instead.
        foreach ($c in $AutoFitColumns) { $Worksheet.Columns.Item($c).AutoFit() | Out-Null }

        if ($DataRowHeight -gt 0) {
            $Worksheet.Range($Worksheet.Cells.Item(2,1), $Worksheet.Cells.Item($rowCount+1,1)).EntireRow.RowHeight = $DataRowHeight
        }

        $Worksheet.Application.ActiveWindow.SplitRow = 1
        $Worksheet.Application.ActiveWindow.FreezePanes = $true
    }

    Function Set-ProfileSheetFonts {
        ## Matches the reference workbook's font sizes: main columns 10pt, LSV columns 10pt bold,
        ## DataSourceCount + per-data-source columns 9pt (smaller, to offset the wrapped text height).
        param($Worksheet, [int]$RowCount, [int]$ColCount)

        $lastRow = $RowCount + 1
        $Worksheet.Range($Worksheet.Cells.Item(2,1), $Worksheet.Cells.Item($lastRow,6)).Font.Size  = 10
        $Worksheet.Range($Worksheet.Cells.Item(2,7), $Worksheet.Cells.Item($lastRow,9)).Font.Size  = 10
        $Worksheet.Range($Worksheet.Cells.Item(2,7), $Worksheet.Cells.Item($lastRow,9)).Font.Bold   = $true
        $Worksheet.Range($Worksheet.Cells.Item(2,10), $Worksheet.Cells.Item($lastRow,$ColCount)).Font.Size = 9
    }

    Function Export-CombinedWorkbook {
        param([array]$DeviceRows, [array]$ProfileRows, [string]$OutputFile)

        Write-Host "  Writing workbook (Excel COM)..." -ForegroundColor Cyan
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible       = $false
        $xl.DisplayAlerts = $false
        try {
            $wb = $xl.Workbooks.Add()
            while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

            $wsDevices = $wb.Worksheets.Item(1)
            Add-DataSheetToWorkbook -Workbook $wb -Worksheet $wsDevices -SheetName "DeviceProfileReport" -Rows $DeviceRows

            $wsProfiles = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
            Add-DataSheetToWorkbook -Workbook $wb -Worksheet $wsProfiles -SheetName "ProfileSummary" -Rows $ProfileRows `
                -ColumnWidths $Script:ProfileSheetColumnWidths -BoldHeader -WrapTextColumns $Script:ProfileSheetWrapColumns -DataRowHeight 75 -AutoFitColumns (6..10)
            if ($ProfileRows.Count -gt 0) {
                Set-ProfileSheetFonts -Worksheet $wsProfiles -RowCount $ProfileRows.Count -ColCount $ProfileRows[0].PSObject.Properties.Name.Count
            }

            $wsDevices.Activate()
            $xlOpenXMLWorkbook = 51
            $wb.SaveAs($OutputFile, $xlOpenXMLWorkbook)
            $wb.Close($false)
        } finally {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
#endregion ----- Excel Export ----

#region ----- Main ----
    Send-APICredentialsCookie
    $targetPartner = Get-CovePartnerInfo -RequestedPartnerName $PartnerName
    $Script:PartnerId = [int]$targetPartner.Id
    $Script:PartnerName = $targetPartner.Name
    $Script:Uid = $targetPartner.Uid
    $Script:Level = $targetPartner.Level

    Write-Output $Script:strLineSeparator
    Write-Output "  $($Script:Level) - $($Script:PartnerName) - $($Script:PartnerId) - $($Script:Uid)"
    Write-Output $Script:strLineSeparator

    $deviceRows = Get-DeviceProfileRows -PartnerId $Script:PartnerId -MaxDevices $MaxDevices

    ## Cross-reference DeviceCount per profile directly from the device rows just fetched -
    ## avoids a second EnumerateAccountStatistics call just to count devices per profile.
    $deviceCountsByProfileId = @{}
    foreach ($d in $deviceRows) {
        if ($d.ProfileId) {
            $key = [int]$d.ProfileId
            $deviceCountsByProfileId[$key] = ($deviceCountsByProfileId[$key] ?? 0) + 1
        }
    }

    $profileRows = Get-ProfileSummaryRows -PartnerId $Script:PartnerId -DeviceCountsByProfileId $deviceCountsByProfileId
    Write-Progress -Id 1 -Activity "Cove report" -Completed

    if ($deviceRows.Count -eq 0 -and $profileRows.Count -eq 0) {
        Write-Host "  Nothing to export. Exiting." -ForegroundColor Yellow
        exit 0
    }

    $deviceRowsSorted  = @($deviceRows | Sort-Object CompanyName, MachineName)

    if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null }
    $safeLabel  = ($Script:PartnerName -replace '[^\w-]', '-')
    $outputFile = Join-Path $ExportPath "CombinedDeviceProfileReport_${safeLabel}_$($Script:PartnerId)_$CurrentDate.xlsx"

    Export-CombinedWorkbook -DeviceRows $deviceRowsSorted -ProfileRows $profileRows -OutputFile $outputFile

    Write-Host "`n  Report saved: $outputFile" -ForegroundColor Green
    if ($Launch) { Start-Process $outputFile }
#endregion ----- Main ----
