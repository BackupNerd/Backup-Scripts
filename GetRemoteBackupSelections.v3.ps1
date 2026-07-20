<# ----- About: ----
    # Cove Data Protection - Get Remote Device Backup Selections
    # Revision v3 - 2026-07-20
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
    # GitHub: @BackupNerd  https://github.com/BackupNerd
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
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices and metadata via EnumerateAccountStatistics
    # Maintain a persistent master CSV of all known devices across runs
    # For each device connect via RCG and call EnumerateBackupSelections + EnumerateBackupSchedule
    # Merge live selections and schedules into master CSV; preserve data for unreachable devices
    # Export full master in analyst-friendly format to a dated Output subfolder (CSV + XLS)
    # Flag anomalies: UNREACHABLE, ORPHANED, PROFILE, EXCLUSIONS, SPECIFIC_FS, ORPHANED_DS
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -AllDevices switch parameter to skip GUI device selection
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -Export switch parameter to export results to CSV/XLS files
    # Use the -ExportPath (?:\Folder) parameter to specify export file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the CSV delimiter
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    # Use the -DeviceThrottle ## (default=40) parameter to set max parallel RCG threads
    # Use the -RetryFailed switch to retry unreachable devices after the first pass
    # Use the -RetryOnly switch to skip the main pass and only retry previously unreachable devices
    # Use the -ActiveWithinDays ## (default=7) parameter to filter devices by last heartbeat
    # Use the -FilterAccountIDs parameter to process only specific device AccountIDs
    # Use the -ExcludeColumns parameter to omit datasource columns from export (e.g. -ExcludeColumns Exch,SP,ORC)
    # Use the -DebugCDP switch to enable verbose debug output
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [string]$PartnerName  = "",                       ## Override root partner name (default: partner from stored credentials)
        [Parameter(Mandatory=$False)] [switch]$AllPartners  = $true,                    ## Skip partner selection GUI
        [Parameter(Mandatory=$False)] [switch]$AllDevices   = $true,                    ## Skip device selection GUI
        [Parameter(Mandatory=$False)] [int]$DeviceCount     = 5000,                     ## Maximum number of devices to return
        [Parameter(Mandatory=$False)] [switch]$Export       = $true,                    ## Generate CSV / XLS Output Files
        [Parameter(Mandatory=$False)] [switch]$Launch       = $true,                    ## Launch XLS or CSV file
        [Parameter(Mandatory=$False)] [string]$Delimiter    = ',',                      ## Delimiter for CSV file
        [Parameter(Mandatory=$False)] $ExportPath           = "$PSScriptRoot",          ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials,                        ## Remove Stored API Credentials at start of script

        [Parameter(Mandatory=$False)] [int]$DeviceThrottle  = 20,                       ## Max parallel device threads (PS7+) — keep low to avoid overloading RCG relay servers
        [Parameter(Mandatory=$False)] [switch]$RetryFailed,                    ## Retry unreachable devices once after first pass
        [Parameter(Mandatory=$False)] [int]$ActiveWithinDays = 7,                       ## Only process devices with heartbeat within N days (0 = all devices)
        [Parameter(Mandatory=$False)] [switch]$RetryOnly,                              ## Skip main pass; only retry devices marked ReachableStatus=No in master CSV
        [Parameter(Mandatory=$False)] [switch]$DebugCDP,                                ## Show debug output (column counts, System State merge details, etc.)
        [Parameter(Mandatory=$False)] [int[]]$FilterAccountIDs = @(),                   ## If set, only process devices with these AccountIDs
        [Parameter(Mandatory=$False)]
        [ValidateSet('FS','SS','HypV','SQL','MySQ','Net','VM','SP','ORC','Exch')]
        [string[]]$ExcludeColumns = @('VM','SP','ORC','Exch')                           ## Datasource columns to omit from export (e.g. 'Exch','SP')
    )

    ## Translate short ExcludeColumns aliases to internal abbreviations used as column prefixes
    $Script:ExcludeColAliasMap = @{
        'FS'   = 'FS';       'SS'   = 'SysState'; 'HypV' = 'HyperV'
        'SQL'  = 'MSSQL';    'MySQ' = 'MySQL';     'Net'  = 'NetShares'
        'VM'   = 'VMware';   'SP'   = 'SharePt';   'ORC'  = 'Oracle'
        'Exch' = 'Exchange'
    }
    $Script:ExcludeColumnSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]($ExcludeColumns | ForEach-Object { $Script:ExcludeColAliasMap[$_] }),
        [System.StringComparer]::OrdinalIgnoreCase
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Remote Device Backup Selections via RCG"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

    Write-Output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"

    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $Script:DebugCDP = $DebugCDP
    $urlJSON = "https://api.backup.management/jsonapi"

    $Script:SelectionResults = [System.Collections.Generic.List[object]]::new()  ## one row per device (pivoted) — for CSV/XLS export
    $Script:SelectionDetails  = [System.Collections.Generic.List[object]]::new()  ## one row per selection entry — for on-screen summary

    ## Plugin ID → column abbreviation, ordered by popularity
    ## Plugins sharing the same abbreviation have their results merged into one column
    ## Order: FS > SysState > HyperV > MSSQL > MySQL > NetShares > VMware > SharePt > Oracle > Exchange
    $Script:DataSources = [ordered]@{
        'FileSystem'       = 'FS'
        'VssSystemState'   = 'SysState'    ## D07 — Windows VSS System State
        'SystemState'      = 'SysState'    ## D02 — macOS System State
        'LinuxSystemState' = 'SysState'    ## Linux System State — same output column as all other SysState variants
        'VssHyperV'        = 'HyperV'
        'VssMsSql'         = 'MSSQL'
        'MySql'            = 'MySQL'
        'NetworkShares'    = 'NetShares'
        'VMWare'           = 'VMware'      ## schema name is VMWare (capital M)
        'VssSharePoint'    = 'SharePt'
        'Oracle'           = 'Oracle'
        'VssExchange'      = 'Exchange'    ## schema name is VssExchange, not Exchange
    }

    ## Maps DataSource schema names to the BackupPlugin class names used in EnumerateBackupSchedule Plugins field
    $Script:PluginClassMap = @{
        'FileSystem'       = 'FsBackupPlugin'
        'VssSystemState'   = 'SystemStateBackupPlugin'
        'SystemState'      = 'SystemStateBackupPlugin'
        'LinuxSystemState' = 'LinuxSystemStateBackupPlugin'
        'VssHyperV'        = 'VssHyperVBackupPlugin'
        'VssMsSql'         = 'VssMsSqlBackupPlugin'
        'MySql'            = 'MySqlBackupPlugin'
        'NetworkShares'    = 'NetworkSharesBackupPlugin'
        'VMWare'           = 'VmwareBackupPlugin'
        'VssSharePoint'    = 'VssSharePointBackupPlugin'
        'Oracle'           = 'OracleBackupPlugin'
        'VssExchange'      = 'VssExchangeBackupPlugin'
    }

    ## AP column (I78) legacy single-char codes → plugin name
    ## Source: https://developer.n-able.com/n-able-cove/docs/column-codes
    $Script:ApCharToPlugin = @{
        'F' = 'FileSystem'    ## D01 Files and Folders
        'S' = 'SystemState'   ## D02 System State
        'Z' = 'VssMsSql'      ## D10 VSS MSSQL
        'H' = 'VssHyperV'     ## D14 Hyper-V
        'N' = 'NetworkShares' ## D06 Network Shares
        'L' = 'MySql'         ## D15 MySQL
        'X' = 'VssExchange'   ## D04 VSS Exchange (schema: VssExchange)
        'W' = 'VMWare'        ## D08 VMware (schema: VMWare)
        'P' = 'VssSharePoint' ## D11 VSS SharePoint
        'Y' = 'Oracle'        ## D12 Oracle
    }

    ## I78 new-format 3-char D-codes → plugin name
    $Script:ApDCodeToPlugin = @{
        'D01' = 'FileSystem';      'D02' = 'SystemState';      'D04' = 'VssExchange'   ## schema: VssExchange
        'D06' = 'NetworkShares';   'D07' = 'VssSystemState';   'D08' = 'VMWare'        ## schema: VMWare
        'D10' = 'VssMsSql';        'D11' = 'VssSharePoint';    'D12' = 'Oracle'
        'D14' = 'VssHyperV';       'D15' = 'MySql'
    }

    Write-Output "  Current Parameters:"
    Write-Output "  -AllPartners  = $AllPartners"
    Write-Output "  -AllDevices   = $AllDevices"
    Write-Output "  -DeviceCount  = $DeviceCount"
    Write-Output "  -Export       = $Export"
    Write-Output "  -ExportPath   = $ExportPath"
    Write-Output "  -Launch       = $Launch"
    Write-Output "  -Delimiter    = $Delimiter"
    Write-Output "  -ActiveWithinDays = $ActiveWithinDays  $(if ($ActiveWithinDays -gt 0) { "(devices not seen in $ActiveWithinDays+ days will be skipped)" } else { '(all devices regardless of last heartbeat)' })"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {
        Write-Output $Script:strLineSeparator
        Write-Output "  Setting Backup API Credentials"
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath }

        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Data Protection | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($PartnerName.length -eq 0)
        $PartnerName | Out-File $APIcredfile

        $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove Data Protection API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName"

        $BackupCred.UserName | Out-File -Append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-File -Append $APIcredfile

        Start-Sleep -Milliseconds 300

        Send-APICredentialsCookie  ## Attempt API Authentication

    }  ## Set API credentials if not present

    Function Get-APICredentials {
        $Script:True_path = "C:\ProgramData\MXB\"
        $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
        $Script:APIcredpath = Split-Path -Path $APIcredfile

        if (($ClearCredentials) -and (Test-Path $APIcredfile)) {
            Remove-Item -Path $Script:APIcredfile
            $ClearCredentials = $Null
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Cleared"
            Send-APICredentialsCookie  ## Retry Authentication
        } else {
            Write-Output $Script:strLineSeparator
            Write-Output "  Getting Backup API Credentials"

            if (Test-Path $APIcredfile) {
                Write-Output $Script:strLineSeparator
                "  Backup API Credential File Present"
                $APIcredentials = Get-Content $APIcredfile

                $Script:cred0 = [string]$APIcredentials[0]
                $Script:cred1 = [string]$APIcredentials[1]
                $Script:cred2 = $APIcredentials[2] | ConvertTo-SecureString
                $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))

                Write-Output $Script:strLineSeparator
                Write-Output "  Stored Backup API Partner  = $Script:cred0"
                Write-Output "  Stored Backup API User     = $Script:cred1"
                Write-Output "  Stored Backup API Password = Encrypted"
            } else {
                Write-Output $Script:strLineSeparator
                Write-Output "  Backup API Credential File Not Present"
                Set-APICredentials  ## Create API Credential File if Not Found
            }
        }

    }  ## Get API credentials if present

    Function Send-APICredentialsCookie {
        Get-APICredentials  ## Read API Credential File before Authentication

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.method = 'Login'
        $data.params = @{}
        $data.params.partner   = $Script:cred0
        $data.params.username  = $Script:cred1
        $data.params.password  = $Script:cred2

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            $Script:cookies    = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Authenticate = $webrequest | ConvertFrom-Json

        if ($authenticate.visa) {
            $Script:visa   = $authenticate.visa
            $Script:UserId = $authenticate.result.result.id
        } else {
            Write-Output $Script:strLineSeparator
            Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output $Script:strLineSeparator
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Send-APICredentialsCookie
            }
        }
    }

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
            $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
            $epoch = $epoch.ToUniversalTime()
            $epoch = $epoch.AddSeconds($inputUnixTime)
            return $epoch
        }else{ return "" }
    }  ## Convert epoch time to date time

    Function Save-CSVasExcel {
        param (
            [string]$CSVFile = $(Throw 'No file provided.')
        )
        BEGIN {
            function Resolve-FullPath ([string]$Path) {
                if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) { $Path = "$PWD\$Path" }
                [IO.Path]::GetFullPath($Path)
            }
            function Release-Ref ($ref) {
                ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            $CSVFile = Resolve-FullPath $CSVFile
            $xl = New-Object -ComObject Excel.Application
        }
        PROCESS {
            $wb = $xl.workbooks.open($CSVFile)
            $xlOut = $CSVFile -replace '\.csv$', '.xlsx'
            $ws = $wb.Worksheets.Item(1)
            $range = $ws.UsedRange
            [void]$range.AutoFilter()
            [void]$range.EntireColumn.Autofit()
            $num = 1
            $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
            $nextname = $xlOut
            while (Test-Path $nextname) {
                $nextname = Join-Path (Split-Path $xlOut) $($base + "-$num" + '.xlsx')
                $num++
            }
            $wb.SaveAs($nextname, 51)
            $nextname  ## return actual saved path
        }
        END {
            $xl.Quit()
            $null = $ws, $wb, $xl | ForEach-Object { Release-Ref $_ }
        }
    }  ## Save as output XLS Routine

    Function Sort-CSVByAccount {
        param (
            [string]$CSVFile = $(Throw 'No file provided.')
        )
        ## Memory-efficient sort: read CSV, sort by AccountID then DeviceName, write back
        $rows = @(Import-Csv -Path $CSVFile)
        if ($rows.Count -gt 0) {
            $rows = $rows | Sort-Object { [int]$_.AccountID }, DeviceName
            $rows | Export-Csv -Path $CSVFile -Delimiter ',' -NoTypeInformation -Encoding UTF8 -Force
        }
    }  ## Sort CSV by AccountID and DeviceName

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function Send-GetPartnerInfo ($PartnerName) {
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'GetPartnerInfo'
        $data.params = @{}
        $data.params.name = [String]$PartnerName

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data -Depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            $Script:websession = $websession
            $Script:Partner = $webrequest | ConvertFrom-Json

        $RestrictedPartnerLevel = @("Root","Sub-root")

        if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$Script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-Output "  $Level - $PartnerName - $PartnerId - $Uid"
            Write-Output $Script:strLineSeparator
        } else {
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:PartnerName
        }

        if ($partner.error) {
            Write-Output "  $($partner.error.message)"
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:PartnerName
        }

    }  ## get PartnerID and Partner Level

    Function Send-EnumeratePartners {
        $objEnumeratePartners = (New-Object PSObject |
            Add-Member -PassThru NoteProperty jsonrpc '2.0' |
            Add-Member -PassThru NoteProperty visa $Script:visa |
            Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
            Add-Member -PassThru NoteProperty params @{
                parentPartnerId  = $PartnerId
                fetchRecursively = "true"
                fields           = (0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22)
            } |
            Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json -Depth 5

        [array]$Script:EnumeratePartnersSession = Invoke-RestMethod -Uri $urlJSON -Method POST `
            -ContentType 'application/json' -Body $objEnumeratePartners

        $Script:visa = $EnumeratePartnersSession.visa

        Start-Sleep -Milliseconds 100

        if ($EnumeratePartnersSession.error.code) {
            Write-Output $Script:strLineSeparator
            Write-Output "  EnumeratePartners Error: $($EnumeratePartnersSession.error.message)"
            Break
        }

        $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result |
            Select-Object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* `
            -ExcludeProperty Company -ErrorAction Ignore

        $Script:SelectedPartners = $Script:EnumeratePartnersSessionResults |
            Where-Object { $_.name -notlike "001???????????????- Recycle Bin" } |
            Where-Object { $_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????' }

        $Script:SelectedPartners += @( [PSCustomObject]@{ Name=$PartnerName; Id=[string]$PartnerId; Level='<ParentPartner>' } )

        if ($AllPartners) {
            $Script:Selection = $Script:SelectedPartners | Select-Object id,Name,Level | Sort-Object Level,Name
            Write-Output $Script:strLineSeparator
            Write-Output "  All Partners Selected"
        } else {
            $Script:Selection = $Script:SelectedPartners | Select-Object id,Name,Level | Sort-Object Level,Name |
                Out-GridView -Title "Select Partner | $PartnerName" -OutputMode Single

            if ($null -eq $Script:Selection) {
                Write-Output $Script:strLineSeparator
                Write-Output "  No Partner Selected"
                Break
            } else {
                [int]$Script:PartnerId = $Script:Selection.Id
                [String]$Script:PartnerName = $Script:Selection.Name
            }
        }

    }  ## EnumeratePartners API Call

    Function Send-GetDevices {
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        ## Base filter: Backup Manager agents only (I59==1)
        ## Add heartbeat cutoff if ActiveWithinDays > 0 — excludes devices whose RCG token has likely expired
        $tsFilter = if ($ActiveWithinDays -gt 0) {
            $cutoffTs = [int64]([datetime]::UtcNow.AddDays(-$ActiveWithinDays) - [datetime]'1970-01-01').TotalSeconds
            "I59 == 1 AND I6 > $cutoffTs"
        } else {
            "I59 == 1"
        }
        $data.params.query.Filter  = $tsFilter
        $data.params.query.Columns = @(
            ## Device identity
            "I0","I1","I2","I4","I6","I8","I9","I10","I11",
            ## Installation details
            "I16","I17","I18","I19","I20","I32","I44","I45",
            ## Profile / product
            "I54","I56",
            ## Feature / hardware
            "I78","I81","I84","I85",
            ## Overall last success
            "TL",
            ## Per-datasource last successful session (new shortcodes)
            "FL",   ## Files and Folders
            "SL",   ## System State (Windows)
            "KL",   ## Linux System State
            "XL",   ## Exchange (VSS)
            "NL",   ## Network Shares
            "WL",   ## VMware
            "ZL",   ## VSS MSSQL
            "PL",   ## VSS SharePoint
            "YL",   ## Oracle
            "HL",   ## Hyper-V
            "LL"    ## MySQL
        )
        $data.params.query.OrderBy = "I1 ASC"   ## Device name ascending
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)")

        $params = @{
            Uri         = $url
            Method      = 'POST'
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $data -Depth 6)))
            ContentType = 'application/json; charset=utf-8'
        }

        $Script:DeviceResponse = Invoke-RestMethod @params

        $Script:DeviceDetail = @()
        ForEach ($DeviceResult in $Script:DeviceResponse.result.result) {
            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{
                AccountID    = [Int]$DeviceResult.AccountId
                PartnerID    = [string]$DeviceResult.PartnerId
                DeviceName   = $DeviceResult.Settings.I1  -join ''  ## AN
                ComputerName = $DeviceResult.Settings.I18 -join ''  ## MN
                DeviceAlias  = $DeviceResult.Settings.I2  -join ''  ## AL
                PartnerName  = $DeviceResult.Settings.I8  -join ''  ## AR
                Location     = $DeviceResult.Settings.I11 -join ''  ## LN
                Product      = $DeviceResult.Settings.I10 -join ''  ## PN
                ProductID    = $DeviceResult.Settings.I9  -join ''  ## PD
                Profile      = $DeviceResult.Settings.I56 -join ''  ## OP
                ProfileID    = $DeviceResult.Settings.I54 -join ''  ## OI
                OS           = $DeviceResult.Settings.I16 -join ''  ## OS
                OSType       = $DeviceResult.Settings.I32 -join ''  ## OT
                ## Decode I78 to comma-separated plugin names ONCE here so all downstream code
                ## can simple comma-split without caring about D-code vs legacy-char format
                DataSources  = $(
                    $raw = ($DeviceResult.Settings.I78 | Where-Object { $_ }) -join ''
                    if ($raw -match '^D\d\d') {
                        ## New 3-char D-code format: "D01D02D10"
                        ([regex]::Matches($raw, 'D\d\d') | ForEach-Object { $Script:ApDCodeToPlugin[$_.Value] } | Where-Object { $_ } | Select-Object -Unique) -join ','
                    } elseif ($raw) {
                        ## Old single-char format: "FSZ"
                        ($raw.ToCharArray() | ForEach-Object { $Script:ApCharToPlugin["$_"] } | Where-Object { $_ } | Select-Object -Unique) -join ','
                    } else { "" }
                )
                TimeStamp    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.I6  -join '')  ## TS
                LastSuccess  = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL  -join '')
                CreationDate = Convert-UnixTimeToDateTime ($DeviceResult.Settings.I4  -join '')  ## CD
                ClientVersion = ($DeviceResult.Settings.I17 -join '').Trim()  ## VN
                IPAddress    = $DeviceResult.Settings.I19 -join ''  ## IP (internal); I20 = external IP
                Manufacturer = $DeviceResult.Settings.I44 -join ''  ## MF
                Model        = $DeviceResult.Settings.I45 -join ''  ## MO
                Physicality  = $DeviceResult.Settings.I81 -join ''
                CPUCores     = $DeviceResult.Settings.I84 -join ''
                RAMBytes     = $DeviceResult.Settings.I85 -join ''
                FS_LastOK         = Convert-UnixTimeToDateTime ($DeviceResult.Settings.FL -join '')
                SysState_LastOK    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.SL -join '')  ## SL covers both D02 and D07
                LinuxSS_LastOK    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.KL -join '')
                Exchange_LastOK   = Convert-UnixTimeToDateTime ($DeviceResult.Settings.XL -join '')
                NetShares_LastOK  = Convert-UnixTimeToDateTime ($DeviceResult.Settings.NL -join '')
                VMware_LastOK     = Convert-UnixTimeToDateTime ($DeviceResult.Settings.WL -join '')
                MSSQL_LastOK      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.ZL -join '')
                SharePt_LastOK    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.PL -join '')
                Oracle_LastOK     = Convert-UnixTimeToDateTime ($DeviceResult.Settings.YL -join '')
                HyperV_LastOK     = Convert-UnixTimeToDateTime ($DeviceResult.Settings.HL -join '')
                MySQL_LastOK      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.LL -join '')
            }
        }

    }  ## EnumerateAccountStatistics API Call to Get Devices

    Function Get-RcgSelections ($AccountId, $DeviceName, $DeviceDataSources, $Visa = $null, $DsMap = $null, $ApMap = $null) {
        ## $Visa / $DsMap / $ApMap allow self-contained execution inside parallel runspaces via $using:
        if (-not $Visa)  { $Visa  = $Script:visa }
        if (-not $DsMap) { $DsMap = $Script:DataSources }
        if (-not $ApMap) { $ApMap = $Script:ApCharToPlugin }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        # Stage 1: Get RCG endpoint — EnumerateAccountRemoteAccessEndpoints
        $rcgPayload = @{
            jsonrpc = '2.0'; id = 'jsonrpc'; visa = $Visa
            method  = 'EnumerateAccountRemoteAccessEndpoints'
            params  = @{ accountId = $AccountId }
        }
        $rcgResp = Invoke-RestMethod -Uri "https://api.backup.management/jsonapi" -Method POST `
            -ContentType 'application/json' -Body ($rcgPayload | ConvertTo-Json -Depth 5) `
            -TimeoutSec 30 -ErrorAction Stop
        $t1 = $sw.ElapsedMilliseconds

        if ($rcgResp.error) {
            Write-Host "  EnumerateAccountRemoteAccessEndpoints error: $($rcgResp.error.message)" -ForegroundColor Red
            return $null
        }

        # Result may be double-nested (result.result) or flat (result)
        $endpointsList = if ($rcgResp.result.result) { $rcgResp.result.result } else { $rcgResp.result }
        if (-not $endpointsList -or @($endpointsList).Count -eq 0) {
            Write-Host "  No RCG endpoint (device offline or unavailable)" -ForegroundColor Yellow
            return $null
        }

        $endpoint   = @($endpointsList)[0]
        $rcgUrl     = $endpoint.WebRcgUrl
        $rcgBaseUrl = $rcgUrl -replace '\?.*$', '' -replace '/$', ''
        $rcgRpcUrl  = "$rcgBaseUrl/jsonrpcv1"

        # Extract one-time auth_token from WebRcgUrl query param
        if (-not $rcgUrl) {
            Write-Host "  WebRcgUrl is empty — device has no RCG endpoint" -ForegroundColor Yellow
            return $null
        }
        if (-not ($rcgUrl -match "auth_token=([^&]+)")) {
            $urlPreview = $rcgUrl.Substring(0, [Math]::Min(100, $rcgUrl.Length))
            Write-Host "  No auth_token in WebRcgUrl: $urlPreview" -ForegroundColor Red
            return $null
        }
        $authToken = $matches[1]

        # Stage 2: GET /content/data/progress-status with auth_token to obtain RCG visa
        # Use -SessionVariable so the cookie is captured automatically regardless of path scope
        $rcgVisa = $null
        try {
            $getResp = Invoke-WebRequest -Uri "$rcgBaseUrl/content/data/progress-status?auth_token=$authToken" `
                -Method Get -UseBasicParsing -SessionVariable rcgSession -TimeoutSec 90 -ErrorAction Stop

            # Prefer session cookie (path-scoped cookies won't appear in Headers on some PS versions)
            $visCookie = $rcgSession.Cookies.GetCookies($rcgBaseUrl) | Where-Object { $_.Name -eq 'visa' }
            if ($visCookie) {
                $rcgVisa = $visCookie.Value
            } else {
                # Fallback: parse Set-Cookie header
                foreach ($cookieVal in @($getResp.Headers["Set-Cookie"])) {
                    if ($cookieVal -match "visa=([^;]+)") { $rcgVisa = $matches[1].Trim(); break }
                }
            }

            if (-not $rcgVisa) {
                Write-Debug "  No visa in Set-Cookie — auth may have failed (token already used?)"
                return $null
            }
        } catch {
            Write-Host "  RCG auth GET failed: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
        $t2 = $sw.ElapsedMilliseconds

        # Stage 3: Fire EnumerateBackupSelections plugin calls, filtered by I78 (ActiveDS)
        # 'Undefined' returns null via RCG — must call per-plugin; parallelize to avoid serial round-trips
        # Visa goes in Cookie header only — NOT in JSON body
        # $using: captures local vars into each parallel runspace
        
        ## Build list of plugins to query from $DeviceDataSources (I78 decoding)
        ## Special case: if ANY System State variant is enabled, query all three to find which has data
        $dsPluginsToQuery = @()
        if ([string]::IsNullOrWhiteSpace($DeviceDataSources)) {
            ## Fallback: if I78 is empty/missing, query all datasources (backward compat)
            $dsPluginsToQuery = [array]$DsMap.Keys
        } else {
            $enabledPlugins = $DeviceDataSources -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            $dsPluginsToQuery = @($enabledPlugins)
            
            ## Special handling for System State: if any variant (VssSystemState, SystemState, LinuxSystemState)
            ## is enabled, query all three and keep results with actual selections
            $sysStateVariants = @('VssSystemState', 'SystemState', 'LinuxSystemState')
            $hasSysState = $enabledPlugins | Where-Object { $_ -in $sysStateVariants }
            if ($hasSysState) {
                ## Add all System State variants (will filter empty results later)
                $dsPluginsToQuery += @($sysStateVariants | Where-Object { $_ -notin $dsPluginsToQuery })
            }
        }
        
        $allSelections = $dsPluginsToQuery | ForEach-Object -Parallel {
            $plugin   = $_
            $rpcUrl   = $using:rcgRpcUrl
            $cookie   = "visa=$($using:rcgVisa)"
            $headers  = @{ "Content-Type"="application/json"; "Cookie"=$cookie; "Cache-Control"="no-cache" }

            ## Same method for all plugins — identical to local GetLocalBackupSelections.v2 approach
            $payload = @{
                id      = 1; jsonrpc = "2.0"
                method  = "EnumerateBackupSelections"
                params  = @{ plugin = $plugin }
            } | ConvertTo-Json -Depth 4

            ## Retry logic for individual datasource failures (transient network errors, RCG timeouts)
            $maxRetries = 3
            $retryDelay = 100   ## milliseconds
            $attempt    = 0
            $lastError  = $null
            $result     = $null

            while ($attempt -lt $maxRetries -and $null -eq $result) {
                $attempt++
                try {
                    $resp = Invoke-RestMethod -Uri $rpcUrl -Method POST -Headers $headers `
                        -TimeoutSec 15 -Body $payload -ErrorAction Stop
                    $result = $resp.result.result   ## emit rows (null emits nothing)
                } catch {
                    $lastError = $_
                    if ($attempt -lt $maxRetries) {
                        Start-Sleep -Milliseconds $retryDelay
                    }
                }
            }

            ## Emit result (null results naturally filtered by Where-Object)
            $result
        } -ThrottleLimit 10 | Where-Object { $_ }

        ## Stage 3b: EnumerateBackupSchedule + GetHighFrequentBackupSchedule — same RCG visa, no additional auth
        $rawSchedules     = $null
        $rawHFSchedules   = $null
        $schedHeaders = @{ "Content-Type"="application/json"; "Cookie"="visa=$rcgVisa"; "Cache-Control"="no-cache" }
        try {
            $schedPayload = @{ id = 1; jsonrpc = "2.0"; method = "EnumerateBackupSchedule"; params = @{} } | ConvertTo-Json -Depth 3
            $rawSchedules = Invoke-RestMethod -Uri $rcgRpcUrl -Method POST -Headers $schedHeaders `
                -Body $schedPayload -TimeoutSec 15 -ErrorAction Stop
        } catch {
            $rawSchedules = @{ _error = $_.Exception.Message }
        }
        try {
            $hfPayload = @{ id = 1; jsonrpc = "2.0"; method = "GetHighFrequentBackupSchedule"; params = @{} } | ConvertTo-Json -Depth 3
            $rawHFSchedules = Invoke-RestMethod -Uri $rcgRpcUrl -Method POST -Headers $schedHeaders `
                -Body $hfPayload -TimeoutSec 15 -ErrorAction Stop
        } catch {
            $rawHFSchedules = @{ _error = $_.Exception.Message }
        }

        $sw.Stop()

        return @{
            Selections      = @($allSelections | Sort-Object PluginId)
            RawSchedules    = $rawSchedules
            RawHFSchedules  = $rawHFSchedules
            PerfMs       = [ordered]@{
                S1_Endpoint = $t1
                S2_RcgAuth  = $t2 - $t1
                S3_Plugins  = $sw.ElapsedMilliseconds - $t2
                Total       = $sw.ElapsedMilliseconds
            }
        }

    }  ## Get backup selections via RCG for a single device

#endregion ----- Backup.Management JSON Calls ----

#region ----- Master CSV Persistence ----

    Function Get-MasterCsvPath {
        ## Build master CSV filename from partner name and ID
        $sanitizedName = $Script:PartnerName -replace ' \(.*\)', '' -replace '[^a-zA-Z_0-9]', ''
        return "$ExportPath\RemoteSelections_${sanitizedName}_$($Script:PartnerId)_MASTER.csv"
    }

    Function Load-MasterCsv {
        ## Load existing master CSV into hashtable keyed by AccountID
        ## Returns hashtable: @{ AccountID => @{ DeviceName => "...", Selections => [...], ReachableStatus => "Yes/No", LastValidated => "...", ... } }
        $masterPath = Get-MasterCsvPath
        $masterHash = @{}

        ## Acquire exclusive file lock — blocks concurrent script instances from corrupting the master
        $lockPath = $masterPath + ".lock"
        try {
            $Script:MasterLockStream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        } catch {
            Write-Host "  [!] Master CSV is locked by another running instance of this script." -ForegroundColor Red
            Write-Host "      Lock file: $lockPath" -ForegroundColor Red
            Write-Host "      Wait for the other run to complete, or delete the lock file if it is stale." -ForegroundColor Red
            exit 1
        }

        if (Test-Path $masterPath) {
            Write-Host "  [*] Loading master CSV: $(Split-Path $masterPath -Leaf)" -ForegroundColor Cyan
            $masterRows = @(Import-Csv -Path $masterPath)
            foreach ($row in $masterRows) {
                $acctId = [string]$row.AccountID  ## Convert to string for consistent key matching
                if ($acctId) {
                    $masterHash[$acctId] = $row
                }
            }
            Write-Host "    — Loaded $($masterHash.Count) devices from master" -ForegroundColor Cyan
        } else {
            Write-Host "  [*] No existing master CSV found. Will create new: $(Split-Path $masterPath -Leaf)" -ForegroundColor Yellow
        }
        
        return $masterHash
    }

    Function Get-SelectionSignature {
        ## Compute a signature (hash) of selections for a datasource to detect changes
        ## Input: array of selection objects from EnumerateBackupSelections
        ## Output: canonical string representation (stable across runs)
        param (
            [object[]]$Selections,
            [string]$PluginId
        )
        
        if (-not $Selections) { return "" }
        
        $filtered = @($Selections | Where-Object { $_.PluginId -eq $PluginId })
        if ($filtered.Count -eq 0) { return "" }
        
        ## Build canonical representation: sorted by (Type, Priority, Path, Flags)
        $sig = $filtered | Sort-Object Type, Priority, Path, @{Expression={($_.Flags -join ';')}; Ascending=$true} `
            | ForEach-Object { "$($_.Type)|$($_.Priority)|$($_.Path)|$($_.Flags -join ';')" } `
            | ConvertTo-Json -Compress
        
        return $sig
    }

    Function Compare-SelectionsForChanges {
        ## Compare current selections to master row selections for a specific datasource
        ## Returns: $true if changed, $false if same or doesn't exist
        param (
            [object]$CurrentSelections,
            [object]$MasterRow,
            [string]$PluginId,       ## Full datasource name (e.g., "FileSystem")
            [string]$DSAbbr          ## Abbreviation for CSV columns (e.g., "FS")
        )
        
        $currentSig = Get-SelectionSignature -Selections $CurrentSelections -PluginId $PluginId
        $masterColName = "${DSAbbr}_Signature"
        $masterSig = if ($MasterRow -and $MasterRow.PSObject.Properties[$masterColName]) { $MasterRow.$masterColName } else { "" }
        
        return ($currentSig -ne $masterSig)
    }

    Function Format-SelectionForCsv {
        ## Format selection array into CSV columns: Include, Exclude, Signature, LastValidated, Changed
        ## Returns: hashtable with keys: {PluginId}_Include, {PluginId}_Exclude, {PluginId}_Signature, {PluginId}_LastValidated, {PluginId}_Changed
        param (
            [object[]]$Selections,
            [string]$PluginId,
            [object]$MasterRow,
            [bool]$IsChanged = $false,
            [datetime]$ValidatedTime = (Get-Date -AsUTC)
        )
        
        $result = @{}
        $sig = Get-SelectionSignature -Selections $Selections -PluginId $PluginId
        
        if (-not $Selections -or $Selections.Count -eq 0) {
            ## No selections for this datasource
            $result["${PluginId}_Include"] = ""
            $result["${PluginId}_Exclude"] = ""
            $result["${PluginId}_Signature"] = ""
            $result["${PluginId}_LastValidated"] = ""
            $result["${PluginId}_Changed"] = ""
        } else {
            ## API may return Type as "Inclusive"/"Exclusive" or "Include"/"Exclude"
            ## CRITICAL: Check path INSIDE ForEach loop, not in Where-Object filter
            ## This allows empty/missing paths (e.g., System State) to convert to "(all)"
            $includes = @($Selections | Where-Object { $_.Type -in @('Include','Inclusive') } |
                         ForEach-Object {
                             $p = if ($_.Path) { $_.Path } else { '(all)' }
                             if ($_.Flags -contains 'CreatedByAccountProfile') { "$p [p]" } else { $p }
                         } |
                         Sort-Object) -join ";"
            $excludes = @($Selections | Where-Object { $_.Type -in @('Exclude','Exclusive') -and $_.Path } |
                         ForEach-Object {
                             if ($_.Flags -contains 'CreatedByAccountProfile') { "$($_.Path) [p]" } else { $_.Path }
                         } | Sort-Object) -join ";"
            
            ## Clean up multiple "(all)" entries (e.g., "(all);(all)" -> "(all)")
            if ($includes -match '^\(all\);' -or $includes -match ';\(all\)') {
                $includes = '(all)'
            }
            
            $result["${PluginId}_Include"] = $includes
            $result["${PluginId}_Exclude"] = $excludes
            $result["${PluginId}_Signature"] = $sig
            $result["${PluginId}_LastValidated"] = $ValidatedTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
            $result["${PluginId}_Changed"] = if ($IsChanged) { "Yes" } else { "" }
        }
        
        return $result
    }

    Function Format-ScheduleForDs {
        ## Returns "N | Days | HH:mm,HH:mm" for schedules that cover a given plugin class list.
        ## Returns "" if no schedules match or $RawSchedules is null.
        param (
            [object]$RawSchedules,   ## raw JSON-RPC response object from EnumerateBackupSchedule
            [string[]]$PluginClasses ## e.g. @('FsBackupPlugin') or @('VssSystemState','SystemState','LinuxSystemState')
        )
        if (-not $RawSchedules) { return "" }
        try {
            $schedList = $RawSchedules.result.result
        } catch { return "" }
        if (-not $schedList) { return "" }

        ## DaysOfWeek bitmask: bit0=Sun,1=Mon,2=Tue,3=Wed,4=Thu,5=Fri,6=Sat
        $dayLabels = @('Su','Mo','Tu','We','Th','Fr','Sa')
        function DaysLabel([int]$mask) {
            if ($mask -eq 127) { return "Daily" }
            if ($mask -eq 62)  { return "Mon-Fri" }
            if ($mask -eq 65)  { return "Wknd" }
            $letters = ""
            for ($i = 0; $i -lt 7; $i++) {
                if ($mask -band (1 -shl $i)) { $letters += $dayLabels[$i][0] } else { $letters += "_" }
            }
            return $letters
        }

        $matched = @()
        foreach ($entry in $schedList) {
            if ($entry.Count -lt 2) { continue }
            $meta  = $entry[0]
            $scope = $entry[1]
            if (-not $meta -or -not $scope) { continue }
            if (-not $meta.Enabled) { continue }
            $entryPlugins = ($scope.Plugins -split "`t") | Where-Object { $_ -ne "" }
            $covers = $PluginClasses | Where-Object { $_ -in $entryPlugins }
            if ($covers.Count -eq 0) { continue }
            $hhmm = '{0:D2}:{1:D2}' -f [int]([int]$meta.Policy.FireTime / 3600), [int](([int]$meta.Policy.FireTime % 3600) / 60)
            $matched += [PSCustomObject]@{ Days = DaysLabel([int]$meta.Policy.DaysOfWeek); Time = $hhmm }
        }

        if ($matched.Count -eq 0) { return "" }
        $days  = ($matched | Select-Object -ExpandProperty Days | Select-Object -Unique) -join "/"
        $times = ($matched | Sort-Object Time | Select-Object -ExpandProperty Time) -join ","
        return "$($matched.Count) | $days | $times"
    }

    Function Format-HFScheduleForDs {
        ## Returns "Xh @ HH:mm | skip HH-HH" for HF schedules covering a given DataSource name list.
        ## Returns "" if no HF schedule exists for those datasources or $RawHFSchedules is null.
        param (
            [object]$RawHFSchedules,  ## raw JSON-RPC response from GetHighFrequentBackupSchedule
            [string[]]$DsNames        ## DataSource schema names e.g. @('VssHyperV') or @('VssSystemState','SystemState','LinuxSystemState')
        )
        if (-not $RawHFSchedules) { return "" }
        try { $hf = $RawHFSchedules.result.result } catch { return "" }
        if (-not $hf -or -not $hf.BackupScheduleItems) { return "" }

        $freqMap = @{
            'Every15Minutes' = '15m'; 'Every30Minutes' = '30m'; 'Every45Minutes' = '45m'
            'Every1Hour'     = '1h';  'Every2Hours'    = '2h';  'Every3Hours'    = '3h'
            'Every4Hours'    = '4h';  'Every6Hours'    = '6h';  'Every8Hours'    = '8h'
            'Every12Hours'   = '12h'
        }

        $items = @($hf.BackupScheduleItems | Where-Object { $_.PluginId -in $DsNames })
        if ($items.Count -eq 0) { return "" }

        $parts = @()
        foreach ($item in $items) {
            $freq  = if ($freqMap[$item.Frequency]) { $freqMap[$item.Frequency] } else { $item.Frequency }
            $first = '{0:D2}:{1:D2}' -f [int]([int]$item.TimeOfFirstBackup / 3600), [int](([int]$item.TimeOfFirstBackup % 3600) / 60)
            $part  = "$freq @ $first"
            if ($item.DoNotStartDuringWorkingHours -and $hf.WorkingHours) {
                $s = '{0:D2}:{1:D2}' -f [int]([int]$hf.WorkingHours.StartTime / 3600), [int](([int]$hf.WorkingHours.StartTime % 3600) / 60)
                $e = '{0:D2}:{1:D2}' -f [int]([int]$hf.WorkingHours.EndTime   / 3600), [int](([int]$hf.WorkingHours.EndTime   % 3600) / 60)
                $part += " | skip $s-$e"
            }
            $parts += $part
        }
        return $parts -join "; "
    }

    Function Merge-ResultIntoMaster {
        ## Merge a single device result into master hashtable
        ## Handles: new devices, reachable updates, unreachable preservation
        param (
            [hashtable]$MasterHash,
            [object]$DeviceObj,
            [object]$SelectionsObj,   ## Can be $null if unreachable
            [object]$RawSchedules,    ## Can be $null; raw EnumerateBackupSchedule response
            [object]$RawHFSchedules,  ## Can be $null; raw GetHighFrequentBackupSchedule response
            [bool]$WasReachable = $true
        )
        
        $acctId = [string]$DeviceObj.AccountID  ## Convert to string for consistent key matching with CSV-loaded keys
        $now = Get-Date -AsUTC
        $isNewDevice = -not $MasterHash.ContainsKey($acctId)
        Write-Debug "[MERGE] Device $acctId ($($DeviceObj.DeviceName)): isNew=$isNewDevice, reachable=$WasReachable"
        
        if (-not $MasterHash.ContainsKey($acctId)) {
            ## NEW DEVICE — create base row
            $newRow = [PSCustomObject]@{
                PartnerName = $DeviceObj.PartnerName
                PartnerID   = if ($DeviceObj.PartnerID) { $DeviceObj.PartnerID } else { "" }
                AccountID = $acctId
                DeviceName = $DeviceObj.DeviceName
                ComputerName = $DeviceObj.ComputerName
                IPAddress = $DeviceObj.IPAddress
                OS = $DeviceObj.OS
                Physicality = $DeviceObj.Physicality
                Manufacturer = $DeviceObj.Manufacturer
                Model = $DeviceObj.Model
                CPUCores  = $DeviceObj.CPUCores
                RAMBytes  = $DeviceObj.RAMBytes
                ProductID = if ($DeviceObj.ProductID) { $DeviceObj.ProductID } else { "" }
                Product   = if ($DeviceObj.Product)   { $DeviceObj.Product }   else { "" }
                ProfileID = if ($DeviceObj.ProfileID) { $DeviceObj.ProfileID } else { "" }
                Profile   = if ($DeviceObj.Profile)   { $DeviceObj.Profile }   else { "" }
                ClientVersion = if ($DeviceObj.ClientVersion) { $DeviceObj.ClientVersion } else { "" }
                CreationDate = if ($DeviceObj.CreationDate -is [datetime]) { $DeviceObj.CreationDate.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.CreationDate }
                TimeStamp = if ($DeviceObj.TimeStamp -is [datetime]) { $DeviceObj.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.TimeStamp }
                LastSuccess = if ($DeviceObj.LastSuccess -is [datetime]) { $DeviceObj.LastSuccess.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.LastSuccess }
                DataSources = $(
                    if ($WasReachable -and $SelectionsObj) {
                        $ds = ($SelectionsObj | Select-Object -ExpandProperty PluginId -Unique | Where-Object { $Script:DataSources.Contains($_) } | Sort-Object) -join ','
                        if ($ds) { $ds } else { $DeviceObj.DataSources }
                    } else { if ($DeviceObj.DataSources) { $DeviceObj.DataSources } else { "" } }
                )
                ReachableStatus = if ($WasReachable) { "Yes" } else { "Unknown" }
                LastValidated = $now.ToString("yyyy-MM-dd HH:mm:ss UTC")
                OrphanedStatus = "Active"
            }
            
            ## Add all datasource columns (empty for new devices until first query)
            foreach ($ds in $Script:DataSources.Values | Select-Object -Unique) {
                $newRow | Add-Member -NotePropertyName "${ds}_Include" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_Exclude" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_Signature" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_LastValidated" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_Changed" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_Sched" -NotePropertyValue ""
                $newRow | Add-Member -NotePropertyName "${ds}_HFSched" -NotePropertyValue ""
            }

            ## If reachable, populate with selections
            if ($WasReachable -and $SelectionsObj) {
                ## DEBUG: Log System State handling
                $sysStateInSelections = @($SelectionsObj | Where-Object { $_.PluginId -in @('VssSystemState', 'SystemState', 'LinuxSystemState') })
                if ($sysStateInSelections.Count -gt 0) {
                    if ($Script:DebugCDP) { Write-Host "    [DEBUG] $($DeviceObj.DeviceName): Found $($sysStateInSelections.Count) System State selections to merge" -ForegroundColor DarkGray }
                }
                
                ## Iterate by unique abbreviation — combine all plugins sharing the same abbrev (e.g. SysState = VssSystemState+SystemState+LinuxSystemState)
                $seenAbbrevs = [System.Collections.Generic.HashSet[string]]::new()
                foreach ($dsName in $Script:DataSources.Keys) {
                    $dsAbbr = $Script:DataSources[$dsName]
                    if (-not $seenAbbrevs.Add($dsAbbr)) { continue }
                    $sharedPlugins = @($Script:DataSources.GetEnumerator() | Where-Object { $_.Value -eq $dsAbbr } | Select-Object -ExpandProperty Key)
                    $dsSelections  = @($SelectionsObj | Where-Object { $_.PluginId -in $sharedPlugins })
                    $formatted = Format-SelectionForCsv -Selections $dsSelections -PluginId $dsAbbr -ValidatedTime $now
                    foreach ($k in $formatted.Keys) {
                        $newRow | Add-Member -NotePropertyName $k -NotePropertyValue $formatted[$k] -Force
                    }
                    $schedPlugins = @($sharedPlugins | ForEach-Object { $Script:PluginClassMap[$_] } | Where-Object { $_ })
                    $newRow | Add-Member -NotePropertyName "${dsAbbr}_Sched"   -NotePropertyValue (Format-ScheduleForDs   -RawSchedules   $RawSchedules   -PluginClasses $schedPlugins) -Force
                    $newRow | Add-Member -NotePropertyName "${dsAbbr}_HFSched" -NotePropertyValue (Format-HFScheduleForDs -RawHFSchedules $RawHFSchedules -DsNames       $sharedPlugins) -Force
                }
            }

            $MasterHash[$acctId] = $newRow
            Write-Host "    [+] NEW device: $($DeviceObj.DeviceName) (ID: $acctId)" -ForegroundColor Green
        } else {
            ## EXISTING DEVICE — update status and selections
            $existingRow = $MasterHash[$acctId]
            
            if ($WasReachable -and $SelectionsObj) {
                ## Device is currently reachable — update all metadata and selections
                $existingRow.ReachableStatus = "Yes"
                $existingRow.LastValidated    = $now.ToString("yyyy-MM-dd HH:mm:ss UTC")
                $existingRow.OrphanedStatus   = "Active"
                $existingRow.PartnerName      = $DeviceObj.PartnerName
                $existingRow.PartnerID        = if ($DeviceObj.PartnerID) { $DeviceObj.PartnerID } else { $existingRow.PartnerID }
                $existingRow.DeviceName       = $DeviceObj.DeviceName
                $existingRow.ComputerName     = $DeviceObj.ComputerName
                $existingRow.IPAddress        = $DeviceObj.IPAddress
                $existingRow.OS               = $DeviceObj.OS
                $existingRow.Physicality      = $DeviceObj.Physicality
                $existingRow.Manufacturer     = $DeviceObj.Manufacturer
                $existingRow.Model            = $DeviceObj.Model
                $existingRow.CPUCores         = $DeviceObj.CPUCores
                $existingRow.RAMBytes         = $DeviceObj.RAMBytes
                $existingRow | Add-Member -NotePropertyName 'ProductID'    -NotePropertyValue $DeviceObj.ProductID    -Force
                $existingRow | Add-Member -NotePropertyName 'Product'      -NotePropertyValue $DeviceObj.Product      -Force
                $existingRow | Add-Member -NotePropertyName 'ProfileID'    -NotePropertyValue $DeviceObj.ProfileID    -Force
                $existingRow | Add-Member -NotePropertyName 'Profile'      -NotePropertyValue $DeviceObj.Profile      -Force
                $existingRow | Add-Member -NotePropertyName 'ClientVersion' -NotePropertyValue $(if ($DeviceObj.ClientVersion) { $DeviceObj.ClientVersion } else { "" }) -Force
                $existingRow.TimeStamp        = if ($DeviceObj.TimeStamp -is [datetime]) { $DeviceObj.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.TimeStamp }
                $existingRow.LastSuccess      = if ($DeviceObj.LastSuccess -is [datetime]) { $DeviceObj.LastSuccess.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.LastSuccess }
                ## Derive DataSources from confirmed selections (more reliable than I78 from EnumerateAccounts)
                $confirmedDS = ($SelectionsObj | Select-Object -ExpandProperty PluginId -Unique | Where-Object { $Script:DataSources.Contains($_) } | Sort-Object) -join ','
                $dsValue = if ($confirmedDS) { $confirmedDS } elseif ($DeviceObj.DataSources) { $DeviceObj.DataSources } else { $existingRow.DataSources }
                $existingRow | Add-Member -NotePropertyName 'DataSources' -NotePropertyValue $dsValue -Force
                
                $seenAbbrevs = [System.Collections.Generic.HashSet[string]]::new()
                foreach ($dsName in $Script:DataSources.Keys) {
                    $dsAbbr = $Script:DataSources[$dsName]
                    if (-not $seenAbbrevs.Add($dsAbbr)) { continue }
                    $sharedPlugins = @($Script:DataSources.GetEnumerator() | Where-Object { $_.Value -eq $dsAbbr } | Select-Object -ExpandProperty Key)
                    $dsSelections  = @($SelectionsObj | Where-Object { $_.PluginId -in $sharedPlugins })
                    $isChanged = Compare-SelectionsForChanges -CurrentSelections $SelectionsObj -MasterRow $existingRow -PluginId $dsAbbr -DSAbbr $dsAbbr
                    $formatted = Format-SelectionForCsv -Selections $dsSelections -PluginId $dsAbbr -MasterRow $existingRow -IsChanged $isChanged -ValidatedTime $now
                    foreach ($k in $formatted.Keys) {
                        $existingRow | Add-Member -NotePropertyName $k -NotePropertyValue $formatted[$k] -Force
                    }
                    $schedPlugins = @($sharedPlugins | ForEach-Object { $Script:PluginClassMap[$_] } | Where-Object { $_ })
                    $existingRow | Add-Member -NotePropertyName "${dsAbbr}_Sched"   -NotePropertyValue (Format-ScheduleForDs   -RawSchedules   $RawSchedules   -PluginClasses $schedPlugins) -Force
                    $existingRow | Add-Member -NotePropertyName "${dsAbbr}_HFSched" -NotePropertyValue (Format-HFScheduleForDs -RawHFSchedules $RawHFSchedules -DsNames       $sharedPlugins) -Force
                }
                Write-Host "    [+] UPDATED device: $($DeviceObj.DeviceName) (ID: $acctId)" -ForegroundColor Green
            } else {
                ## Device is currently unreachable — preserve old selections, only update status and timestamp
                $existingRow.ReachableStatus = "No"
                $existingRow.LastValidated = $now.ToString("yyyy-MM-dd HH:mm:ss UTC")
                $existingRow.OrphanedStatus = "Active"
                if ($DeviceObj.PartnerID -and -not $existingRow.PartnerID) {
                    $existingRow | Add-Member -NotePropertyName 'PartnerID' -NotePropertyValue $DeviceObj.PartnerID -Force
                }
                if ($DeviceObj.DataSources) {
                    $existingRow | Add-Member -NotePropertyName 'DataSources' -NotePropertyValue $DeviceObj.DataSources -Force
                }
                if ($DeviceObj.ClientVersion) {
                    $existingRow | Add-Member -NotePropertyName 'ClientVersion' -NotePropertyValue $DeviceObj.ClientVersion -Force
                }
                $existingRow | Add-Member -NotePropertyName 'ProductID' -NotePropertyValue $DeviceObj.ProductID -Force
                $existingRow | Add-Member -NotePropertyName 'Product'   -NotePropertyValue $DeviceObj.Product   -Force
                $existingRow | Add-Member -NotePropertyName 'ProfileID' -NotePropertyValue $DeviceObj.ProfileID -Force
                $existingRow | Add-Member -NotePropertyName 'Profile'   -NotePropertyValue $DeviceObj.Profile   -Force
                Write-Host "    [-] UNREACHABLE device: $($DeviceObj.DeviceName) (ID: $acctId) — old selections preserved" -ForegroundColor Yellow
            }
        }
    }

    Function Save-MasterCsv {
        ## Save master hashtable to CSV, sorted by CreationDate ASC
        param (
            [hashtable]$MasterHash,
            [switch]$PassThru
        )
        
        $masterPath = Get-MasterCsvPath
        
        ## Convert hashtable values to array and sort by CreationDate (oldest first)
        $rows = @($MasterHash.Values)
        $rows = $rows | Sort-Object {
            try {
                [datetime]::ParseExact($_.CreationDate, "yyyy-MM-dd HH:mm:ss UTC", $null)
            } catch {
                [datetime]::MinValue
            }
        }
        
        ## ---- Safety check: detect and self-heal corrupt rows ----
        ## Two tiers of corruption:
        ##   Metadata quote: literal " in PartnerName/OS/etc — row unrecoverable, drop it (re-added next full run)
        ##   Schedule quote:  literal " in _Sched/_HFSched  — clear those columns only, metadata stays intact
        $metaColsToCheck  = @('PartnerName','OS','DeviceName','ComputerName','AccountID','Manufacturer','Model')
        $schedColNames    = @($rows[0].PSObject.Properties.Name | Where-Object { $_ -match '_Sched$|_HFSched$' })
        $droppedRows      = [System.Collections.Generic.List[object]]::new()
        $schedHealedCount = 0

        $rows = $rows | Where-Object {
            $row = $_
            ## Drop rows where metadata columns contain embedded quotes (unrecoverable column shift)
            $metaBad = $metaColsToCheck | Where-Object { $row.$_ -match '"' }
            if ($metaBad) { $droppedRows.Add($row); return $false }
            ## Heal rows where only schedule columns are corrupt
            $schedBad = $schedColNames | Where-Object { $row.$_ -match '"' }
            if ($schedBad) {
                foreach ($col in $schedColNames) { $row.$col = '' }
                $schedHealedCount++
            }
            return $true
        }

        if ($droppedRows.Count -gt 0) {
            Write-Host "  [!] Dropped $($droppedRows.Count) unrecoverable corrupt row(s) — will re-add on next full run:" -ForegroundColor Red
            $droppedRows | ForEach-Object { Write-Host "      $($_.AccountID) / $($_.DeviceName)" -ForegroundColor Red }
        }
        if ($schedHealedCount -gt 0) {
            Write-Host "  [!] Schedule column corruption healed on $schedHealedCount device(s) — will refresh on next run" -ForegroundColor Yellow
        }

        ## Define master schema (metadata columns first, then per-datasource columns in popularity order)
        $metaColumns = @(
            'PartnerName', 'PartnerID', 'AccountID', 'DeviceName', 'ComputerName', 'IPAddress',
            'OS', 'Physicality', 'Manufacturer', 'Model', 'CPUCores', 'RAMBytes', 'ProductID', 'Product', 'ProfileID', 'Profile', 'ClientVersion',
            'CreationDate', 'TimeStamp', 'LastSuccess', 'DataSources',
            'ReachableStatus', 'LastValidated', 'OrphanedStatus'
        )
        
        ## Build datasource columns in order (preserve order, remove duplicates)
        $dsColumns = @()
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($dsName in $Script:DataSources.Keys) {
            $dsAbbr = $Script:DataSources[$dsName]
            if ($seen.Add($dsAbbr)) {
                $dsColumns += "${dsAbbr}_Sched", "${dsAbbr}_HFSched", "${dsAbbr}_Include", "${dsAbbr}_Exclude", "${dsAbbr}_Signature", "${dsAbbr}_LastValidated", "${dsAbbr}_Changed"
            }
        }
        
        $allColumns = $metaColumns + $dsColumns
        
        ## Ensure all rows have all columns (backfill with empty strings if needed)
        $standardRows = @()
        foreach ($row in $rows) {
            $stdRow = [PSCustomObject]@{}
            foreach ($col in $allColumns) {
                $val = if ($row.PSObject.Properties[$col]) { $row.$col } else { "" }
                $stdRow | Add-Member -NotePropertyName $col -NotePropertyValue $val
            }
            $standardRows += $stdRow
        }
        
        ## Export to CSV — recreate rows with properties in exact order, then export
        ##  ConvertTo-Csv/Export-Csv always sort columns alphabetically, so we build CSV manually
        
        if ($Script:DebugCDP) { Write-Host "    [DEBUG] $($allColumns.Count) columns, first 5 are: $(($allColumns | Select-Object -First 5) -join ', ')" -ForegroundColor DarkGray }
        if ($Script:DebugCDP) { Write-Host "    [DEBUG] DS columns (Include only): $(($allColumns | Where-Object { $_ -match '_Include' }) -join ', ')" -ForegroundColor DarkGray }
        
        ## Header
        $headerLine = $allColumns -join ','
        if ($Script:DebugCDP) { Write-Host "    [DEBUG] Header line first 100 chars: $($headerLine.Substring(0, 100))" -ForegroundColor DarkGray }
        $csvLines = @($headerLine)
        
        ## Data rows - build from existing row, preserving property order
        foreach ($row in $standardRows) {
            $values = @()
            foreach ($colName in $allColumns) {
                $val = $row.$colName
                # Escape CSV values properly
                if ([string]::IsNullOrEmpty($val)) {
                    $values += ""
                } elseif ($val -like '* *' -or $val -like '*,*' -or $val -like '*"*' -or $val -match '[\r\n]') {
                    $values += "`"$($val -replace '"', '""')`""
                } else {
                    $values += $val
                }
            }
            $csvLines += $values -join ','
        }
        
        ## Write to file
        $csvLines -join "`r`n" | Out-File -FilePath $masterPath -Encoding UTF8 -Force
        
        Write-Host "  [*] Master CSV saved: $(Split-Path $masterPath -Leaf) ($($standardRows.Count) devices)" -ForegroundColor Cyan
        
        if ($PassThru) { return $masterPath }
    }

#endregion ----- Master CSV Persistence ----

#endregion ----- Functions ----

#region ----- Main Execution ----

    Send-APICredentialsCookie

    $scriptSw = [System.Diagnostics.Stopwatch]::StartNew()  ## total elapsed timer

    Write-Output $Script:strLineSeparator
    Write-Output ""

    Send-GetPartnerInfo $(if ($PartnerName) { $PartnerName } else { $Script:cred0 })

    if ($AllPartners) {} else { Send-EnumeratePartners }

    ## Load master CSV before device selection so RetryOnly can build its device list from it
    $masterHash = Load-MasterCsv
    Write-Output $Script:strLineSeparator

    if ($RetryOnly) {
        ## Skip full enumeration — build device list directly from master CSV "No" rows
        $Script:SelectedDevices = @($masterHash.Values | Where-Object { $_.ReachableStatus -in @('No','Unknown') } |
            ForEach-Object {
                [PSCustomObject]@{
                    AccountID    = [int]$_.AccountID
                    DeviceName   = $_.DeviceName
                    ComputerName = $_.ComputerName
                    PartnerName  = $_.PartnerName
                    PartnerID    = ""
                    IPAddress    = $_.IPAddress
                    OS           = $_.OS
                    Physicality  = $_.Physicality
                    Manufacturer = $_.Manufacturer
                    Model        = $_.Model
                    CPUCores     = $_.CPUCores
                    RAMBytes     = $_.RAMBytes
                    DataSources  = ""   ## empty = query all plugins (fallback in Get-RcgSelections)
                    CreationDate = $_.CreationDate
                    TimeStamp    = $_.TimeStamp
                    LastSuccess  = $_.LastSuccess
                }
            })
        $retryOnlyBeforeCount = $Script:SelectedDevices.Count
        Write-Host "  RetryOnly: $retryOnlyBeforeCount unreachable device(s) loaded from master CSV" -ForegroundColor Cyan
    } else {
        Send-GetDevices

        Write-Output $Script:strLineSeparator
        Write-Output "  $($Script:DeviceDetail.Count) devices found under $Script:PartnerName"
        Write-Output $Script:strLineSeparator

        if ($AllDevices) {
            $Script:SelectedDevices = $Script:DeviceDetail |
                Select-Object PartnerID,PartnerName,AccountID,DeviceName,ComputerName,IPAddress,OS,Physicality,Manufacturer,Model,CPUCores,RAMBytes,DataSources,CreationDate,TimeStamp,LastSuccess,FS_LastOK,SysState_LastOK,LinuxSS_LastOK,MSSQL_LastOK,HyperV_LastOK,NetShares_LastOK,MySQL_LastOK,Exchange_LastOK,VMware_LastOK,SharePt_LastOK,Oracle_LastOK
        } else {
            $Script:SelectedDevices = $Script:DeviceDetail |
                Select-Object PartnerID,PartnerName,AccountID,DeviceName,ComputerName,IPAddress,OS,Physicality,Manufacturer,Model,CPUCores,RAMBytes,DataSources,CreationDate,TimeStamp,LastSuccess,FS_LastOK,SysState_LastOK,LinuxSS_LastOK,MSSQL_LastOK,HyperV_LastOK,NetShares_LastOK,MySQL_LastOK,Exchange_LastOK,VMware_LastOK,SharePt_LastOK,Oracle_LastOK |
                Out-GridView -Title "Select Devices | $Script:PartnerName" -OutputMode Multiple
        }

        ## Sort by CreationDate ASC (oldest device first) — ensures new devices append at end of master CSV
        $Script:SelectedDevices = $Script:SelectedDevices |
            Sort-Object { if ($_.CreationDate -is [datetime]) { $_.CreationDate } else { [datetime]::MaxValue } }

        if ($null -eq $Script:SelectedDevices) {
            Write-Output $Script:strLineSeparator
            Write-Output "  No Devices Selected"
            Break
        }
    }

    ## Optional AccountID filter (for targeted test runs)
    if ($FilterAccountIDs.Count -gt 0) {
        $Script:SelectedDevices = @($Script:SelectedDevices | Where-Object { $_.AccountID -in $FilterAccountIDs })
        Write-Output "  Filtered to $($Script:SelectedDevices.Count) device(s) by -FilterAccountIDs"
    }

    ## Capture list of enumerated AccountIDs for orphan detection
    $enumeratedIds = @($Script:SelectedDevices | Select-Object -ExpandProperty AccountID)
    Write-Output "  Enumerated $($enumeratedIds.Count) device(s) for this run"

    Write-Output $Script:strLineSeparator
    Write-Output "  Processing $($Script:SelectedDevices.Count) device(s) | Throttle: $DeviceThrottle | Retry: $RetryFailed"
    Write-Output $Script:strLineSeparator

    ## Dated output subfolder — exports go here; master CSV stays in ExportPath root
    $OutputPath = "$ExportPath\Output\$(($CurrentDate -split '_')[0])"
    $null = New-Item -ItemType Directory -Path $OutputPath -Force

    ## Live CSV path set here so the file is ready before the parallel loop starts
    $Script:csvoutputfile = "$OutputPath\$($CurrentDate)_RemoteSelections_$($Script:PartnerName -replace(' \(.*\)','') -replace('[^a-zA-Z_0-9]',''))_$($Script:PartnerId).csv"

    ## Mutex for thread-safe CSV appends; flag tracks whether header row has been written yet
    $csvMutex         = [System.Threading.Mutex]::new($false, "RcgCsvWrite_$PID")
    $csvHeaderWritten = $false   ## accessed only from main thread (after parallel block)

    ## Thread-safe bags for collecting results from parallel runspaces
    $resultBag    = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $retryBag     = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $perfBag      = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $scheduleDump = [System.Collections.Concurrent.ConcurrentBag[object]]::new()  ## raw schedule responses for analysis

    ## Pass master hash to parallel runspace via $using: scope
    $masterHashForUse = $masterHash
    
    ## Track merged devices to prevent duplicates (important: each device should merge ONCE per run)
    $mergedIds = [System.Collections.Generic.HashSet[string]]::new()

    ## Capture function text + deps for re-import inside parallel runspaces
    ## ($Script: variables are not accessible across runspace boundaries)
    $rcgFuncText    = ${function:Get-RcgSelections}.ToString()
    $currentVisa    = $Script:visa
    $currentPartner = $Script:PartnerName
    $currentDsMap   = $Script:DataSources
    $currentApMap   = $Script:ApCharToPlugin

    if ($RetryOnly) {
        ## Feed "No" devices into retryBag so the immediate retry pass handles them
        foreach ($d in $Script:SelectedDevices) { $retryBag.Add($d) }
    }

    ## ---- Helper: build pivoted device row from selections ----
    function Build-DeviceRow ($DeviceObj, $Selections, $ReachableTag) {
        ## DataSources is already decoded to comma-separated plugin names — map to abbreviations
        $activeAbbrevs = if ($DeviceObj.DataSources) {
            ($DeviceObj.DataSources -split ',') | Where-Object { $_ } |
                ForEach-Object { $Script:DataSources[$_] } | Where-Object { $_ }
        } else { @() }

        ## Detect anomalies for this device
        $anomalies = @()

        ## Check for unreachable
        if ($ReachableTag -eq 'No') {
            $anomalies += 'UNREACHABLE'
        }


        ## Check for specific FS path selections (not full backup)
        $fsIncludes = @($Selections | Where-Object { $_.PluginId -eq 'FileSystem' -and $_.Type -eq 'Inclusive' })
        if ($fsIncludes.Count -gt 0) {
            $hasSpecific = $fsIncludes | Where-Object { $_.Path -ne '' -and $_.Path -ne '\' -and $_.Path -ne '/' }
            if ($hasSpecific) { $anomalies += 'SPECIFIC_FS' }
        }

        ## Check for orphaned datasources (enabled but no selections)
        foreach ($abbr in $activeAbbrevs) {
            $dsPlugin = $Script:DataSources.GetEnumerator() | Where-Object { $_.Value -eq $abbr } | Select-Object -First 1 -ExpandProperty Key
            if ($dsPlugin) {
                $hasSelections = @($Selections | Where-Object { $_.PluginId -eq $dsPlugin }).Count -gt 0
                if (-not $hasSelections) {
                    $anomalies += 'ORPHANED_DS'
                    break
                }
            }
        }

        ## Check for profile-based selections
        $hasProfile = @($Selections | Where-Object { $_.Flags -contains 'CreatedByAccountProfile' }).Count -gt 0
        if ($hasProfile) {
            $anomalies += 'PROFILE_BASED'
        }

        ## Check for devices with no exclusions (unusual if large fleet)
        $hasAnyExclusion = @($Selections | Where-Object { $_.Type -eq 'Exclusive' -and $_.Path }).Count -gt 0
        if (-not $hasAnyExclusion) {
            $anomalies += 'NO_EXCLUSIONS'
        }

        ## Flag if device is unreachable on retry
        if ($ReachableTag -eq 'Retry-OK') {
            $anomalies += 'RETRY_RECOVERY'
        }

        $anomalyString = if ($anomalies.Count -gt 0) { $anomalies -join '; ' } else { '' }

        $deviceRow = [ordered]@{
            Anomalies    = $anomalyString
            PartnerID    = $DeviceObj.PartnerID
            PartnerName  = $DeviceObj.PartnerName
            AccountID    = $DeviceObj.AccountID
            DeviceName   = $DeviceObj.DeviceName
            ComputerName = $DeviceObj.ComputerName
            IPAddress    = $DeviceObj.IPAddress
            OS           = $DeviceObj.OS
            Physicality  = $DeviceObj.Physicality
            Manufacturer = $DeviceObj.Manufacturer
            Model        = $DeviceObj.Model
            CPUCores     = $DeviceObj.CPUCores
            RAMBytes     = $DeviceObj.RAMBytes
            ProductID    = $DeviceObj.ProductID
            Product      = $DeviceObj.Product
            ProfileID    = $DeviceObj.ProfileID
            Profile      = $DeviceObj.Profile
            ActiveDS     = $activeAbbrevs -join '; '
            CreationDate = if ($DeviceObj.CreationDate -is [datetime]) { $DeviceObj.CreationDate.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { $DeviceObj.CreationDate }
            LastSuccess  = if ($DeviceObj.LastSuccess  -is [datetime]) { $DeviceObj.LastSuccess.ToString("yyyy-MM-dd HH:mm:ss UTC")  } else { $DeviceObj.LastSuccess  }
            TimeStamp    = if ($DeviceObj.TimeStamp    -is [datetime]) { $DeviceObj.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss UTC")    } else { $DeviceObj.TimeStamp    }
            Reachable    = $ReachableTag
        }

        foreach ($ds in $Script:DataSources.Keys) {
            $abbrev = $Script:DataSources[$ds]

            ## Skip if this abbreviation was already written by an earlier plugin that shares it
            if ($deviceRow.Contains("$abbrev Inc+")) { continue }

            ## Collect rows from ALL plugins that share this abbreviation (e.g. VssSystemState + SystemState)
            $sharedPlugins = @($Script:DataSources.GetEnumerator() |
                Where-Object { $_.Value -eq $abbrev } | Select-Object -ExpandProperty Key)
            $dsRows = @($Selections | Where-Object { $_.PluginId -in $sharedPlugins })

            if ($dsRows.Count -eq 0) {
                ## null = plugin not configured or not reporting path-based selections — leave blank
                $includes = ''
                $excludes = ''
            } else {
                $includes = ($dsRows | Where-Object { $_.Type -eq 'Inclusive' } |
                             ForEach-Object {
                                 ## For FileSystem, SystemState, and MSSQL: normalize \, /, or System State to (all); for others, use path as-is
                                 if ($abbrev -in @('FS', 'SysState', 'MSSQL') -and ($_.Path -eq '\' -or $_.Path -eq '/' -or $_.Path -eq 'System State' -or $_.Path -eq '' -or -not $_.Path)) {
                                     $p = '(all)'
                                 } else {
                                     $p = if ($_.Path) { $_.Path } else { '(all)' }
                                 }
                                 if ($_.Flags -contains 'CreatedByAccountProfile') { "$p [p]" } else { $p }
                             } | Select-Object -Unique) -join "`n"
                $excludes = ($dsRows | Where-Object { $_.Type -eq 'Exclusive' -and $_.Path } |
                             ForEach-Object {
                                 if ($_.Flags -contains 'CreatedByAccountProfile') { "$($_.Path) [p]" } else { $_.Path }
                             } | Select-Object -Unique) -join "`n"
            }
            ## Column order: Last, Inc+, Exc- (LastOK timestamp appears before selections for easier scanning)
            $deviceRow["$abbrev Last"] = $DeviceObj."${abbrev}_LastOK"
            $deviceRow["$abbrev Inc+"] = $includes
            $deviceRow["$abbrev Exc-"] = $excludes
        }
        return [PSCustomObject]$deviceRow
    }

    ## ---- Helper: append one device row to CSV (main-thread only) ----
    function Append-ToCsv ($Row) {
        if (-not $Export) { return }
        if (-not $Script:csvHeaderWritten) {
            $Row | Export-Csv -Path $Script:csvoutputfile -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
            $Script:csvHeaderWritten = $true
        } else {
            $Row | Export-Csv -Path $Script:csvoutputfile -Delimiter $Delimiter -NoTypeInformation -Append -Encoding UTF8
        }
    }

    ## ---- Helper: convert a master CSV row into the analyst-friendly export row format ----
    function Convert-MasterRowToExportRow ($MasterRow) {
        $anomalies = @()
        if ($MasterRow.ReachableStatus -notin @('Yes')) { $anomalies += 'UNREACHABLE' }
        if ($MasterRow.OrphanedStatus -eq 'NotInCurrentInventory') { $anomalies += 'ORPHANED' }

        ## Specific FS path selections (not full backup)
        $fsInc = $MasterRow.FS_Include
        if ($fsInc) {
            $hasSpecific = ($fsInc -split ';') | Where-Object { $_ -ne '' -and $_ -ne '\' -and $_ -ne '/' -and $_ -ne '(all)' }
            if ($hasSpecific) { $anomalies += 'SPECIFIC_FS' }
        }

        ## Profile: any _Include column contains [p]
        $incCols = $MasterRow.PSObject.Properties.Name | Where-Object { $_ -match '_Include$' }
        if ($incCols | Where-Object { $MasterRow.$_ -match '\[p\]' }) { $anomalies += 'PROFILE' }

        ## Exclusions: any _Exclude column has data
        $excCols = $MasterRow.PSObject.Properties.Name | Where-Object { $_ -match '_Exclude$' }
        if ($excCols | Where-Object { "$($MasterRow.$_)".Trim() -ne '' }) { $anomalies += 'EXCLUSIONS' }

        $row = [ordered]@{
            Anomalies    = $anomalies -join '; '
            PartnerID    = $MasterRow.PartnerID
            PartnerName  = $MasterRow.PartnerName
            AccountID    = $MasterRow.AccountID
            DeviceName   = $MasterRow.DeviceName
            ComputerName = $MasterRow.ComputerName
            IPAddress    = $MasterRow.IPAddress
            OS           = $MasterRow.OS
            Physicality  = $MasterRow.Physicality
            Manufacturer = $MasterRow.Manufacturer
            Model        = $MasterRow.Model
            CPUCores      = $MasterRow.CPUCores
            RAMBytes      = $MasterRow.RAMBytes
            ProductID     = $MasterRow.ProductID
            Product       = $MasterRow.Product
            ProfileID     = $MasterRow.ProfileID
            Profile       = $MasterRow.Profile
            ClientVersion = $MasterRow.ClientVersion
            DataSources   = $MasterRow.DataSources
            CreationDate = $MasterRow.CreationDate
            LastSuccess  = $MasterRow.LastSuccess
            TimeStamp    = $MasterRow.TimeStamp
            Reachable    = $MasterRow.ReachableStatus
        }

        $activeDsList = @()
        $seenAbbrevs  = [System.Collections.Generic.HashSet[string]]::new()

        foreach ($dsName in $Script:DataSources.Keys) {
            $abbrev = $Script:DataSources[$dsName]
            if (-not $seenAbbrevs.Add($abbrev)) { continue }   ## skip duplicate abbreviations (SysState variants)
            if ($Script:ExcludeColumnSet.Contains($abbrev)) { continue }   ## user-requested column exclusion

            $inc  = $MasterRow."${abbrev}_Include"
            $exc  = $MasterRow."${abbrev}_Exclude"
            $last = $MasterRow."${abbrev}_LastValidated"

            ## Convert ";" path separator to newline for Excel readability
            $incFmt = if ($inc)  { $inc -replace ';', "`n" } else { '' }
            $excFmt = if ($exc)  { $exc -replace ';', "`n" } else { '' }

            if ($inc -or $exc) { $activeDsList += $abbrev }

            $row["$abbrev Sched"]   = $MasterRow."${abbrev}_Sched"
            $row["$abbrev HFSched"] = $MasterRow."${abbrev}_HFSched"
            $row["$abbrev Last"]    = $last
            $row["$abbrev Inc+"]  = $incFmt
            $row["$abbrev Exc-"]  = $excFmt
        }

        return [PSCustomObject]$row
    }

    if (-not $RetryOnly) {

    ## ======== FIRST PASS — parallel across all selected devices ========
    $totalDevices   = @($Script:SelectedDevices).Count
    $firstPassStart = $scriptSw.Elapsed
    $Script:SelectedDevices | ForEach-Object -Parallel {
        $device = $_

        ## Re-import Get-RcgSelections into this runspace
        New-Item -Path 'function:\Get-RcgSelections' `
            -Value ([scriptblock]::Create($using:rcgFuncText)) -Force | Out-Null

        $xb    = $using:resultBag
        $total = $using:totalDevices
        Write-Progress -Id 1 `
            -Activity   "RCG Selections — $($using:currentPartner)" `
            -Status     "[$($xb.Count+1)/$total] $($device.DeviceName) — connecting..." `
            -PercentComplete ([int]($xb.Count * 100 / $total))

        $result = $null
        try {
            $result = Get-RcgSelections `
                -AccountId         $device.AccountID `
                -DeviceName        $device.DeviceName `
                -DeviceDataSources $device.DataSources `
                -Visa              $using:currentVisa `
                -DsMap             $using:currentDsMap `
                -ApMap             $using:currentApMap
            if ($result -and $result.Selections) {
                $queriedCount = @($using:currentDsMap.Keys | ForEach-Object { $_ }).Count
                $activeDs = if ($device.DataSources) { ($device.DataSources -split ',').Count } else { 0 }
                Write-Host "  [$($activeDs) active / $queriedCount total plugins] $($device.DeviceName)" -ForegroundColor DarkGray
            }
        } catch {
            Write-Host "  [!] $($device.DeviceName): $($_.Exception.Message)" -ForegroundColor Red
        }

        $selections     = if ($result) { $result.Selections     } else { $null }
        $perfMs         = if ($result) { $result.PerfMs         } else { $null }
        $rawSchedules   = if ($result) { $result.RawSchedules   } else { $null }
        $rawHFSchedules = if ($result) { $result.RawHFSchedules } else { $null }

        ## Capture raw schedule response for markdown dump (regardless of success/failure)
        $sdb = $using:scheduleDump
        $sdb.Add([PSCustomObject]@{ Device = $device; RawSchedules = $rawSchedules; RawHFSchedules = $rawHFSchedules })

        ## Defer failed devices to retry pass (will do fresh auth there)
        $rb = $using:retryBag
        $xb = $using:resultBag

        if ($null -eq $selections) {
            $rb.Add($device)
        }

        $xb.Add([PSCustomObject]@{
            Device          = $device
            Selections      = $selections
            RawSchedules    = $rawSchedules
            RawHFSchedules  = $rawHFSchedules
            PerfMs          = $perfMs
            IsRetry         = $false
            FailureReason   = ""
        })

        $done      = $xb.Count
        $total     = $using:totalDevices
        $pct       = [int]($done * 100 / $total)
        $remaining = $total - $done

        Write-Progress -Id 1 `
            -Activity        "RCG Selections — $($using:currentPartner)" `
            -Status          "[$done/$total] $($device.DeviceName) — done ($remaining left)" `
            -PercentComplete $pct

        if ($null -eq $selections) {
            Write-Host "  [-] $($device.DeviceName): unreachable — retry   $pct% ($done/$total, $remaining left)" -ForegroundColor Yellow
        } else {
            Write-Host "  [+] $($device.DeviceName): $($selections.Count) selection(s)   $pct% ($done/$total, $remaining left)" -ForegroundColor Green
        }

    } -ThrottleLimit $DeviceThrottle
    Write-Progress -Id 1 -Activity "RCG Selections" -Completed
    $firstPassElapsed = $scriptSw.Elapsed - $firstPassStart

    ## ---- Process first-pass results (main thread — safe to write CSV and update master) ----
    foreach ($item in $resultBag) {
        $tag = if ($null -ne $item.Selections) { 'Yes' } else { 'No (retry pending)' }
        $row = Build-DeviceRow -DeviceObj $item.Device -Selections $item.Selections -ReachableTag $tag

        ## Update master hash with result (new device, update reachable, or preserve unreachable)
        ## Only merge if not already merged (prevent duplicates from multiple retry passes)
        $addResult = $mergedIds.Add([string]$item.Device.AccountID)  ## Convert to string for consistent matching
        if ($addResult) {
            Merge-ResultIntoMaster -MasterHash $masterHashForUse -DeviceObj $item.Device -SelectionsObj $item.Selections -RawSchedules $item.RawSchedules -RawHFSchedules $item.RawHFSchedules -WasReachable ($null -ne $item.Selections)
        }

        if ($null -ne $item.Selections) {
            ## Write reachable devices live; unreachable written after retry pass
            $Script:SelectionResults.Add($row)
            Append-ToCsv $row

            foreach ($sel in $item.Selections) {
                $Script:SelectionDetails.Add([PSCustomObject]@{
                    PartnerName = $item.Device.PartnerName; AccountID  = $item.Device.AccountID
                    DeviceName  = $item.Device.DeviceName;  DataSource = $sel.PluginId
                    Type        = $sel.Type; Priority = $sel.Priority; Path = $sel.Path
                    Flags       = ($sel.Flags -join ',')
                })
            }
        }

        if ($item.PerfMs) {
            $perfBag.Add([PSCustomObject]@{
                DeviceName = $item.Device.DeviceName
                Pass       = 'First'
                S1_ms      = $item.PerfMs.S1_Endpoint
                S2_ms      = $item.PerfMs.S2_RcgAuth
                S3_ms      = $item.PerfMs.S3_Plugins
                Total_ms   = $item.PerfMs.Total
            })
        }
    }

    } ## end if (-not $RetryOnly) — first pass + result processing

    ## ======== IMMEDIATE RETRY PASS (Main Thread — Per-Device Fresh Auth) ========
    if ($retryBag.Count -gt 0) {
        Write-Output $Script:strLineSeparator
        Write-Host "  Immediate retry pass (parallel, per-device fresh auth) — $($retryBag.Count) device(s)..." -ForegroundColor Yellow

        ## Refresh main visa before retry; re-auth if expired
        Get-VisaTime
        $retryVisa = $Script:visa

        $immediateBag   = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $totalImmediate = $retryBag.Count
        $immediateSuccessCount = 0

        ## Each Get-RcgSelections call handles Stage 1+2+3 internally using the main visa
        $retryBag | ForEach-Object -Parallel {
            $device = $_

            New-Item -Path 'function:\Get-RcgSelections' `
                -Value ([scriptblock]::Create($using:rcgFuncText)) -Force | Out-Null

            $ib    = $using:immediateBag
            $visa  = $using:retryVisa
            $dsMap = $using:currentDsMap
            $apMap = $using:currentApMap

            $result  = $null
            $success = $false
            try {
                $result = Get-RcgSelections `
                    -AccountId         $device.AccountID `
                    -DeviceName        $device.DeviceName `
                    -DeviceDataSources $device.DataSources `
                    -Visa              $visa `
                    -DsMap             $dsMap `
                    -ApMap             $apMap
                if ($result -and $result.Selections) {
                    Write-Host "  [+] $($device.DeviceName): retry OK ($($result.Selections.Count) selections)" -ForegroundColor Green
                    $success = $true
                } else {
                    Write-Host "  [-] $($device.DeviceName): retry - no selections returned" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  [-] $($device.DeviceName): retry - $($_.Exception.Message)" -ForegroundColor Red
            }

            $ib.Add([PSCustomObject]@{
                Device     = $device
                Selections = if ($result) { $result.Selections } else { $null }
                PerfMs     = if ($result) { $result.PerfMs     } else { $null }
                Success    = $success
            })
        } -ThrottleLimit 10

        ## Process immediate retry results
        foreach ($item in $immediateBag) {
            if ($item.Success) {
                ## Successful on immediate retry — write to results
                $row = Build-DeviceRow -DeviceObj $item.Device -Selections $item.Selections -ReachableTag 'Yes (immediate retry)'
                
                ## Update master hash with successful retry result (only if not already merged)
                $addResult = $mergedIds.Add([string]$item.Device.AccountID)  ## Convert to string for consistent matching
                if ($addResult) {
                    Merge-ResultIntoMaster -MasterHash $masterHashForUse -DeviceObj $item.Device -SelectionsObj $item.Selections -RawSchedules $null -RawHFSchedules $null -WasReachable $true
                }
                
                $Script:SelectionResults.Add($row)
                Append-ToCsv $row

                foreach ($sel in $item.Selections) {
                    $Script:SelectionDetails.Add([PSCustomObject]@{
                        PartnerName = $item.Device.PartnerName; AccountID  = $item.Device.AccountID
                        DeviceName  = $item.Device.DeviceName;  DataSource = $sel.PluginId
                        Type        = $sel.Type; Priority = $sel.Priority; Path = $sel.Path
                        Flags       = ($sel.Flags -join ',')
                    })
                }

                if ($item.PerfMs) {
                    $perfBag.Add([PSCustomObject]@{
                        DeviceName = $item.Device.DeviceName
                        Pass       = 'Immediate Retry'
                        S1_ms      = $item.PerfMs.S1_Endpoint
                        S2_ms      = $item.PerfMs.S2_RcgAuth
                        S3_ms      = $item.PerfMs.S3_Plugins
                        Total_ms   = $item.PerfMs.Total
                    })
                }

                ## Remove from retry bag since it succeeded
                $retryBag = $retryBag | Where-Object { $_.AccountID -ne $item.Device.AccountID }
                $immediateSuccessCount++
            } else {
                ## Failed on immediate retry — mark as No in master
                $addResult = $mergedIds.Add([string]$item.Device.AccountID)
                if ($addResult) {
                    Merge-ResultIntoMaster -MasterHash $masterHashForUse -DeviceObj $item.Device -SelectionsObj $null -RawSchedules $null -RawHFSchedules $null -WasReachable $false
                }
            }
        }

        $immediateFailCount = $totalImmediate - $immediateSuccessCount
        Write-Host "  Immediate retry results: $immediateSuccessCount recovered, $immediateFailCount still unreachable" -ForegroundColor Cyan
        Write-Output $Script:strLineSeparator
    }

    if ($RetryOnly) {
        $retryOnlyAfterCount = @($masterHash.Values | Where-Object { $_.ReachableStatus -in @('No','Unknown') }).Count
        $retryOnlyRecovered  = $retryOnlyBeforeCount - $retryOnlyAfterCount
        Write-Output $Script:strLineSeparator
        Write-Host "  RetryOnly Summary:" -ForegroundColor Cyan
        Write-Host "    Unreachable before : $retryOnlyBeforeCount"
        Write-Host "    Recovered          : $retryOnlyRecovered" -ForegroundColor $(if ($retryOnlyRecovered -gt 0) { 'Green' } else { 'Yellow' })
        Write-Host "    Still unreachable  : $retryOnlyAfterCount" -ForegroundColor $(if ($retryOnlyAfterCount -eq 0) { 'Green' } else { 'Red' })
        Write-Output $Script:strLineSeparator
    }

    ## ======== RETRY PASS ========
    if (-not $RetryOnly -and $RetryFailed -and $retryBag.Count -gt 0) {
        Write-Output $Script:strLineSeparator
        Write-Host "  End-of-script retry pass — $($retryBag.Count) device(s)..." -ForegroundColor Yellow
        Write-Output $Script:strLineSeparator

        Get-VisaTime
        $currentVisa = $Script:visa
        Write-Host "  Visa refreshed for end-of-script retry pass." -ForegroundColor DarkGray
        $retryStart = $scriptSw.Elapsed

        $retryResultBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $totalRetry = $retryBag.Count

        $retryBag | ForEach-Object -Parallel {
            $device = $_

            New-Item -Path 'function:\Get-RcgSelections' `
                -Value ([scriptblock]::Create($using:rcgFuncText)) -Force | Out-Null

            Write-Host "  [R] $($device.DeviceName)  (ID: $($device.AccountID))" -ForegroundColor Yellow

            $rrb   = $using:retryResultBag
            $rtot  = $using:totalRetry
            Write-Progress -Id 1 `
                -Activity        "RCG Selections — RETRY — $($using:currentPartner)" `
                -Status          "[$($rrb.Count+1)/$rtot] $($device.DeviceName) — connecting..." `
                -PercentComplete ([int]($rrb.Count * 100 / $rtot))

            $result = $null
            $resultStatus = ""
            try {
                $result = Get-RcgSelections `
                    -AccountId         $device.AccountID `
                    -DeviceName        $device.DeviceName `
                    -DeviceDataSources $device.DataSources `
                    -Visa              $using:currentVisa `
                    -DsMap             $using:currentDsMap `
                    -ApMap             $using:currentApMap
                
                if ($result -and $result.Selections) {
                    $resultStatus = "OK"
                } else {
                    $resultStatus = "empty"
                }
            } catch {
                $resultStatus = "error: $($_.Exception.Message)"
            }

            $selections = if ($result) { $result.Selections } else { $null }
            $perfMs     = if ($result) { $result.PerfMs     } else { $null }

            $rrb = $using:retryResultBag
            $rrb.Add([PSCustomObject]@{
                Device     = $device
                Selections = $selections
                PerfMs     = $perfMs
                IsRetry    = $true
            })

            $rdone  = $rrb.Count
            $rtotal = $using:totalRetry
            $rpct   = [int]($rdone * 100 / $rtotal)
            $rrem   = $rtotal - $rdone

            Write-Progress -Id 1 `
                -Activity        "RCG Selections — RETRY — $($using:currentPartner)" `
                -Status          "[$rdone/$rtotal] $($device.DeviceName) — done ($rrem left)" `
                -PercentComplete $rpct

            if ($null -eq $selections) {
                Write-Host "  [R] $($device.DeviceName) — retry failed: $resultStatus" -ForegroundColor Red
            } else {
                Write-Host "  [+] $($device.DeviceName) — retry OK: $($selections.Count) selection(s)" -ForegroundColor Green
            }

        } -ThrottleLimit $DeviceThrottle
        Write-Progress -Id 1 -Activity "RCG Selections" -Completed
        $retryElapsed = $scriptSw.Elapsed - $retryStart

        ## Process retry results (main thread)
        foreach ($item in $retryResultBag) {
            $tag = if ($null -ne $item.Selections) { 'Retry-OK' } else { 'No' }
            $row = Build-DeviceRow -DeviceObj $item.Device -Selections $item.Selections -ReachableTag $tag

            ## Update master hash with retry result (successful or still unreachable, only if not already merged)
            $addResult = $mergedIds.Add([string]$item.Device.AccountID)  ## Convert to string for consistent matching
            if ($addResult) {
                Merge-ResultIntoMaster -MasterHash $masterHashForUse -DeviceObj $item.Device -SelectionsObj $item.Selections -RawSchedules $null -RawHFSchedules $null -WasReachable ($null -ne $item.Selections)
            }

            $Script:SelectionResults.Add($row)
            Append-ToCsv $row

            if ($null -ne $item.Selections) {
                foreach ($sel in $item.Selections) {
                    $Script:SelectionDetails.Add([PSCustomObject]@{
                        PartnerName = $item.Device.PartnerName; AccountID  = $item.Device.AccountID
                        DeviceName  = $item.Device.DeviceName;  DataSource = $sel.PluginId
                        Type        = $sel.Type; Priority = $sel.Priority; Path = $sel.Path
                        Flags       = ($sel.Flags -join ',')
                    })
                }
            }

            if ($item.PerfMs) {
                $perfBag.Add([PSCustomObject]@{
                    DeviceName = $item.Device.DeviceName
                    Pass       = 'Retry'
                    S1_ms      = $item.PerfMs.S1_Endpoint
                    S2_ms      = $item.PerfMs.S2_RcgAuth
                    S3_ms      = $item.PerfMs.S3_Plugins
                    Total_ms   = $item.PerfMs.Total
                })
            }
        }
    } elseif (-not $RetryOnly -and $retryBag.Count -gt 0) {
        ## RetryFailed=$false — write unreachable rows to CSV now
        foreach ($item in ($resultBag | Where-Object { $null -eq $_.Selections })) {
            $row = Build-DeviceRow -DeviceObj $item.Device -Selections $null -ReachableTag 'No'
            $Script:SelectionResults.Add($row)
            Append-ToCsv $row
        }
    }

    $csvMutex.Dispose()
    $scriptSw.Stop()

    ## ---- Write raw schedule dump markdown (debug only) ----
    if ($DebugCDP -and $scheduleDump.Count -gt 0) {
        $mdPath = "$OutputPath\$($CurrentDate)_ScheduleDump_$($Script:PartnerName -replace '[^a-zA-Z_0-9]','').md"
        $md = [System.Text.StringBuilder]::new()
        $null = $md.AppendLine("# Raw EnumerateBackupSchedule Dump")
        $null = $md.AppendLine("Partner: $Script:PartnerName  |  Run: $CurrentDate  |  Devices: $($scheduleDump.Count)")
        $null = $md.AppendLine("")

        foreach ($entry in ($scheduleDump | Sort-Object { $_.Device.DeviceName })) {
            $null = $md.AppendLine("## $($entry.Device.DeviceName)  (ID: $($entry.Device.AccountID))")
            $null = $md.AppendLine("**Partner:** $($entry.Device.PartnerName)")
            $null = $md.AppendLine("")
            $null = $md.AppendLine("### EnumerateBackupSchedule")
            if ($entry.RawSchedules) {
                $null = $md.AppendLine('```json')
                $null = $md.AppendLine(($entry.RawSchedules | ConvertTo-Json -Depth 10))
                $null = $md.AppendLine('```')
            } else {
                $null = $md.AppendLine("_No response (device unreachable or call skipped)_")
            }
            $null = $md.AppendLine("")
            $null = $md.AppendLine("### GetHighFrequentBackupSchedule")
            if ($entry.RawHFSchedules) {
                $null = $md.AppendLine('```json')
                $null = $md.AppendLine(($entry.RawHFSchedules | ConvertTo-Json -Depth 10))
                $null = $md.AppendLine('```')
            } else {
                $null = $md.AppendLine("_No response (device unreachable or call skipped)_")
            }
            $null = $md.AppendLine("")
        }

        $md.ToString() | Set-Content -Path $mdPath -Encoding UTF8
        Write-Output $Script:strLineSeparator
        Write-Host "  Schedule dump (raw): $mdPath" -ForegroundColor Cyan
        Write-Output $Script:strLineSeparator
    }

    ## ---- Mark orphaned devices (only on full-fleet runs) ----
    ## Partial runs (-DeviceCount N, -RetryOnly) must not stomp OrphanedStatus for unseen devices
    if (-not $RetryOnly -and $enumeratedIds.Count -ge $masterHashForUse.Count) {
        foreach ($acctId in $masterHashForUse.Keys) {
            if ($acctId -notin $enumeratedIds) {
                $masterHashForUse[$acctId].OrphanedStatus = "NotInCurrentInventory"
            }
        }
    }

    ## ---- Save master CSV (persistent across runs) ----
    $masterCsvPath = Save-MasterCsv -MasterHash $masterHashForUse -PassThru
    Write-Output $Script:strLineSeparator
    if ($masterCsvPath) {
        Write-Host "  Master CSV saved for next run: $masterCsvPath" -ForegroundColor Green
    }
    Write-Output $Script:strLineSeparator

#endregion ----- Main Execution ----

#region ----- Report ----

    Write-Output ""
    Write-Output $Script:strLineSeparator
    Write-Host "  Analysis Report - $Script:PartnerName" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator

    # Per-device human-readable summary (gated behind -DebugCDP)
    if ($Script:DebugCDP) {
    Write-Output $Script:strLineSeparator
    Write-Host "  Per-Device Summary" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator

    ForEach ($devName in ($Script:SelectionDetails | Select-Object -ExpandProperty DeviceName -Unique)) {
        $devRows = $Script:SelectionDetails | Where-Object { $_.DeviceName -eq $devName }
        Write-Host "`n  $devName" -ForegroundColor Yellow

        ForEach ($ds in ($devRows | Select-Object -ExpandProperty DataSource -Unique)) {
            $abbrev   = if ($Script:DataSources[$ds]) { $Script:DataSources[$ds] } else { $ds }
            $rows     = $devRows | Where-Object { $_.DataSource -eq $ds }
            $excPaths = @($rows | Where-Object { $_.Type -eq 'Exclusive' -and $_.Path } |
                          Select-Object -ExpandProperty Path)
            $incList  = ($rows | Where-Object { $_.Type -eq 'Inclusive' } |
                         ForEach-Object {
                             ## For FileSystem and SystemState: normalize \, /, or System State to (all); for others, use path as-is
                             if ($abbrev -in @('FS', 'SysState') -and ($_.Path -eq '\' -or $_.Path -eq '/' -or $_.Path -eq 'System State' -or $_.Path -eq '' -or -not $_.Path)) {
                                 $p = '(all)'
                             } else {
                                 $p = if ($_.Path) { $_.Path } else { '(all)' }
                             }
                             $flag = if ($_.Flags -like '*CreatedByAccountProfile*') { '(p)' } else { '' }
                             "$p`t$flag"  ## tab-separated for easier parsing during output
                         }) -join ' | '

            Write-Host "    [$abbrev] " -ForegroundColor Cyan -NoNewline
            
            ## Write incList with special handling for (p) badge
            $parts = $incList -split ' \| '
            for ($i = 0; $i -lt $parts.Count; $i++) {
                if ($i -gt 0) { Write-Host " | " -ForegroundColor Green -NoNewline }
                $pathAndFlag = $parts[$i] -split "`t"
                $path = $pathAndFlag[0]
                $flag = if ($pathAndFlag.Count -gt 1) { $pathAndFlag[1] } else { '' }
                
                Write-Host $path -ForegroundColor Green -NoNewline
                if ($flag) {
                    Write-Host " " -NoNewline
                    Write-Host $flag -ForegroundColor Magenta -NoNewline
                }
            }
            
            if ($excPaths) {
                Write-Host "  excl: " -ForegroundColor DarkGray -NoNewline
                Write-Host ($excPaths -join ' | ') -ForegroundColor Red -NoNewline
            }
            Write-Host ""
        }
    }
    } ## end if ($Script:DebugCDP) — per-device summary
    Write-Output ""

    if ($Export -and $masterHashForUse.Count -gt 0) {
        ## Export is always the full master converted to analyst-friendly format
        $exportRows = @($masterHashForUse.Values |
            ForEach-Object { Convert-MasterRowToExportRow $_ } |
            Sort-Object AccountID)

        $exportRows | Export-Csv -Path $Script:csvoutputfile -NoTypeInformation -Encoding UTF8

        $xlsoutputfile = Save-CSVasExcel $Script:csvoutputfile

        Write-Output $Script:strLineSeparator
        Write-Host "  Exports ($($exportRows.Count) devices — full master in analyst format):" -ForegroundColor Cyan
        Write-Host "    CSV: $Script:csvoutputfile" -ForegroundColor Green
        Write-Host "    XLS: $xlsoutputfile" -ForegroundColor Green

        if ($Launch) {
            if (Test-Path HKLM:SOFTWARE\Classes\Excel.Application) {
                Start-Process $xlsoutputfile
            } else {
                Start-Process $Script:csvoutputfile
            }
        }
    }

    ## ---- Run summary ----
    $totalSelected   = $Script:SelectionResults.Count
    $reachedFirst    = @($Script:SelectionResults | Where-Object { $_.Reachable -eq 'Yes'      }).Count
    $reachedRetry    = @($Script:SelectionResults | Where-Object { $_.Reachable -eq 'Retry-OK' }).Count
    $failedFinal     = @($Script:SelectionResults | Where-Object { $_.Reachable -eq 'No'       }).Count
    $totalReached    = $reachedFirst + $reachedRetry
    $retryAttempted  = $retryBag.Count
    $overallPct      = if ($totalSelected  -gt 0) { '{0:0.0}' -f ($totalReached   * 100 / $totalSelected)  } else { 'N/A' }
    $retrySuccessPct = if ($retryAttempted -gt 0) { '{0:0.0}' -f ($reachedRetry   * 100 / $retryAttempted) } else { 'N/A' }
    $totalSelEntries = $Script:SelectionDetails.Count
    $avgPerDevice    = if ($totalReached   -gt 0) { '{0:0.0}' -f ($totalSelEntries * 1.0  / $totalReached)  } else { 'N/A' }

    Write-Output $Script:strLineSeparator
    Write-Host "  Run Summary — $Script:PartnerName" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator
    Write-Host "  Devices selected       : $totalSelected"
    Write-Host "  Reached (first pass)   : $reachedFirst"
    Write-Host "  Queued for retry       : $retryAttempted"
    if ($RetryFailed -and $retryAttempted -gt 0) {
        Write-Host "  Reached on retry       : $reachedRetry  ($retrySuccessPct% retry success)" -ForegroundColor $(if ($reachedRetry -gt 0) { 'Green' } else { 'Yellow' })
    }
    Write-Host "  Still unreachable      : $failedFinal"   -ForegroundColor $(if ($failedFinal  -eq 0) { 'Green' } else { 'Red' })
    $reachColor = if ($overallPct -eq 'N/A') { 'Gray' } elseif ([double]$overallPct -ge 90) { 'Green' } elseif ([double]$overallPct -ge 70) { 'Yellow' } else { 'Red' }
    Write-Host "  Overall reachability   : $totalReached / $totalSelected  ($overallPct%)" -ForegroundColor $reachColor
    Write-Host "  Total selection entries: $totalSelEntries  (avg $avgPerDevice per reachable device)"
    if ($Export) {
        Write-Host "  Output CSV             : $Script:csvoutputfile" -ForegroundColor DarkCyan
    }

    ## ---- Elapsed time breakdown ----
    $fp  = if ($firstPassElapsed)  { $firstPassElapsed }  else { [TimeSpan]::Zero }
    $rp  = if ($retryElapsed)      { $retryElapsed }      else { [TimeSpan]::Zero }
    $tot = $scriptSw.Elapsed

    $fpReached    = $reachedFirst
    $fpUnreached  = $retryAttempted
    $fpTotal      = $fpReached + $fpUnreached
    $fpAvgReached  = if ($fpReached   -gt 0) { [int]($fp.TotalSeconds / $fpReached   * 1000) } else { 0 }
    $fpAvgMissed   = if ($fpUnreached -gt 0) { [int]($fp.TotalSeconds / $fpUnreached * 1000) } else { 0 }

    $noVisaCount  = @($perfBag | Where-Object { $_.Pass -eq 'First' -and $_.S2_ms -ge 85000 }).Count  ## S2 near 90s timeout
    $noVisaRate   = if ($fpTotal -gt 0) { '{0:0.0}' -f ($noVisaCount * 100.0 / $fpTotal) } else { '0.0' }

    ## Throttle recommendation based on error rate
    $recommendedThrottle = switch ([double]$noVisaRate) {
        { $_ -gt 20 } { 10 }
        { $_ -gt 10 } { 15 }
        { $_ -gt 5  } { 20 }
        { $_ -gt 2  } { 30 }
        default       { 40 }
    }

    Write-Output $Script:strLineSeparator
    Write-Host "  Elapsed Time — Throttle: $DeviceThrottle" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator
    Write-Host ("  First pass   : {0:mm\:ss\.f}  ({1} reached ~{2}ms avg  |  {3} missed ~{4}ms avg)" -f $fp, $fpReached, $fpAvgReached, $fpUnreached, $fpAvgMissed)
    if ($rp.TotalSeconds -gt 0) {
        Write-Host ("  Retry pass   : {0:mm\:ss\.f}  ({1} reached  |  {2} still unreachable)" -f $rp, $reachedRetry, $failedFinal)
    }
    Write-Host ("  Total script : {0:mm\:ss\.f}" -f $tot)
    Write-Host ("  No-visa errors (Stage 2 timeout) : ~$noVisaCount / $fpTotal  ($noVisaRate%)")
    $throttleColor = if ($recommendedThrottle -lt $DeviceThrottle) { 'Yellow' } else { 'Green' }
    Write-Host "  Recommended throttle   : $recommendedThrottle  (current: $DeviceThrottle)" -ForegroundColor $throttleColor
    Write-Output $Script:strLineSeparator

    ## ---- Device configuration analysis ----
    Write-Output $Script:strLineSeparator
    Write-Host "  Device Configuration Analysis" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator

    ## Profile vs no-profile count
    $profileEntries   = @($Script:SelectionDetails | Where-Object { $_.Flags -like '*CreatedByAccountProfile*' })
    $nonProfileCount  = @($Script:SelectionResults | Where-Object { $_.Reachable -ne 'No' }).Count
    $profileDevCount  = @($profileEntries | Select-Object DeviceName -Unique).Count
    $nonProfileDevCount = $nonProfileCount - $profileDevCount

    Write-Host "  Profile-based devices : $profileDevCount" -ForegroundColor $(if ($profileDevCount -gt 0) { 'Green' } else { 'Yellow' })
    Write-Host "  Manual selections     : $nonProfileDevCount"
    Write-Output ""

    ## Unique datasource configurations and device counts
    $dsConfigs = @{}
    foreach ($dev in @($Script:SelectionResults | Where-Object { $_.Reachable -ne 'No' })) {
        $activeDs = $dev.ActiveDS
        if (-not $dsConfigs.ContainsKey($activeDs)) {
            $dsConfigs[$activeDs] = @()
        }
        $dsConfigs[$activeDs] += $dev
    }

    Write-Host "  Unique datasource configurations: $($dsConfigs.Count)" -ForegroundColor Cyan
    $dsConfigs.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object {
        $config = $_.Key
        $devList = $_.Value
        $percent = '{0:0.0}' -f ($devList.Count * 100 / $nonProfileCount)
        Write-Host "    [$($devList.Count) devices]  ($percent%)  $config" -ForegroundColor DarkCyan
        
        ## Sub-breakdown by OS type (Windows, macOS, Linux, Other)
        $osByType = @{}
        foreach ($dev in $devList) {
            $os = $dev.OS
            $osType = if ($os -like '*Windows*' -or $os -like '*Server*') { 'Windows' }
                      elseif ($os -like '*macOS*' -or $os -like '*Mac OS*') { 'macOS' }
                      elseif ($os -like '*Linux*' -or $os -like '*Ubuntu*' -or $os -like '*CentOS*' -or $os -like '*Red Hat*') { 'Linux' }
                      else { 'Other' }
            if (-not $osByType.ContainsKey($osType)) { $osByType[$osType] = 0 }
            $osByType[$osType]++
        }
        
        ## Only show OS breakdown if there's diversity (not all Windows)
        if ($osByType.Count -gt 1 -or ($osByType.ContainsKey('macOS') -or $osByType.ContainsKey('Linux') -or $osByType.ContainsKey('Other'))) {
            foreach ($osType in @('Windows', 'macOS', 'Linux', 'Other') | Where-Object { $osByType.ContainsKey($_) }) {
                $count = $osByType[$osType]
                $osPct = '{0:0.0}' -f ($count * 100 / $devList.Count)
                Write-Host "      ├─ $osType`: $count ($osPct%)" -ForegroundColor Gray
            }
        }
    }
    Write-Output ""

    ## Devices with limited/non-standard selections (less than median per datasource)
    $selCountByDsAndDev = @{}
    foreach ($detail in $Script:SelectionDetails) {
        $key = "$($detail.DeviceName)::$($detail.DataSource)"
        if (-not $selCountByDsAndDev.ContainsKey($key)) {
            $selCountByDsAndDev[$key] = 0
        }
        ## Inclusive selections with any path (including legacy \ or System State) count as a selection
        if ($detail.Type -eq 'Inclusive' -and ($detail.Path -or $detail.Path -eq '\' -or $detail.Path -eq 'System State')) {
            $selCountByDsAndDev[$key]++
        }
    }

    ## Find median and outliers per datasource
    $outlierDevices = @{}
    foreach ($dsPlugin in @($Script:DataSources.Keys)) {
        $dsAbbrev = $Script:DataSources[$dsPlugin]
        $countsForDs = @($selCountByDsAndDev.GetEnumerator() |
            Where-Object { $_.Key -like "*::$dsPlugin" } |
            ForEach-Object { $_.Value })

        if ($countsForDs.Count -gt 0) {
            $sorted = @($countsForDs | Sort-Object)
            $median = if ($sorted.Count % 2 -eq 0) {
                ($sorted[[int]($sorted.Count/2)-1] + $sorted[[int]($sorted.Count/2)]) / 2
            } else {
                $sorted[[int]($sorted.Count/2)]
            }

            ## Flag devices with less than 50% of median selections
            foreach ($kvp in $selCountByDsAndDev.GetEnumerator()) {
                if ($kvp.Key -like "*::$dsPlugin" -and $kvp.Value -gt 0 -and $kvp.Value -lt ($median * 0.5)) {
                    $devName = $kvp.Key.Split('::')[0]
                    if (-not $outlierDevices.ContainsKey($devName)) {
                        $outlierDevices[$devName] = @()
                    }
                    $outlierDevices[$devName] += "$($dsAbbrev):$($kvp.Value) sel (median:$([int]$median))"
                }
            }
        }
    }

    if ($outlierDevices.Count -gt 0) {
        Write-Host "  Devices with non-standard (limited) selections:" -ForegroundColor Yellow
        $outlierDevices.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object {
            Write-Host "    $($_.Key)" -ForegroundColor Yellow
            $_.Value | ForEach-Object { Write-Host "      └─ $_" -ForegroundColor DarkYellow }
        }
    } else {
        Write-Host "  All devices have standard selection counts" -ForegroundColor Green
    }
    Write-Output ""

    ## Recommended datasource column order
    $dsFreq = @{}
    foreach ($detail in $Script:SelectionDetails) {
        if (-not $dsFreq.ContainsKey($detail.DataSource)) {
            $dsFreq[$detail.DataSource] = 0
        }
        $dsFreq[$detail.DataSource]++
    }

    ## Suggested order: mandatory > common > windows-specific > linux-specific > niche
    $suggestedOrder = @(
        'FileSystem',      ## Always present, mandatory
        'VssSystemState', 'SystemState', 'LinuxSystemState',  ## OS state (windows/mac/linux)
        'VssExchange',     ## Common enterprise
        'VssMsSql',        ## Common enterprise
        'VssHyperV',       ## Common enterprise
        'VMWare',          ## Virtualization
        'VssSharePoint',   ## Office 365 adjacent
        'NetworkShares',   ## File services
        'MySql',           ## Databases
        'Oracle'           ## Enterprise databases
    )

    Write-Host "  Recommended datasource column order (in export):" -ForegroundColor Cyan
    $idx = 1
    foreach ($ds in $suggestedOrder) {
        $abbrev = if ($Script:DataSources[$ds]) { $Script:DataSources[$ds] } else { $ds }
        $freq   = if ($dsFreq[$ds]) { $dsFreq[$ds] } else { 0 }
        $pct    = if ($totalSelEntries -gt 0) { '{0:0.0}' -f ($freq * 100 / $totalSelEntries) } else { '0.0' }
        $hasData = if ($freq -gt 0) { '✓' } else { '○' }
        Write-Host "    $idx. [$abbrev]  $ds  $hasData  ($freq entries, $pct%)" -ForegroundColor DarkCyan
        $idx++
    }
    Write-Output $Script:strLineSeparator

    ## ---- Selection pattern analysis ----
    Write-Output $Script:strLineSeparator
    Write-Host "  Selection Pattern Analysis" -ForegroundColor Cyan
    Write-Output $Script:strLineSeparator

    ## "All" vs path-specific selections per datasource
    $allVsPathByDs = @{}
    foreach ($detail in $Script:SelectionDetails | Where-Object { $_.Type -eq 'Inclusive' }) {
        $ds = $detail.DataSource
        if (-not $allVsPathByDs.ContainsKey($ds)) {
            $allVsPathByDs[$ds] = @{ All = 0; Specific = 0 }
        }
        ## Normalize: empty, \, "System State" all mean full backup (all)
        $isFullBackup = ($detail.Path -eq $null -or $detail.Path -eq '' -or $detail.Path -eq '(all)' -or $detail.Path -eq '\' -or $detail.Path -eq 'System State')
        if ($isFullBackup) {
            $allVsPathByDs[$ds].All++
        } else {
            $allVsPathByDs[$ds].Specific++
        }
    }

    Write-Host "  Backup scope per datasource (All vs. Specific paths):" -ForegroundColor Cyan
    foreach ($ds in ($suggestedOrder | Where-Object { $allVsPathByDs.ContainsKey($_) })) {
        $abbrev = $Script:DataSources[$ds]
        $all    = $allVsPathByDs[$ds].All
        $spec   = $allVsPathByDs[$ds].Specific
        $total  = $all + $spec
        if ($total -gt 0) {
            $allPct  = '{0:0.0}' -f ($all  * 100 / $total)
            $specPct = '{0:0.0}' -f ($spec * 100 / $total)
            Write-Host "    [$abbrev]  Full backup: $all ($allPct%)  |  Path-specific: $spec ($specPct%)" -ForegroundColor DarkCyan
        }
    }
    Write-Output ""

    ## Exclusion analysis
    $exclusiveCount = @($Script:SelectionDetails | Where-Object { $_.Type -eq 'Exclusive' -and $_.Path }).Count
    $inclusiveCount = @($Script:SelectionDetails | Where-Object { $_.Type -eq 'Inclusive' }).Count
    $devicesWithExclusions = @($Script:SelectionDetails | Where-Object { $_.Type -eq 'Exclusive' -and $_.Path } | Select-Object DeviceName -Unique).Count
    $exclusionPct = if ($inclusiveCount -gt 0) { '{0:0.0}' -f ($exclusiveCount * 100 / $inclusiveCount) } else { '0.0' }

    Write-Host "  Exclusion Coverage:" -ForegroundColor Cyan
    $exclPct = if ($nonProfileCount -gt 0) { '{0:0.0}' -f ($devicesWithExclusions * 100 / $nonProfileCount) } else { 'N/A' }
    Write-Host "    Devices with exclusions: $devicesWithExclusions / $nonProfileCount  ($exclPct%)" -ForegroundColor DarkCyan
    Write-Host "    Exclusion/Inclusion ratio: $exclusionPct%" -ForegroundColor DarkCyan
    Write-Output ""

    ## Orphaned datasources (enabled in I78 but no selections)
    $orphanedByDevice = @{}
    foreach ($dev in @($Script:SelectionResults | Where-Object { $_.Reachable -ne 'No' })) {
        if ($dev.ActiveDS) {
            $enabledDsPlugins = $dev.ActiveDS -split '; ' | Where-Object { $_ }
            foreach ($abbr in $enabledDsPlugins) {
                $dsPlugin = $Script:DataSources.GetEnumerator() | Where-Object { $_.Value -eq $abbr } | Select-Object -First 1 -ExpandProperty Key
                if ($dsPlugin) {
                    $hasSelections = @($Script:SelectionDetails | Where-Object { $_.DeviceName -eq $dev.DeviceName -and $_.DataSource -eq $dsPlugin }).Count -gt 0
                    if (-not $hasSelections) {
                        if (-not $orphanedByDevice.ContainsKey($dev.DeviceName)) {
                            $orphanedByDevice[$dev.DeviceName] = @()
                        }
                        $orphanedByDevice[$dev.DeviceName] += $abbr
                    }
                }
            }
        }
    }

    if ($orphanedByDevice.Count -gt 0) {
        Write-Host "  ⚠ Orphaned datasources (enabled but no selections configured):" -ForegroundColor Yellow
        $orphanedByDevice.GetEnumerator() | ForEach-Object {
            Write-Host "    $($_.Key): $($_.Value -join ', ')" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ✓ No orphaned datasources found" -ForegroundColor Green
    }
    Write-Output ""

    ## Selection consistency — identify identical configs and high-variance datasources
    $configSignatures = @{}  ## device => config hash (simplified)
    foreach ($dev in @($Script:SelectionResults | Where-Object { $_.Reachable -ne 'No' })) {
        $sig = $dev.ActiveDS
        if (-not $configSignatures.ContainsKey($sig)) {
            $configSignatures[$sig] = @()
        }
        $configSignatures[$sig] += $dev.DeviceName
    }

    $mostCommonConfig = $configSignatures.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | Select-Object -First 1
    if ($mostCommonConfig) {
        $commonCount = $mostCommonConfig.Value.Count
        $commonPct = '{0:0.0}' -f ($commonCount * 100 / $nonProfileCount)
        Write-Host "  Configuration standardization:" -ForegroundColor Cyan
        Write-Host "    Most common datasource config: $($mostCommonConfig.Key)" -ForegroundColor DarkCyan
        Write-Host "    Devices sharing this config: $commonCount / $nonProfileCount  ($commonPct%)" -ForegroundColor DarkCyan
        if ($commonPct -lt 30) {
            Write-Host "    → Recommendation: High variability detected. Consider standardizing on 1-2 templates." -ForegroundColor Yellow
        } elseif ($commonPct -ge 70) {
            Write-Host "    → Recommendation: Good standardization. Most devices follow common pattern." -ForegroundColor Green
        }
    }
    Write-Output ""

    ## Datasource prevalence (coverage %) — grouped by abbreviation to avoid duplicate SysState variants
    $reachableCount = @($Script:SelectionResults | Where-Object { $_.Reachable -ne 'No' }).Count
    Write-Host "  Datasource Prevalence (% of reachable devices):" -ForegroundColor Cyan
    
    ## Get unique abbreviations in order (FS, SysState, MSSQL, MySQL, HyperV, NetShares, Exchange, VMware, SharePt, Oracle)
    $uniqueAbbrevs = @()
    foreach ($ds in $suggestedOrder) {
        $abbrev = $Script:DataSources[$ds]
        if ($abbrev -notin $uniqueAbbrevs) {
            $uniqueAbbrevs += $abbrev
        }
    }
    
    foreach ($abbrev in $uniqueAbbrevs) {
        $devicesWithDs = @($Script:SelectionResults |
            Where-Object { $_.Reachable -ne 'No' -and $_.ActiveDS -like "*$abbrev*" }).Count
        if ($devicesWithDs -gt 0) {
            $dsPercent = '{0:0.0}' -f ($devicesWithDs * 100 / $reachableCount)
            Write-Host "    $abbrev`: $devicesWithDs devices  ($dsPercent%)" -ForegroundColor DarkCyan
        }
    }
    Write-Output $Script:strLineSeparator

    Write-Output $Script:strLineSeparator

    ## ---- Flags observed across all selection entries ----
    $knownFlags   = @('CreatedByAccountProfile')  ## expected flags
    $allFlagsSeen = $Script:SelectionDetails |
        Where-Object { $_.Flags } |
        ForEach-Object { $_.Flags -split ',' } |
        Where-Object { $_ } |
        Sort-Object -Unique

    ## Report only if unexpected flags appear
    $unexpected = $Script:SelectionDetails | Where-Object {
        if (-not $_.Flags) { return $false }
        $flagList = @($_.Flags -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        @($flagList | Where-Object { $_ -notin $knownFlags }).Count -gt 0
    }
    if ($unexpected) {
        Write-Output $Script:strLineSeparator
        Write-Host "  ⚠️  Unexpected Selection Flags Found" -ForegroundColor Yellow
        Write-Output $Script:strLineSeparator
        $unexpected | Select-Object DeviceName, DataSource, Type, Path, Flags |
            Format-Table -AutoSize | Out-String -Width 160 | Write-Host
    }

#endregion ----- Report ----

#region ----- Cleanup ----
    ## Release master CSV file lock (also released automatically by OS if process dies)
    if ($Script:MasterLockStream) {
        $Script:MasterLockStream.Close()
        $Script:MasterLockStream.Dispose()
        $Script:MasterLockStream = $null
        $lockPath = (Get-MasterCsvPath) + ".lock"
        if (Test-Path $lockPath) { Remove-Item $lockPath -Force -ErrorAction SilentlyContinue }
    }
#endregion ----- Cleanup ----
