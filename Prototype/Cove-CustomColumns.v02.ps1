<# ----- About: ----
    # Cove Data Protection - Custom Column Manager
    # Revision v02 - 2026-05-15
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Authenticates to the Cove Data Protection API
    # Iterates over one or more partners (via -PartnerList, -PartnerName, or built-in default list)
    # For each partner, enumerates all devices that have at least one custom column value set
    # Pivots results into wide format — one row per device, one column per custom column definition
    # Exports a single combined XLSX workbook covering all partners
    #
    # OUTPUT COLUMNS:
    #   PartnerName | AccountId | DeviceId | DeviceName | ComputerName | DeviceType | [CustomCol1] | [CustomCol2] | ...
    #
    # PARAMETERS:
    #   -PartnerList      : String array of partner names to process
    #                       (default: 7 Sika LATAM partners defined in $Script:DefaultPartnerList)
    #   -PartnerName      : Single partner name (used when -PartnerList is not provided)
    #   -ClearCredentials : Remove stored API credentials and re-prompt
    #   -Launch           : Open the exported XLSX automatically (default: $true)
    #   -ExportPath       : Override export folder (default: script directory)
    #
    # REQUIRES: ImportExcel module (installed automatically for current user if absent)
    #
    # API DOCUMENTATION:
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
# -----------------------------------------------------------#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)] [string[]]$PartnerList,       ## Array of partner names to iterate
    [Parameter(Mandatory=$False)] [string]$PartnerName,         ## Single partner (used if PartnerList absent)
    [Parameter(Mandatory=$False)] [switch]$ClearCredentials,    ## Remove stored credentials at script start
    [Parameter(Mandatory=$False)] [switch]$Launch = $true,      ## Open exported XLSX when complete
    [Parameter(Mandatory=$False)] [string]$ExportPath,          ## Override export folder (default: script dir)
    [Parameter(Mandatory=$False)] [string]$DeviceFilter = 'OT == 2',   ## EnumerateAccountStatistics filter e.g. 'OT == 2'
    [Parameter(Mandatory=$False)] [switch]$Move,                ## After export: show gridview to move devices to target partners
    [Parameter(Mandatory=$False)] [switch]$Revert               ## After export: show gridview to revert previously moved devices
)

#region ----- Environment, Variables, Names and Paths ----

    Set-Location $PSScriptRoot
    Clear-Host

    $ConsoleTitle = "Cove Data Protection - Custom Column Manager"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Script:strLineSeparator = "  ---------"
    $Script:urlJSON          = 'https://api.backup.management/jsonapi'
    $Script:Timestamp        = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Script:True_path        = "C:\ProgramData\MXB\"
    $Script:APIcredfile      = Join-Path -Path $Script:True_path -ChildPath "${env:computername}_${env:username}_API_Credentials.Secure.xml"
    $Script:APIcredpath      = Split-Path -Path $Script:APIcredfile
    $Script:VisaLoginTime    = $null

    ## Move log — persistent CSV recording all moves and reverts
    $Script:MoveLogFile = Join-Path $PSScriptRoot 'DeviceMoveLog.csv'

    ## Default partner list — override with -PartnerList or -PartnerName at runtime
    $Script:DefaultPartnerList = @(
        'CL - Sika S.A. Chile 6605',
        'DR - Sika Dominicana SRL 6709',
        'EC - Sika Ecuatoriana S.A 6703',
        'AR - Sika Argentina SAIC 6601',
        'CO - Sika Colombia S.A. 6701',
        'PE - Sika Peru SA 6606',
        'PY - Sika Paraguay SA (Inatec) 6608'
    )

    ## Source partner ID → [TargetPartnerId, TargetPartnerName] mapping
    $Script:TargetPartnerMap = @{
        264142 = @{ Id = 2870606; Name = 'CL - Sika S.A. Chile_6605'           }
        264189 = @{ Id = 2869960; Name = 'DR - Sika Dominicana SRL_6709'        }
        266853 = @{ Id = 2870561; Name = 'EC - Sika Ecuatoriana S.A_6703'       }
        253738 = @{ Id = 2870600; Name = 'AR - Sika Argentina SAIC_6601'        }
        264168 = @{ Id = 2869958; Name = 'CO - Sika Colombia S.A._6701'         }
        232169 = @{ Id = 2870608; Name = 'PE - Sika Peru SA_6606'               }
        263877 = @{ Id = 2870607; Name = 'PY - Sika Paraguay SA (Inatec)_6608'  }
    }

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

    #region ----- Authentication ----

    Function Set-APICredentials {
        Write-Output $Script:strLineSeparator
        Write-Output "  Setting Backup API Credentials"

        if (Test-Path $Script:APIcredpath) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential Path Present"
        } else {
            New-Item -ItemType Directory -Path $Script:APIcredpath -Force | Out-Null
        }

        Write-Output "  Enter Exact, Case Sensitive Partner Name for Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO { $Script:cred0 = Read-Host "  Enter Login Partner Name" }
        WHILE ($Script:cred0.length -eq 0)

        $BackupCred = Get-Credential -Message 'Enter Login Email and Password for Backup.Management API'

        $CDPCredentials = [PSCustomObject]@{
            PartnerName = $Script:cred0
            Username    = $BackupCred.UserName
            Password    = ($BackupCred.Password | ConvertFrom-SecureString)
        }

        $CDPCredentials | Export-Clixml -Path $Script:APIcredfile -Force
        Write-Output "  Credentials saved to: $Script:APIcredfile"

        Start-Sleep -Milliseconds 300
        Send-APICredentialsCookie
    }  ## Set API credentials if not present

    Function Get-APICredentials {
        Write-Output $Script:strLineSeparator
        Write-Output "  Getting Backup API Credentials"

        if (($ClearCredentials) -and (Test-Path $Script:APIcredfile)) {
            Remove-Item -Path $Script:APIcredfile -Force
            $Script:ClearCredentials = $Null
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Cleared"
            Send-APICredentialsCookie
        } else {
            if (Test-Path $Script:APIcredfile) {
                Write-Output $Script:strLineSeparator
                Write-Output "  Backup API Credential File Present"

                try {
                    $APIcredentials = Import-Clixml -Path $Script:APIcredfile -ErrorAction Stop

                    $Script:cred0 = $APIcredentials.PartnerName
                    $Script:cred1 = $APIcredentials.Username
                    $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
                                    ($APIcredentials.Password | ConvertTo-SecureString)))

                    Write-Output "  Stored Backup API Partner  = $Script:cred0"
                    Write-Output "  Stored Backup API User     = $Script:cred1"
                    Write-Output "  Stored Backup API Password = Encrypted"
                    Write-Output $Script:strLineSeparator
                } catch {
                    Write-Output $Script:strLineSeparator
                    Write-Warning "  Backup API Credential File is corrupted or invalid"
                    Write-Output "  Error: $($_.Exception.Message)"
                    Write-Output "  Removing corrupted file and prompting for new credentials..."
                    Remove-Item -Path $Script:APIcredfile -Force -ErrorAction SilentlyContinue
                    Set-APICredentials
                }
            } else {
                Write-Output $Script:strLineSeparator
                Write-Output "  Backup API Credential File Not Present"
                Set-APICredentials
            }
        }
    }  ## Get API credentials if present

    Function Send-APICredentialsCookie {
        Get-APICredentials

        Write-Output $Script:strLineSeparator
        Write-Output "  Authenticating against Backup.Management API"
        Write-Output $Script:strLineSeparator

        $data = @{
            jsonrpc = '2.0'
            id      = '2'
            method  = 'Login'
            params  = @{
                partner  = $Script:cred0
                username = $Script:cred1
                password = $Script:cred2
            }
        }

        try {
            $Script:Authenticate = Invoke-RestMethod -Method POST `
                -ContentType 'application/json' `
                -Body (ConvertTo-Json $data -Depth 5) `
                -Uri $Script:urlJSON `
                -TimeoutSec 30 `
                -SessionVariable Script:websession `
                -UseBasicParsing

            $Script:cookies = $Script:websession.Cookies.GetCookies($Script:urlJSON)

            if ($Script:Authenticate.error) {
                Write-Output $Script:strLineSeparator
                Write-Output "  Authentication API Error: $($Script:Authenticate.error.message)"
                Write-Output $Script:strLineSeparator
                Set-APICredentials
                return
            }

            if ($Script:Authenticate.visa) {
                $Script:visa          = $Script:Authenticate.visa
                $Script:VisaLoginTime = Get-Date
                Write-Output "  Authentication Successful"
                Write-Output $Script:strLineSeparator
            } else {
                Write-Output $Script:strLineSeparator
                Write-Output "  Authentication Failed: No visa token received"
                Write-Output "  Please confirm your Backup.Management Partner Name and Credentials"
                Write-Output "  Note: Multiple failed attempts could temporarily lock your account"
                Write-Output $Script:strLineSeparator
                Set-APICredentials
            }
        } catch {
            Write-Output $Script:strLineSeparator
            Write-Output "  Authentication Network Error: $($_.Exception.Message)"
            Write-Output $Script:strLineSeparator
            Set-APICredentials
        }
    }  ## Use credentials to authenticate and get visa token

    Function Invoke-RefreshVisaIfNeeded {
        ## Re-authenticate using cached credentials if the visa is older than 50 minutes
        if ($null -eq $Script:VisaLoginTime -or (Get-Date) -gt $Script:VisaLoginTime.AddMinutes(50)) {
            Write-Host "  Refreshing visa token (session age > 50 min)..."
            Send-APICredentialsCookie
        }
    }  ## Silently refresh visa token before it expires

    #endregion ----- Authentication ----

    #region ----- Partner Lookup ----

    Function Send-GetPartnerInfo ([string]$LookupName) {
        $data = @{
            jsonrpc = '2.0'
            id      = '2'
            visa    = $Script:visa
            method  = 'GetPartnerInfo'
            params  = @{ name = $LookupName }
        }

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) `
            -Uri $Script:urlJSON `
            -SessionVariable Script:websession `
            -UseBasicParsing

        $Script:Partner = $webrequest | ConvertFrom-Json

        if ($Script:Partner.result.result.Uid) {
            $Script:PartnerId   = [int]$Script:Partner.result.result.Id
            $Script:PartnerName = $Script:Partner.result.result.Name
            $Script:Level       = $Script:Partner.result.result.Level

            Write-Output $Script:strLineSeparator
            Write-Output "  Partner Level : $Script:Level"
            Write-Output "  Partner Name  : $Script:PartnerName"
            Write-Output "  Partner ID    : $Script:PartnerId"
            Write-Output $Script:strLineSeparator
        } else {
            Write-Warning "  Partner '$LookupName' not found"
            $Script:PartnerId   = $null
            $Script:PartnerName = $null
            $Script:Level       = $null
        }
    }  ## Resolve partner name to partner ID

    #endregion ----- Partner Lookup ----

    #region ----- Custom Column API Functions ----

    Function Invoke-EnumerateCustomColumns {
        Param (
            [Parameter(Mandatory=$False)] [int]$PartnerId = $Script:PartnerId
        )

        $data = @{
            jsonrpc = '2.0'
            id      = '2'
            visa    = $Script:visa
            method  = 'EnumerateCustomColumns'
            params  = @{ partnerId = $PartnerId }
        }

        $response = Invoke-WebRequest -Method POST `
            -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) `
            -Uri $Script:urlJSON `
            -SessionVariable Script:websession `
            -UseBasicParsing | ConvertFrom-Json

        if ($response.error) {
            Write-Warning "  EnumerateCustomColumns Error: $($response.error.message)"
            return $null
        }

        return $response.result.result
    }  ## List all custom columns for a partner

    Function Invoke-GetAccountCustomColumnValues {
        Param (
            [Parameter(Mandatory=$True)] [int]$AccountId
        )

        $data = @{
            jsonrpc = '2.0'
            id      = '2'
            visa    = $Script:visa
            method  = 'GetAccountCustomColumnValues'
            params  = @{ accountId = $AccountId }
        }

        $response = Invoke-WebRequest -Method POST `
            -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) `
            -Uri $Script:urlJSON `
            -SessionVariable Script:websession `
            -UseBasicParsing | ConvertFrom-Json

        if ($response.error) {
            Write-Warning "  GetAccountCustomColumnValues Error: $($response.error.message)"
            return $null
        }

        return $response.result.result
    }  ## Get custom column values for a device

    Function Invoke-EnumerateAccountProfiles {
        Param (
            [Parameter(Mandatory=$True)] [int]$PartnerId
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'EnumerateAccountProfiles'
            params  = @{ partnerId = $PartnerId }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  EnumerateAccountProfiles Error: $($response.error.message)"; return @() }
        return $response.result.result
    }  ## List all profiles for a partner

    Function Invoke-GetAccountProfileInfo {
        Param (
            [Parameter(Mandatory=$True)] [int]$ProfileId
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'GetAccountProfileInfo'
            params  = @{ accountProfileId = $ProfileId }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 3) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  GetAccountProfileInfo Error: $($response.error.message)"; return $null }
        return $response.result.result
    }  ## Get full profile definition by ID

    Function Invoke-AddAccountProfile {
        Param (
            [Parameter(Mandatory=$True)] $ProfileInfo
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'AddAccountProfile'
            params  = @{ accountProfileInfo = $ProfileInfo }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 100) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  AddAccountProfile Error: $($response.error.message)"; return $null }
        return $response.result.result
    }  ## Create a new profile; returns new profile ID

    #endregion ----- Custom Column API Functions ----

    #region ----- Account Enumeration ----

    Function Invoke-EnumerateAccounts {
        Param (
            [Parameter(Mandatory=$False)] [int]$PartnerId    = $Script:PartnerId,
            [Parameter(Mandatory=$False)] [int]$RecordsCount = 5000,
            [Parameter(Mandatory=$False)] [string]$Filter    = ''
        )

        $data = @{
            jsonrpc = '2.0'
            id      = '2'
            visa    = $Script:visa
            method  = 'EnumerateAccountStatistics'
            params  = @{
                query = @{
                    PartnerId         = $PartnerId
                    Filter            = $Filter
                    Columns           = @('AU','AN','MN','OT','OI','OP','PD','PN')
                    OrderBy           = 'AN ASC'
                    StartRecordNumber = 0
                    RecordsCount      = $RecordsCount
                    Totals            = @('COUNT(AT==1)')
                }
            }
        }

        $response = Invoke-WebRequest -Method POST `
            -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 10) `
            -Uri $Script:urlJSON `
            -SessionVariable Script:websession `
            -UseBasicParsing | ConvertFrom-Json

        if ($response.error) {
            Write-Warning "  EnumerateAccountStatistics Error: $($response.error.message)"
            return $null
        }

        $accounts = @()
        foreach ($result in $response.result.result) {
            $otStr = ($result.Settings.OT -join '').Trim(); $otRaw = if ($otStr -match '\d+') { [int]$otStr } else { 0 }
            $accounts += [PSCustomObject]@{
                AccountId         = [int]($result.Settings.AU -join '')
                DeviceName        = ($result.Settings.AN -join '')
                ComputerName      = ($result.Settings.MN -join '')
                DeviceType        = switch ($otRaw) { 1 { 'Workstation' } 2 { 'Server' } default { 'Undefined' } }
                ProfileId         = ($result.Settings.OI -join '')
                ProfileName       = ($result.Settings.OP -join '')
                ProductId         = ($result.Settings.PD -join '')
                ProductName       = ($result.Settings.PN -join '')
            }
        }
        return $accounts
    }  ## Get all accounts for a partner

    #endregion ----- Account Enumeration ----

    #region ----- Product (Retention Policy) API Functions ----

    Function Invoke-EnumerateProducts {
        Param (
            [Parameter(Mandatory=$True)] [int]$PartnerId,
            [Parameter(Mandatory=$False)] [switch]$CurrentPartnerOnly
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'EnumerateProducts'
            params  = @{ partnerId = $PartnerId; currentPartnerOnly = [bool]$CurrentPartnerOnly; skipDefaultFeatures = $false }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  EnumerateProducts Error: $($response.error.message)"; return @() }
        return $response.result.result
    }  ## List all products/retention policies for a partner

    Function Invoke-GetProductInfo {
        Param (
            [Parameter(Mandatory=$True)] [int]$ProductId
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'GetProductInfo'
            params  = @{ productId = $ProductId; returnModifiedFeaturesOnly = $false }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  GetProductInfo Error: $($response.error.message)"; return $null }
        return $response.result.result
    }  ## Get full product/retention policy definition by ID

    Function Invoke-AddProduct {
        Param (
            [Parameter(Mandatory=$True)] $ProductInfo
        )
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'AddProduct'
            params  = @{ productInfo = $ProductInfo }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 20) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  AddProduct Error: $($response.error.message)"; return $null }
        return $response.result.result
    }  ## Create a new product/retention policy; returns new product ID

    #endregion ----- Product (Retention Policy) API Functions ----

    #region ----- Move API Functions ----

    Function Invoke-ModifyAccount {
        Param (
            [Parameter(Mandatory=$True)]  [int]$AccountId,
            [Parameter(Mandatory=$True)]  [int]$PartnerId,
            [Parameter(Mandatory=$False)] [int]$ProfileId = 0,
            [Parameter(Mandatory=$False)] [int]$ProductId = 0
        )
        $accountInfo = @{ Id = $AccountId; PartnerId = $PartnerId }
        if ($ProfileId -gt 0) { $accountInfo['ProfileId'] = $ProfileId }
        if ($ProductId -gt 0) { $accountInfo['ProductId'] = $ProductId }
        $data = @{
            jsonrpc = '2.0'; id = '2'; visa = $Script:visa
            method  = 'ModifyAccount'
            params  = @{ accountInfo = $accountInfo; forceRemoveCustomColumnValuesInOldScope = $true }
        }
        $response = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -Depth 5) -Uri $Script:urlJSON `
            -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
        if ($response.error) { Write-Warning "  ModifyAccount Error: $($response.error.message)"; return $false }
        return $true
    }  ## Move a device to a new partner/profile/product; returns $true on success

    #endregion ----- Move API Functions ----

#endregion ----- Functions ----

#region ----- Main Execution ----

    ## Authenticate
    Send-APICredentialsCookie

    if (-not $Script:visa) {
        Write-Error "  Authentication failed. Exiting."
        Exit 1
    }

    ## Ensure ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "  ImportExcel module not found — installing for current user..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel -ErrorAction Stop

    ## Resolve the partner list to process
    $resolvedList = if ($PartnerList -and $PartnerList.Count -gt 0) {
        $PartnerList
    } elseif ($PartnerName) {
        @($PartnerName)
    } else {
        $Script:DefaultPartnerList
    }

    Write-Host "`n$Script:strLineSeparator"
    Write-Host "  Partners to process : $($resolvedList.Count)"
    $resolvedList | ForEach-Object { Write-Host "    - $_" }
    Write-Host $Script:strLineSeparator

    ## Accumulate flat rows and track all distinct column names across all partners
    $allFlatRows = [System.Collections.Generic.List[PSObject]]::new()
    $allColNames = [System.Collections.Generic.SortedSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($pName in $resolvedList) {

        Invoke-RefreshVisaIfNeeded

        Write-Host "`n$Script:strLineSeparator"
        Write-Host "  Processing: $pName"
        Write-Host $Script:strLineSeparator

        Send-GetPartnerInfo $pName

        if (-not $Script:PartnerId) {
            Write-Warning "  Skipping '$pName' — partner not found."
            continue
        }

        ## Custom column definitions for this partner
        $columnDefs = Invoke-EnumerateCustomColumns

        if (-not $columnDefs -or $columnDefs.Count -eq 0) {
            Write-Host "  No custom columns defined for '$Script:PartnerName' — skipping."
            continue
        }

        $colLookup = @{}
        foreach ($col in $columnDefs) {
            $colLookup[[int]$col.Id] = $col.Name
            $allColNames.Add($col.Name) | Out-Null
        }
        Write-Host "  Custom columns : $($columnDefs.Count)  ($($columnDefs.Name -join ', '))"

        ## All device accounts for this partner
        $accounts = Invoke-EnumerateAccounts -Filter $DeviceFilter

        if (-not $accounts -or $accounts.Count -eq 0) {
            Write-Host "  No accounts found for '$Script:PartnerName' — skipping."
            continue
        }
        Write-Host "  Accounts       : $($accounts.Count)"

        ## Per-device: fetch values; include ALL devices, empty values if none set
        $i          = 0
        $withValues = 0

        foreach ($account in $accounts) {
            $i++
            Write-Progress -Activity "Fetching column values — $Script:PartnerName" `
                -Status "$i / $($accounts.Count): $($account.DeviceName)" `
                -PercentComplete (($i / $accounts.Count) * 100)

            $valueMap = Invoke-GetAccountCustomColumnValues -AccountId $account.AccountId

            $deviceValues = @{}
            if ($valueMap -is [System.Object[]] -and $valueMap.Count -gt 0) {
                foreach ($pair in $valueMap) {
                    $colId   = [int]$pair[0]
                    $colName = $colLookup[$colId]
                    if ($colName) { $deviceValues[$colName] = [string]$pair[1] }
                }
                $withValues++
            }

            $allFlatRows.Add([PSCustomObject]@{
                PartnerName     = $Script:PartnerName
                SourcePartnerId = $Script:PartnerId
                AccountId       = $account.AccountId
                DeviceId        = $account.AccountId
                DeviceName      = $account.DeviceName
                ComputerName    = $account.ComputerName
                DeviceType      = $account.DeviceType
                ProfileId       = $account.ProfileId
                ProfileName     = $account.ProfileName
                ProductId       = $account.ProductId
                ProductName     = $account.ProductName
                _Values         = $deviceValues
            }) | Out-Null
        }
        Write-Progress -Activity "Fetching column values — $Script:PartnerName" -Completed
        Write-Host "  Devices with values: $withValues / $($accounts.Count)"
    }

    ## Clone unique source profiles to target partners (skip if profile of same name already exists)
    $Script:ProfileCloneMap = @{}  ## key = "srcPartnerId|srcProfileId", value = @{Id=...; Name=...}

    Write-Host "`n$Script:strLineSeparator"
    Write-Host "  Cloning profiles to target partners..."

    $allFlatRows | Group-Object SourcePartnerId | ForEach-Object {
        $srcPartnerId = [int]$_.Name
        $tgt = $Script:TargetPartnerMap[$srcPartnerId]
        if (-not $tgt) {
            Write-Host "  No target mapping for source partner $srcPartnerId — skipping"
            return
        }

        Write-Host "  $srcPartnerId → Target: $($tgt.Id) ($($tgt.Name))"
        Invoke-RefreshVisaIfNeeded

        ## Build lookup of names already in target partner
        $existingProfiles = Invoke-EnumerateAccountProfiles -PartnerId $tgt.Id
        $existingNames = @{}
        foreach ($ep in $existingProfiles) { $existingNames[$ep.Name.ToLower()] = [int]$ep.Id }

        ## Unique non-zero ProfileIds from devices belonging to this source partner
        $uniqueProfileIds = $_.Group |
            Where-Object { $_.ProfileId -and $_.ProfileId -notin @('','0') } |
            Select-Object -ExpandProperty ProfileId -Unique

        Write-Host "    Source profile IDs to process: $(if ($uniqueProfileIds) { $uniqueProfileIds -join ', ' } else { '(none)' })"

        foreach ($srcProfileId in $uniqueProfileIds) {
            $key = "$srcPartnerId|$srcProfileId"

            $profileInfo = Invoke-GetAccountProfileInfo -ProfileId ([int]$srcProfileId)
            if (-not $profileInfo) { Write-Warning "    Could not retrieve profile $srcProfileId"; continue }

            $profileName = $profileInfo.Name

            if ($existingNames.ContainsKey($profileName.ToLower())) {
                $existingId = $existingNames[$profileName.ToLower()]
                Write-Host "    '$profileName' already exists in target (ID: $existingId) — reusing"
                $Script:ProfileCloneMap[$key] = @{ Id = $existingId; Name = $profileName }
                continue
            }

            ## Strip version + ID, remap PartnerId to target, then create
            $profileInfo.PSObject.Properties.Remove('Version')
            $profileInfo.PSObject.Properties.Remove('Id')
            $profileInfo.PartnerId = $tgt.Id

            $newId = Invoke-AddAccountProfile -ProfileInfo $profileInfo
            if ($newId) {
                Write-Host "    '$profileName' cloned → new ID: $newId" -ForegroundColor Green
                $Script:ProfileCloneMap[$key] = @{ Id = [int]$newId; Name = $profileName }
                $existingNames[$profileName.ToLower()] = [int]$newId

                ## Verify cloned profile matches source
                $clonedInfo = Invoke-GetAccountProfileInfo -ProfileId ([int]$newId)
                if ($clonedInfo) {
                    ## Compare ProfileData as sorted JSON — ignore PartnerId/Id/Name/Version which legitimately differ
                    $srcJson   = ($profileInfo.ProfileData  | ConvertTo-Json -Depth 100 -Compress) -replace '\s',''
                    $cloneJson = ($clonedInfo.ProfileData   | ConvertTo-Json -Depth 100 -Compress) -replace '\s',''
                    if ($srcJson -eq $cloneJson) {
                        Write-Host "      ✔ ProfileData verified — clone matches source" -ForegroundColor Green
                    } else {
                        Write-Warning "      ✘ ProfileData MISMATCH for '$profileName' (src ID: $srcProfileId → clone ID: $newId)"
                        ## Show top-level ProfileData keys that differ
                        $srcObj   = $profileInfo.ProfileData  | ConvertTo-Json -Depth 100 | ConvertFrom-Json
                        $cloneObj = $clonedInfo.ProfileData   | ConvertTo-Json -Depth 100 | ConvertFrom-Json
                        $srcObj.PSObject.Properties | ForEach-Object {
                            $k = $_.Name
                            $sv = $_.Value | ConvertTo-Json -Depth 5 -Compress
                            $cv = ($cloneObj.$k) | ConvertTo-Json -Depth 5 -Compress
                            if ($sv -ne $cv) {
                                Write-Warning "        Diff key: $k"
                                Write-Warning "          Src   : $sv"
                                Write-Warning "          Clone : $cv"
                            }
                        }
                    }
                } else {
                    Write-Warning "      Could not read back cloned profile ID $newId for verification"
                }
            } else {
                Write-Warning "    Failed to clone '$profileName'"
            }
        }
    }

    ## Clone unique source products (retention policies) to target partners
    $Script:ProductCloneMap = @{}  ## key = "srcPartnerId|srcProductId", value = @{Id=...; Name=...}

    Write-Host "`n$Script:strLineSeparator"
    Write-Host "  Cloning products (retention policies) to target partners..."

    $allFlatRows | Group-Object SourcePartnerId | ForEach-Object {
        $srcPartnerId = [int]$_.Name
        $tgt = $Script:TargetPartnerMap[$srcPartnerId]
        if (-not $tgt) { return }

        Write-Host "  $srcPartnerId → Target: $($tgt.Id) ($($tgt.Name))"
        Invoke-RefreshVisaIfNeeded

        ## Build lookup of names already in target partner (all visible products, including inherited)
        $existingProducts = Invoke-EnumerateProducts -PartnerId $tgt.Id
        $existingProductNames = @{}
        foreach ($ep in $existingProducts) { $existingProductNames[$ep.Name.ToLower()] = [int]$ep.Id }

        ## Unique non-zero ProductIds from devices belonging to this source partner
        $uniqueProductIds = $_.Group |
            Where-Object { $_.ProductId -and $_.ProductId -notin @('','0') } |
            Select-Object -ExpandProperty ProductId -Unique

        Write-Host "    Source product IDs to process: $(if ($uniqueProductIds) { $uniqueProductIds -join ', ' } else { '(none)' })"

        foreach ($srcProductId in $uniqueProductIds) {
            $key = "$srcPartnerId|$srcProductId"

            ## Product ID 1 ("All-In") is a built-in default present in every partner — never clone it
            if ([int]$srcProductId -eq 1) {
                Write-Host "    Skipping product ID 1 (All-In) — built-in default, no clone needed"
                continue
            }

            $productInfo = Invoke-GetProductInfo -ProductId ([int]$srcProductId)
            if (-not $productInfo) { Write-Warning "    Could not retrieve product $srcProductId"; continue }

            $productName = $productInfo.Name

            if ($existingProductNames.ContainsKey($productName.ToLower())) {
                $existingId = $existingProductNames[$productName.ToLower()]
                Write-Host "    '$productName' already exists in target (ID: $existingId) — reusing"
                $Script:ProductCloneMap[$key] = @{ Id = $existingId; Name = $productName }
                continue
            }

            ## Deep-copy then strip read-only and historical properties
            $cloned = $productInfo | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            foreach ($prop in @('Id','CreationDate','ModificationDate','HistoricalLimits','HistoricalArchives')) {
                $cloned.PSObject.Properties.Remove($prop)
            }
            ## Strip HistoryLimit* entries from Features — API forbids them when standard
            ## retention limits (IntraDailyRetention/DailyRetention/etc.) are also present
            if ($cloned.Features) {
                $before = $cloned.Features.Count
                $cloned.Features = @($cloned.Features | Where-Object { $_[0] -notlike 'HistoryLimit*' })
                $stripped = $before - $cloned.Features.Count
                if ($stripped -gt 0) {
                    Write-Host "    (stripped $stripped HistoryLimit* feature(s) from '$productName' before clone)" -ForegroundColor DarkYellow
                }
            }
            $cloned.PartnerId = $tgt.Id

            $newId = Invoke-AddProduct -ProductInfo $cloned
            if ($newId) {
                Write-Host "    '$productName' cloned → new ID: $newId" -ForegroundColor Green
                $Script:ProductCloneMap[$key] = @{ Id = [int]$newId; Name = $productName }
                $existingProductNames[$productName.ToLower()] = [int]$newId

                ## Verify: read back and compare Features
                $clonedInfo = Invoke-GetProductInfo -ProductId ([int]$newId)
                if ($clonedInfo) {
                    $srcFeatures   = ($productInfo.Features | ConvertTo-Json -Depth 5 -Compress)
                    $cloneFeatures = ($clonedInfo.Features  | ConvertTo-Json -Depth 5 -Compress)
                    if ($srcFeatures -eq $cloneFeatures) {
                        Write-Host "      ✔ Features verified — clone matches source" -ForegroundColor Green
                    } else {
                        Write-Warning "      ✘ Features MISMATCH for '$productName' (src: $srcProductId → clone: $newId)"
                        Write-Warning "        Src   : $srcFeatures"
                        Write-Warning "        Clone : $cloneFeatures"
                    }
                } else {
                    Write-Warning "      Could not read back cloned product ID $newId for verification"
                }
            } else {
                Write-Warning "    Failed to clone '$productName'"
            }
        }
    }

    if ($allFlatRows.Count -eq 0) {
        Write-Warning "`n  No devices found across any partner. Exiting."
        Exit 0
    }

    ## Pivot to wide format — one row per device, one output column per distinct custom column
    Write-Host "`n$Script:strLineSeparator"
    Write-Host "  Pivoting: $($allFlatRows.Count) device(s), $($allColNames.Count) distinct column(s)..."

    $wideRows = @(foreach ($row in $allFlatRows) {
        $tgt = $Script:TargetPartnerMap[$row.SourcePartnerId]
        $obj = [ordered]@{
            SourcePartnerId = $row.SourcePartnerId
            SourcePartnerName = $row.PartnerName
            TargetPartnerId   = if ($tgt) { $tgt.Id   } else { '' }
            TargetPartnerName  = if ($tgt) { $tgt.Name } else { '' }
            TargetProfileId    = if ($Script:ProfileCloneMap.ContainsKey("$($row.SourcePartnerId)|$($row.ProfileId)")) { $Script:ProfileCloneMap["$($row.SourcePartnerId)|$($row.ProfileId)"].Id   } else { '' }
            TargetProfileName  = if ($Script:ProfileCloneMap.ContainsKey("$($row.SourcePartnerId)|$($row.ProfileId)")) { $Script:ProfileCloneMap["$($row.SourcePartnerId)|$($row.ProfileId)"].Name } else { '' }
            TargetProductId    = if ($Script:ProductCloneMap.ContainsKey("$($row.SourcePartnerId)|$($row.ProductId)")) { $Script:ProductCloneMap["$($row.SourcePartnerId)|$($row.ProductId)"].Id   } else { '' }
            TargetProductName  = if ($Script:ProductCloneMap.ContainsKey("$($row.SourcePartnerId)|$($row.ProductId)")) { $Script:ProductCloneMap["$($row.SourcePartnerId)|$($row.ProductId)"].Name } else { '' }
            AccountId    = $row.AccountId
            DeviceId     = $row.DeviceId
            DeviceName   = $row.DeviceName
            ComputerName = $row.ComputerName
            DeviceType   = $row.DeviceType
            ProfileId    = $row.ProfileId
            ProfileName  = $row.ProfileName
            ProductId    = $row.ProductId
            ProductName  = $row.ProductName
        }
        foreach ($col in $allColNames) {
            $obj[$col] = if ($row._Values.ContainsKey($col)) { $row._Values[$col] } else { '' }
        }
        [PSCustomObject]$obj
    })

    ## Export to XLSX
    $folder = if ($ExportPath) { $ExportPath } else { $PSScriptRoot }
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $xlsxFile = Join-Path $folder "CustomColumnReport_${Script:Timestamp}.xlsx"

    $wideRows | Export-Excel -Path $xlsxFile `
        -WorksheetName 'CustomColumns' `
        -TableName     'CustomColumnData' `
        -TableStyle    Medium6 `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow

    Write-Host "  Exported to : $xlsxFile"
    Write-Host "  Rows        : $($wideRows.Count)"
    Write-Host "  Columns     : $($allColNames.Count) custom column(s)"
    Write-Host "  Partners    : $($resolvedList.Count) requested, $(($allFlatRows | Select-Object -ExpandProperty PartnerName -Unique).Count) with data"
    Write-Host $Script:strLineSeparator

    if ($Launch) { Start-Process $xlsxFile }

    #region ----- Move Devices (gridview loop) ----
    if ($Move) {
        Write-Host "`n$Script:strLineSeparator"
        Write-Host "  MOVE MODE — select devices from grid to move to target partner"
        Write-Host $Script:strLineSeparator

        ## Load move log to determine which devices have already been moved
        $moveLog = @()
        if (Test-Path $Script:MoveLogFile) {
            $moveLog = @(Import-Csv $Script:MoveLogFile)
        }

        ## Build lookup: DeviceId → last action (MOVE or REVERT)
        Function Get-LastMoveAction ($log, $deviceId) {
            $entries = @($log | Where-Object { $_.DeviceId -eq $deviceId } | Sort-Object Timestamp)
            if ($entries.Count -eq 0) { return '' }
            return $entries[-1].Action
        }

        do {
            ## Rebuild last-action lookup each loop iteration (log may have grown)
            $movedSet = @{}
            foreach ($entry in ($moveLog | Group-Object DeviceId)) {
                $last = ($entry.Group | Sort-Object Timestamp)[-1]
                $movedSet[$entry.Name] = $last.Action
            }

            ## Build gridview rows from $wideRows (already in scope)
            $gridRows = $wideRows | ForEach-Object {
                $devId  = [string]$_.DeviceId
                $status = if ($movedSet.ContainsKey($devId)) { $movedSet[$devId] } else { '' }
                [PSCustomObject]@{
                    Status            = $status
                    DeviceId          = $_.DeviceId
                    DeviceName        = $_.DeviceName
                    ComputerName      = $_.ComputerName
                    DeviceType        = $_.DeviceType
                    SourcePartnerId   = $_.SourcePartnerId
                    SourcePartnerName = $_.SourcePartnerName
                    TargetPartnerId   = $_.TargetPartnerId
                    TargetPartnerName = $_.TargetPartnerName
                    SrcProfileId      = $_.ProfileId
                    SrcProfileName    = $_.ProfileName
                    TgtProfileId      = $_.TargetProfileId
                    TgtProfileName    = $_.TargetProfileName
                    SrcProductId      = $_.ProductId
                    SrcProductName    = $_.ProductName
                    TgtProductId      = $_.TargetProductId
                    TgtProductName    = $_.TargetProductName
                }
            }

            $selected = $gridRows | Out-GridView -Title 'Select devices to MOVE to target partner — close window to exit' -PassThru
            if (-not $selected -or $selected.Count -eq 0) { break }

            Write-Host "`n  Moving $($selected.Count) device(s)..."
            foreach ($row in $selected) {
                Invoke-RefreshVisaIfNeeded
                $devId     = [int]$row.DeviceId
                $tgtPid    = [int]$row.TargetPartnerId
                $tgtProf   = if ($row.TgtProfileId) { [int]$row.TgtProfileId } else { 0 }
                $tgtProd   = if ($row.TgtProductId)  { [int]$row.TgtProductId  } else { 0 }

                if ($tgtPid -eq 0) {
                    Write-Warning "    $($row.DeviceName) — no TargetPartnerId mapped, skipping"
                    continue
                }

                Write-Host "    $($row.DeviceName) ($devId)" -NoNewline
                Write-Host "  $($row.SourcePartnerName) → $($row.TargetPartnerName)" -ForegroundColor Cyan

                $ok = Invoke-ModifyAccount -AccountId $devId -PartnerId $tgtPid -ProfileId $tgtProf -ProductId $tgtProd

                ## Verify: re-enumerate device under target partner using AU filter
                $verifyOk = $false; $verifyNote = ''
                if ($ok) {
                    Start-Sleep -Milliseconds 1500
                    $verifyData = @{
                        jsonrpc='2.0'; id='2'; visa=$Script:visa; method='EnumerateAccountStatistics'
                        params=@{ query=@{ PartnerId=$tgtPid; Columns=@('AU','OI','PD'); Filter="AU == $devId" } }
                    }
                    $verifyResp = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
                        -Body (ConvertTo-Json $verifyData -Depth 5) -Uri $Script:urlJSON `
                        -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
                    $verifyAcct = $verifyResp.result.result | Select-Object -First 1
                    if ($verifyAcct) {
                        $newProfileId  = ($verifyAcct.Settings.OI -join '').Trim()
                        $newProductId  = ($verifyAcct.Settings.PD -join '').Trim()
                        $profMatch     = ($tgtProf -eq 0) -or ($newProfileId -eq [string]$tgtProf)
                        $prodMatch     = ($tgtProd -eq 0) -or ($newProductId -eq [string]$tgtProd)
                        $verifyOk      = $profMatch -and $prodMatch
                        $verifyNote    = if ($verifyOk) { 'Verified OK' } else { "ProfileId=$newProfileId (expected $tgtProf) ProdId=$newProductId (expected $tgtProd)" }
                    } else {
                        $verifyNote = 'Not found under target partner after move'
                    }
                }

                $result = if ($ok) { if ($verifyOk) { 'SUCCESS' } else { 'MOVED-UNVERIFIED' } } else { 'FAILED' }
                $colour = if ($result -eq 'SUCCESS') { 'Green' } elseif ($result -eq 'FAILED') { 'Red' } else { 'Yellow' }
                Write-Host "      → $result  $verifyNote" -ForegroundColor $colour

                ## Append to move log
                $logRow = [PSCustomObject]@{
                    Timestamp        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    Action           = 'MOVE'
                    DeviceId         = $devId
                    DeviceName       = $row.DeviceName
                    SrcPartnerId     = $row.SourcePartnerId
                    SrcPartnerName   = $row.SourcePartnerName
                    SrcProfileId     = $row.SrcProfileId
                    SrcProfileName   = $row.SrcProfileName
                    SrcProductId     = $row.SrcProductId
                    SrcProductName   = $row.SrcProductName
                    TgtPartnerId     = $row.TargetPartnerId
                    TgtPartnerName   = $row.TargetPartnerName
                    TgtProfileId     = $row.TgtProfileId
                    TgtProfileName   = $row.TgtProfileName
                    TgtProductId     = $row.TgtProductId
                    TgtProductName   = $row.TgtProductName
                    Result           = $result
                    Notes            = $verifyNote
                }
                $logRow | Export-Csv -Path $Script:MoveLogFile -Append -NoTypeInformation -Encoding UTF8
                $moveLog += $logRow
            }
            Write-Host "  Batch complete. Close gridview to exit, or re-select to move more."
        } while ($true)

        Write-Host "`n  Move session complete. Log: $Script:MoveLogFile"
    }
    #endregion ----- Move Devices ----

    #region ----- Revert Devices (gridview) ----
    if ($Revert) {
        Write-Host "`n$Script:strLineSeparator"
        Write-Host "  REVERT MODE — select previously moved devices to send back to source"
        Write-Host $Script:strLineSeparator

        if (-not (Test-Path $Script:MoveLogFile)) {
            Write-Warning "  No move log found at $Script:MoveLogFile — nothing to revert"
        } else {
            $moveLog = @(Import-Csv $Script:MoveLogFile)

            ## For each DeviceId find last action; keep only those whose last action is MOVE
            $revertCandidates = $moveLog | Group-Object DeviceId | ForEach-Object {
                $last = ($_.Group | Sort-Object Timestamp)[-1]
                if ($last.Action -eq 'MOVE') { $last }
            }

            if (-not $revertCandidates -or @($revertCandidates).Count -eq 0) {
                Write-Host "  No devices eligible for revert (all already reverted or log empty)"
            } else {
                do {
                    $selected = @($revertCandidates) | Out-GridView `
                        -Title 'Select devices to REVERT to source partner — close window to exit' -PassThru
                    if (-not $selected -or $selected.Count -eq 0) { break }

                    Write-Host "`n  Reverting $($selected.Count) device(s)..."
                    foreach ($row in $selected) {
                        Invoke-RefreshVisaIfNeeded
                        $devId    = [int]$row.DeviceId
                        $srcPid   = [int]$row.SrcPartnerId
                        $srcProf  = if ($row.SrcProfileId) { [int]$row.SrcProfileId } else { 0 }
                        $srcProd  = if ($row.SrcProductId)  { [int]$row.SrcProductId  } else { 0 }

                        Write-Host "    $($row.DeviceName) ($devId)" -NoNewline
                        Write-Host "  $($row.TgtPartnerName) → $($row.SrcPartnerName)" -ForegroundColor Cyan

                        $ok = Invoke-ModifyAccount -AccountId $devId -PartnerId $srcPid -ProfileId $srcProf -ProductId $srcProd

                        ## Verify under source partner using AU filter
                        $verifyOk = $false; $verifyNote = ''
                        if ($ok) {
                            Start-Sleep -Milliseconds 1500
                            $verifyData = @{
                                jsonrpc='2.0'; id='2'; visa=$Script:visa; method='EnumerateAccountStatistics'
                                params=@{ query=@{ PartnerId=$srcPid; Columns=@('AU','OI','PD'); Filter="AU == $devId" } }
                            }
                            $verifyResp = Invoke-WebRequest -Method POST -ContentType 'application/json; charset=utf-8' `
                                -Body (ConvertTo-Json $verifyData -Depth 5) -Uri $Script:urlJSON `
                                -SessionVariable Script:websession -UseBasicParsing | ConvertFrom-Json
                            $verifyAcct = $verifyResp.result.result | Select-Object -First 1
                            if ($verifyAcct) {
                                $newProfileId = ($verifyAcct.Settings.OI -join '').Trim()
                                $newProductId = ($verifyAcct.Settings.PD -join '').Trim()
                                $profMatch    = ($srcProf -eq 0) -or ($newProfileId -eq [string]$srcProf)
                                $prodMatch    = ($srcProd -eq 0) -or ($newProductId -eq [string]$srcProd)
                                $verifyOk     = $profMatch -and $prodMatch
                                $verifyNote   = if ($verifyOk) { 'Verified OK' } else { "ProfileId=$newProfileId (expected $srcProf) ProdId=$newProductId (expected $srcProd)" }
                            } else {
                                $verifyNote = 'Not found under source partner after revert'
                            }
                        }

                        $result = if ($ok) { if ($verifyOk) { 'SUCCESS' } else { 'REVERTED-UNVERIFIED' } } else { 'FAILED' }
                        $colour = if ($result -eq 'SUCCESS') { 'Green' } elseif ($result -eq 'FAILED') { 'Red' } else { 'Yellow' }
                        Write-Host "      → $result  $verifyNote" -ForegroundColor $colour

                        ## Append revert row to log
                        $logRow = [PSCustomObject]@{
                            Timestamp        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                            Action           = 'REVERT'
                            DeviceId         = $devId
                            DeviceName       = $row.DeviceName
                            SrcPartnerId     = $row.SrcPartnerId
                            SrcPartnerName   = $row.SrcPartnerName
                            SrcProfileId     = $row.SrcProfileId
                            SrcProfileName   = $row.SrcProfileName
                            SrcProductId     = $row.SrcProductId
                            SrcProductName   = $row.SrcProductName
                            TgtPartnerId     = $row.TgtPartnerId
                            TgtPartnerName   = $row.TgtPartnerName
                            TgtProfileId     = $row.TgtProfileId
                            TgtProfileName   = $row.TgtProfileName
                            TgtProductId     = $row.TgtProductId
                            TgtProductName   = $row.TgtProductName
                            Result           = $result
                            Notes            = $verifyNote
                        }
                        $logRow | Export-Csv -Path $Script:MoveLogFile -Append -NoTypeInformation -Encoding UTF8
                        ## Remove from revert candidates so gridview reflects updated state
                        $revertCandidates = @($revertCandidates | Where-Object { $_.DeviceId -ne [string]$devId })
                    }
                    Write-Host "  Batch complete. Close gridview to exit, or re-select to revert more."
                } while ($revertCandidates.Count -gt 0)
            }
        }
        Write-Host "`n  Revert session complete. Log: $Script:MoveLogFile"
    }
    #endregion ----- Revert Devices ----

#endregion ----- Main Execution ----
