<# ----- About: ----
    # Bulk Set SW Backup GUI Password 
    # Revision v15 - 2026-09-29 - Added -DeviceName (target device(s) by name, bypassing partner
    #   selection entirely via RecurseSubPartners), double-entry masked password confirmation,
    #   optional -GUIPassword (SecureString or plain string), -DebugCDP (masked, decoupled from
    #   PowerShell's native -Debug which also leaks raw HTTP bodies/headers via Invoke-RestMethod)
    # Revision v14 - 2026-09-29 - Modernized credential storage (XML/SecureString) and replaced
    #   deprecated GetPartnerInfo(name=) with GetPartnerInfoById/GetPartnerTree, ported from the
    #   confirmed-working auth/partner-selection pattern in Get-CombinedDeviceProfileReport.v05.ps1
    #   (Invoke-CoveAPI helper, retry/backoff Login, auto-refreshing visa, Get-CovePartnerInfo)
    # Revision v13 - 2021-07-07
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
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
    # For use with the Standalone edition of N-able Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select devices
    # Prompt/ Set/ Wipe GUI password via Remote commands
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -DeviceName parameter to target device(s) by exact Device Name match, bypassing partner selection entirely
    # Use the -AllDevices switch parameter to skip GUI device selection
    #
    # Use the -SetGUIPassword (default) parameter to be prompted to enter a Secure GUI Password to be applied
    # Use the -RestoreOnly parameter with the -SetGUIPassword parameter to allow restores when GUI password is set
    # Use the -WipeGUIPassword parameter to clear the GUI password from selected devices
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script

    # https://developer.n-able.com/n-able-cove/docs/getting-started
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/remote-commands.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="SetGUIPW")]
    Param (
        [Parameter(ParameterSetName="WipeGUIPW",Mandatory=$False)] [Switch]$WipeGUIPassword,  ## Clear GUI Password   
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] [switch]$SetGUIPassword,    ## Specify GUI Password to set
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] $GUIPassword,               ## Optionally supply the GUI Password to set - accepts a SecureString or a plain string (skips the interactive masked/confirm prompt)
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] [Switch]$RestoreOnly,       ## Allow Restore Only GUI Access
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] 
            [Parameter(ParameterSetName="WipeGUIPW",Mandatory=$False)][Switch]$AllPartners,   ## Skip partner selection
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] 
            [Parameter(ParameterSetName="WipeGUIPW",Mandatory=$False)] [string]$DeviceName,   ## Target device(s) by exact Device Name (AN) match - bypasses partner selection entirely, searching the whole authenticated partner tree
        [Parameter(ParameterSetName="SetGUIPW",Mandatory=$False)] 
            [Parameter(ParameterSetName="WipeGUIPW",Mandatory=$False)] [Switch]$AllDevices,   ## Skip device selection             
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials,                              ## Remove Stored API Credentials at start of script
        [Parameter(Mandatory=$False)] [switch]$DebugCDP,                                      ## Show masked per-device remote command output (NOTE: do not use PowerShell's own -Debug/-Verbose instead - Invoke-RestMethod's own tracing under those dumps the RAW, UNMASKED request body/headers including the password and Bearer visa token)
        [Parameter(Mandatory=$False)] [string]$APICredentialFile = "C:\ProgramData\MXB\$($env:computername)_$($env:username)_API_Credentials.Secure.xml"  ## Backup.Management API credential XML file (DPAPI-encrypted Export-Clixml)
        
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

    $ConsoleTitle = "Bulk Set GUI Password"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-Host "`n  $ConsoleTitle`n" -ForegroundColor Cyan
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Host "  Script Parameter Syntax:`n`n  $Syntax" -ForegroundColor DarkGray
    Write-Host "  Current Parameters:" -ForegroundColor DarkGray
    Write-Host "  -AllPartners     = $AllPartners" -ForegroundColor DarkGray
    Write-Host "  -DeviceName      = $(if ($DeviceName) { $DeviceName } else { '<not supplied>' })" -ForegroundColor DarkGray
    Write-Host "  -AllDevices      = $AllDevices" -ForegroundColor DarkGray
    Write-Host "  -SetGUIPassword  = $SetGUIPassword" -ForegroundColor DarkGray
    Write-Host "  -RestoreOnly     = $RestoreOnly" -ForegroundColor DarkGray
    Write-Host "  -WipeGUIPassword = $WipeGUIPassword" -ForegroundColor DarkGray
    Write-Host "  -GUIPassword     = $(if ($GUIPassword) { '<supplied>' } else { '<not supplied - will prompt>' })" -ForegroundColor DarkGray
    Write-Host "  -DebugCDP        = $DebugCDP" -ForegroundColor DarkGray
    Write-Host "  Action           = $($PSCmdlet.ParameterSetName)  (default is SetGUIPW - use -WipeGUIPassword to clear instead)" -ForegroundColor Yellow
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::MaxServicePointIdleTime = 5000000
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $urlJSON = 'https://api.backup.management/jsonapi'

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication (ported from Get-CombinedDeviceProfileReport.v05.ps1's confirmed-working pattern) ----
    Function ConvertFrom-SecureString2 ($SecureString) {
        $ptr  = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
        try { [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
    }  ## Decrypt a SecureString to plaintext only at the point of use

    Function Get-APICredentials {

        $Script:APIcredfile = $APICredentialFile
        $Script:APIcredpath = Split-Path -Path $Script:APIcredfile

        if ($ClearCredentials -and (Test-Path $Script:APIcredfile)) {
            Remove-Item -Path $Script:APIcredfile -Force
            Write-Host "  Backup API Credential File Cleared" -ForegroundColor Yellow
        }

        if (Test-Path $Script:APIcredfile) {
            Write-Host "  Loading credentials from $Script:APIcredfile" -ForegroundColor Cyan
            $stored = Import-Clixml -Path $Script:APIcredfile
            $Script:cred0 = $stored.PartnerName
            $Script:cred1 = $stored.Username
            $Script:cred2 = ConvertFrom-SecureString2 ($stored.Password | ConvertTo-SecureString)

            Write-Host "  Partner  : $Script:cred0" -ForegroundColor DarkCyan
            Write-Host "  API User : $Script:cred1" -ForegroundColor DarkCyan
        }else{
            Write-Host "  No credential file found. Creating..." -ForegroundColor Yellow
            if (-not (Test-Path $Script:APIcredpath)) { New-Item -ItemType Directory -Path $Script:APIcredpath -Force | Out-Null }

            Write-Host "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'" -ForegroundColor DarkGray
            DO{ $LoginPartnerName = Read-Host "  Enter Login Partner Name" }
            WHILE ($LoginPartnerName.length -eq 0)

            $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
            if (-not $BackupCred) { throw "No credentials provided." }

            [PSCustomObject]@{
                PartnerName = $LoginPartnerName
                Username    = $BackupCred.UserName
                Password    = ($BackupCred.Password | ConvertFrom-SecureString)   ## DPAPI encrypted
            } | Export-Clixml -Path $Script:APIcredfile -Force

            $Script:cred0 = $LoginPartnerName
            $Script:cred1 = $BackupCred.UserName
            $Script:cred2 = ConvertFrom-SecureString2 $BackupCred.Password
        }

    }  ## Get (or create) DPAPI-encrypted XML credentials

    Function Test-CoveTransientError ($ErrorRecord) {

        $statusCode = 0
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
            $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
        }
        if ($statusCode -eq 408 -or $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) { return $true }

        $message = $ErrorRecord.Exception.ToString()
        return $message -match '(?i)(timed out|timeout|connection.*(failed|closed|reset|refused)|SSL connection|did not properly respond|temporarily unavailable)'

    }  ## Identify network/API errors worth retrying vs. hard failures

    Function Invoke-WithRetry ($Action, $Label, [int]$MaxRetries = 4) {
        $lastError = $null
        for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
            try {
                return @{ Response = (& $Action); Error = $null }
            }catch{
                $lastError = $_
                $canRetry = $attempt -lt $MaxRetries -and (Test-CoveTransientError $_)
                if (-not $canRetry) { break }
                $delayMs = ([Math]::Min(30000, [int](2000 * [Math]::Pow(2, $attempt)))) + (Get-Random -Minimum 0 -Maximum 1000)
                Write-Host "  [RETRY] $Label failed; retrying in $([Math]::Round($delayMs / 1000, 1)) seconds ($($attempt + 1)/$MaxRetries)" -ForegroundColor Yellow
                Start-Sleep -Milliseconds $delayMs
            }
        }
        return @{ Response = $null; Error = $lastError }
    }  ## Shared exponential-backoff retry wrapper - used by both Login and Invoke-CoveAPI

    Function Send-APICredentialsCookie {

        Get-APICredentials  ## Read (or create) XML Credential File before Authentication

        $body = (ConvertTo-Json @{ jsonrpc='2.0'; id='2'; method='Login'; params=@{ username=$Script:cred1; password=$Script:cred2 } })

        $result = Invoke-WithRetry -Label 'Login' -Action { Invoke-RestMethod -Uri $urlJSON -Method POST -ContentType 'application/json' -Body $body -TimeoutSec 120 -ErrorAction Stop }
        $resp = $result.Response
        $Script:cred2 = $null   ## plaintext password no longer needed once the Login attempt(s) are done

        if ($result.Error -and -not $resp) {
            Write-Host "  Login request failed after 4 retries: $($result.Error.Exception.Message)" -ForegroundColor Red
            Write-Host "  This appears to be a network/API transport failure, not necessarily invalid credentials." -ForegroundColor Yellow
            throw $result.Error
        }

        if ($resp.visa) {
            $Script:visa = $resp.visa
            $Script:AuthPartnerId = [int]$resp.result.result.PartnerId
            Write-Host "  Authenticated as: $Script:cred1" -ForegroundColor Green
            Write-Host "  Authenticated PartnerId (default search scope): $Script:AuthPartnerId" -ForegroundColor DarkCyan
        }else{
            Write-Host "  Authentication Failed: Please confirm your Backup.Management Credentials" -ForegroundColor Red
            Write-Host "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account" -ForegroundColor Yellow
            Remove-Item -Path $Script:APIcredfile -Force -ErrorAction SilentlyContinue  ## Stored creds are stale/invalid - re-prompt
            Send-APICredentialsCookie
        }

    }  ## Use Backup.Management credentials to Authenticate (with transient-error retry/backoff)

    Function Get-VisaTime {
        if ($Script:visa) {
            $visaTime = Convert-UnixTimeToDateTime ([int]$Script:visa.Split('-')[3])
            if ($visaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)) {
                Write-Host "  API visa is older than 10 minutes. Refreshing authentication..." -ForegroundColor Yellow
                Send-APICredentialsCookie
            }
        }
    }  ## Refresh the API visa automatically once it's more than 10 minutes old

#endregion ----- Authentication ----

#region ----- Data Conversion ----
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
}  ## Convert epoch time to date time 

#endregion ----- Data Conversion ----

#region ----- Secure Password Entry ----
    Function Read-MaskedPassword ($Prompt) {
        Write-Host -NoNewline "$Prompt`: "
        $secure = [System.Security.SecureString]::new()
        while ($true) {
            $key = [System.Console]::ReadKey($true)
            if ($key.Key -eq 'Enter') { break }
            elseif ($key.Key -eq 'Backspace') {
                if ($secure.Length -gt 0) { $secure.RemoveAt($secure.Length - 1); Write-Host -NoNewline "`b `b" }
            }else{
                $secure.AppendChar($key.KeyChar)
                Write-Host -NoNewline "*"
            }
        }
        Write-Host ""
        $secure.MakeReadOnly()
        return $secure
    }  ## Read a password one keystroke at a time, echoing '*' instead of the real characters

    Function Test-SecureStringsMatch ($First, $Second) {
        $bstrA = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($First)
        $bstrB = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Second)
        try { return ([Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrA)) -ceq ([Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrB)) }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrA)
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrB)
        }
    }  ## Case-sensitive SecureString comparison, decrypted only for the instant of the compare

    Function Read-ConfirmedPassword ($Prompt) {
        DO {
            $first  = Read-MaskedPassword $Prompt
            $second = Read-MaskedPassword "  Re-enter Password to Confirm"
            $isMatch = Test-SecureStringsMatch $first $second
            if (-not $isMatch) { Write-Host "  Passwords do not match - starting over." -ForegroundColor Yellow }
        } WHILE (-not $isMatch)
        return $first
    }  ## Double-entry confirmation - restarts at entry 1 on a mismatch
#endregion ----- Secure Password Entry ----

#region ----- Backup.Management JSON Calls ----

    Function Invoke-CoveAPI ($Method, $Params, [int]$MaxRetries = 4) {

        Get-VisaTime
        $payload = (ConvertTo-Json @{ jsonrpc='2.0'; id='1'; method=$Method; visa=$Script:visa; params=$Params } -Depth 10)

        $result = Invoke-WithRetry -Label $Method -MaxRetries $MaxRetries -Action { Invoke-RestMethod -Uri $urlJSON -Method POST -ContentType 'application/json' -Body $payload -TimeoutSec 120 -ErrorAction Stop }
        if ($result.Response) { return $result.Response }
        return @{ error = @{ message = $result.Error.Exception.Message }; result = $null }

    }  ## Generic Cove jsonapi POST helper - auto-refreshes the visa and retries transient network errors (via Invoke-WithRetry)

    Function Get-CovePartnerInfoById ($PartnerId) {

        $resp = Invoke-CoveAPI 'GetPartnerInfoById' @{ partnerId = [int]$PartnerId }
        if ($resp.error) { throw "GetPartnerInfoById($PartnerId) failed: $($resp.error.message)" }

        $info = $resp.result
        if ($info -and $info.PSObject.Properties['result']) { $info = $info.result }
        if (-not $info) { throw "GetPartnerInfoById($PartnerId) returned no partner details." }

        [PSCustomObject]@{
            Uid   = [string]$info.Uid
            Id    = [int]$info.Id
            Name  = [string]$info.Name
            Level = [string]$info.Level
        }

    } ## GetPartnerInfoById API Call (replaces deprecated GetPartnerInfo(name=) self-lookup)

    Function Get-CovePartnerInfo ($RequestedPartnerName, [int]$SearchPartnerId = $Script:AuthPartnerId) {
        ## Combined replacement for deprecated GetPartnerInfo(name=...) - CASE-INSENSITIVE SUBSTRING
        ## match. Non-restricted logins with no override name use their own partner directly (no
        ## search needed); Root/Sub-root/Distributor logins (or an explicit name) search via
        ## GetPartnerTree + partnerFilter.NamePattern, disambiguating 2+ matches via Out-GridView.

        $restrictedLevels = @('Root','Sub-root','Distributor')

        if (-not $RequestedPartnerName) {
            $self = Get-CovePartnerInfoById $SearchPartnerId
            if ($self.Level -notin $restrictedLevels) { return $self }

            DO{ $RequestedPartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'" }
            WHILE ($RequestedPartnerName.length -eq 0)
        }

        DO{
            Write-Host "  Searching partner tree for '$RequestedPartnerName' (scope $SearchPartnerId; API timeout 120 seconds)..." -ForegroundColor Cyan
            $treeResp = Invoke-CoveAPI 'GetPartnerTree' @{
                partnerId     = $SearchPartnerId
                fields        = @(0,1,3,4)
                childrenLimit = 5000
                partnerFilter = @{ NamePattern = [string]$RequestedPartnerName }
            }

            if ($treeResp.error) {
                Write-Host "  GetPartnerTree failed: $($treeResp.error.message)" -ForegroundColor Yellow
                $partnerMatches = @()
            }else{
                $treeRoot = $treeResp.result
                if ($treeRoot -and $treeRoot.PSObject.Properties['result']) { $treeRoot = $treeRoot.result }

                $partnerMatches = [System.Collections.Generic.List[object]]::new()
                $partnerTreeInfoById = @{}   ## every node seen, not just matches - needed to walk ancestors for the Path breadcrumb
                Function Add-PartnerTreeMatch ($Node, $Needle) {
                    if ($Node.Info -and $Node.Info.Id) { $partnerTreeInfoById[[int]$Node.Info.Id] = $Node.Info }
                    if ($Node.Info -and $Node.Info.Name -and ($Node.Info.Name -like "*$Needle*") -and ($Node.Info.Level -notin @('Root','Sub-root','Distributor'))) {
                        $partnerMatches.Add([PSCustomObject]@{ Id = "$($Node.Info.Id)"; Name = $Node.Info.Name; Level = $Node.Info.Level; ParentId = "$($Node.Info.ParentId)"; Path = '' }) | Out-Null
                    }
                    foreach ($Child in $Node.Children) { Add-PartnerTreeMatch $Child $Needle }
                } ## Recursively flatten the GetPartnerTree Info/Children nodes into a match list
                if ($treeRoot) { Add-PartnerTreeMatch $treeRoot $RequestedPartnerName }
                $partnerMatches = @($partnerMatches)

                ## Build the "Grandparent > Parent > Name" breadcrumb for each match (matches
                ## Get-CombinedDeviceProfileReport.v05.ps1's gridview style) - walk ParentId up to the root.
                foreach ($partnerMatch in $partnerMatches) {
                    $pathParts = [System.Collections.Generic.List[string]]::new()
                    $pathNode = $partnerTreeInfoById[[int]$partnerMatch.Id]
                    $pathGuard = [System.Collections.Generic.HashSet[int]]::new()
                    while ($pathNode -and $pathGuard.Add([int]$pathNode.Id)) {
                        $pathParts.Insert(0, [string]$pathNode.Name)
                        $pathNode = $partnerTreeInfoById[[int]$pathNode.ParentId]
                    }
                    $partnerMatch.Path = $pathParts -join ' > '
                }
            }

            if ($partnerMatches.Count -eq 0) {
                Write-Host "  No partner found matching '$RequestedPartnerName'." -ForegroundColor Yellow
                DO{ $RequestedPartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup" }
                WHILE ($RequestedPartnerName.length -eq 0)
                continue
            }

            if ($partnerMatches.Count -gt 1) {
                Write-Host "  At least $($partnerMatches.Count) partners match '$RequestedPartnerName'. Select one:" -ForegroundColor Cyan
                $Selected = $partnerMatches | Select-Object Id,Name,Level,ParentId,Path | Out-GridView -Title "Select Partner matching '$RequestedPartnerName'" -OutputMode Single
                if (-not $Selected) {
                    DO{ $RequestedPartnerName = Read-Host "  No partner selected. Enter Customer/ Partner display name (or partial name) to lookup" }
                    WHILE ($RequestedPartnerName.length -eq 0)
                    continue
                }
            }else{
                $Selected = $partnerMatches[0]
            }

            return Get-CovePartnerInfoById ([int]$Selected.Id)

        } WHILE ($true)

    } ## GetPartnerTree partnerFilter.NamePattern search, disambiguate via GridView if needed

    Function Send-GetDevices ($PartnerId, [switch]$RecurseSubPartners) {

        Write-Host "  Enumerating devices for partner $PartnerId$(if($RecurseSubPartners){' (incl. sub-partners)'})..." -ForegroundColor Cyan
        Write-Progress -Id 1 -Activity "Bulk Set GUI Password" -Status "Enumerating devices..." -PercentComplete 10

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        if ($RecurseSubPartners) { $data.params.query.RecurseSubPartners = $true }
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","MN","AL","LN","OP","OI","OS","PD","AP","PF","PN","CD","TS","TL","T3","US","AA843","AA77","AA2048")
        $data.params.query.OrderBy = "CD DESC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = 5000
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            ContentType = 'application/json; charset=utf-8'
        }  

        $Script:DeviceResponse = Invoke-RestMethod @params 
      
        $Script:DeviceDetail = @()

        ForEach ( $DeviceResult in $DeviceResponse.result.result ) {

        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [Int]$DeviceResult.AccountId;
                                                                    PartnerID      = [string]$DeviceResult.PartnerId;
                                                                    DeviceName     = $DeviceResult.Settings.AN -join '' ;
                                                                    ComputerName   = $DeviceResult.Settings.MN -join '' ;
                                                                    DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                                                                    PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                    Reference      = $DeviceResult.Settings.PF -join '' ;
                                                                    Creation       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                                                                    TimeStamp      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;  
                                                                    LastSuccess    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '') ;                                                                                                                                                                                                               
                                                                    SelectedGB     = (($DeviceResult.Settings.T3 -join '') /1GB) ;  
                                                                    UsedGB         = (($DeviceResult.Settings.US -join '') /1GB) ;  
                                                                    DataSources    = $DeviceResult.Settings.AP -join '' ;                                                                
                                                                    Account        = $DeviceResult.Settings.AU -join '' ;
                                                                    Location       = $DeviceResult.Settings.LN -join '' ;
                                                                    Notes          = $DeviceResult.Settings.AA843 -join '' ;
                                                                    GUIPassword    = $DeviceResult.Settings.AA2048 -join '' ;                                                                    
                                                                    TempInfo       = $DeviceResult.Settings.AA77 -join '' ;
                                                                    Product        = $DeviceResult.Settings.PN -join '' ;
                                                                    ProductID      = $DeviceResult.Settings.PD -join '' ;
                                                                    Profile        = $DeviceResult.Settings.OP -join '' ;
                                                                    OS             = $DeviceResult.Settings.OS -join '' ;                                                                
                                                                    ProfileID      = $DeviceResult.Settings.OI -join '' }
        }

        Write-Progress -Id 1 -Activity "Bulk Set GUI Password" -Completed

    } ## EnumerateAccountStatistics API Call

    Function Get-AccountHomeNode ($AccountId) {

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = 'jsonrpc'
        $data.visa = $Script:visa
        $data.method = 'GetAccountInfoById'
        $data.params = @{}
        $data.params.accountId = [int]$AccountId

        $jsondata = (ConvertTo-Json $data -depth 6)

        try {
            $webrequest = Invoke-RestMethod -Method POST `
                -ContentType 'application/json; charset=utf-8' `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($jsondata)) `
                -Uri $url `
                -TimeoutSec 60 `
                -UseBasicParsing `
                -ErrorAction Stop
        }catch{
            throw "GetAccountInfoById failed for AccountId $AccountId (endpoint: $url): $($_.Exception.Message)"
        }

        if ($webrequest.error) { throw "GetAccountInfoById failed for AccountId $AccountId : $($webrequest.error.message)" }

        [PSCustomObject]@{
            Host  = ($webrequest.result.homeNodeInfo.CommonInfo.Host -split ':')[0]   ## strip the :443 port suffix
            Token = $webrequest.result.result.Token
            Name  = $webrequest.result.result.Name
        }

    } ## Resolve the device's home-node host + per-device token needed for SendAccountCommand

    Function Send-RemoteCommand ($Device) { 

        if ($SecurePassword.length -ge 1) {
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
            try { $UnsecureGUIPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }   ## zero+free the native buffer - previously leaked/never wiped
        }
        if ($WipeGUIPassword) {
            $UnsecureGUIPassword = ""
        }
        ## -WipeGUIPassword must clear $UnsecureGUIPassword BEFORE $PasswordParam is built, not after -
        ## building it first then clearing was a no-op left over from v13 (the empty-password wipe
        ## behavior only ever "worked" because $SecurePassword is never populated in the WipeGUIPW
        ## parameter set to begin with).
        if ($RestoreOnly) {
            $PasswordParam = "password $($UnsecureGUIPassword)`nrestore_only allow"
        }else{
            $PasswordParam = "password $($UnsecureGUIPassword)`nrestore_only disallow"
        }
            
        ## SendRemoteCommands / backup.management/jsonrpcv1 is deprecated with no jsonapi equivalent -
        ## the console instead resolves the device's home node via GetAccountInfoById, then posts a
        ## PER-DEVICE SendAccountCommand to that node's own repserv_json endpoint (confirmed via live
        ## capture, 2026-09-29, including the "set gui password" command + parameters shape below -
        ## see /memories/cove-api-endpoints.md).
        $HomeNode = Get-AccountHomeNode -AccountId $Device.accountid

        $url = "https://$($HomeNode.Host)/repserv_json"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = 1
        $data.visa = $Script:visa
        $data.method = 'SendAccountCommand'
        $data.params = @{}
        $data.params.account = $HomeNode.Name
        $data.params.token = $HomeNode.Token
        $data.params.command = "set gui password"
        $data.params.parameters = "$PasswordParam"
    
        $jsondata = (ConvertTo-Json $data -depth 6)
        #$jsondata  ## Debug

        ## Gated behind -DebugCDP (a dedicated switch, NOT PowerShell's native -Debug/-Verbose -
        ## $DebugPreference/$VerbosePreference also enable Invoke-RestMethod's OWN internal tracing,
        ## which dumps the raw, UNMASKED request body/headers including the password and Bearer visa
        ## token - see /memories/powershell-gotchas.md). The real password never appears even here,
        ## only its masked length + which device it's being sent to.
        if ($DebugCDP) {
            $MaskedPassword = if ($UnsecureGUIPassword) { '*' * $UnsecureGUIPassword.Length } else { '<empty>' }
            Write-Host "[DebugCDP] $($Device.DeviceName) ($($Device.AccountID)): $($data.params.command) | password $MaskedPassword | restore_only $(if ($RestoreOnly) { 'allow' } else { 'disallow' })" -ForegroundColor DarkGray
        }

        $params = @{
            Uri         = $url
            Method      = $method
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            ContentType = 'application/json; charset=utf-8'
            TimeoutSec  = 60
            ErrorAction = 'Stop'
        }  

            try {
                $Script:Result = Invoke-RestMethod @params
            }catch{
                throw "SendAccountCommand failed for $($Device.DeviceName) ($($Device.AccountID)) - home node $($HomeNode.Host): $($_.Exception.Message)"
            }

    } ## SendAccountCommand API Call (replaces deprecated SendRemoteCommands)

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    Send-APICredentialsCookie

    $Script:PartnerInfo = if ($AllPartners -or $DeviceName) {
        Get-CovePartnerInfoById $Script:AuthPartnerId   ## -AllPartners/-DeviceName: skip the interactive partner picker
    }else{
        Get-CovePartnerInfo
    }
    $Script:PartnerId   = [int]$Script:PartnerInfo.Id
    $Script:PartnerName = $Script:PartnerInfo.Name
    $Script:Level       = $Script:PartnerInfo.Level
    $Script:Uid         = $Script:PartnerInfo.Uid

    Write-Host "  $Script:PartnerName - $Script:PartnerId - $Script:Uid" -ForegroundColor Green

    if ($DeviceName) {
        $EscapedDeviceName = $DeviceName -replace "'", "''"   ## escape literal quotes in the filter expression, same convention as SQL
        $filter1 = "AT == 1 AND PN != 'Documents' AND AN == '$EscapedDeviceName'"   ## exact match on Device Name only (no MN/ComputerName check)
        Write-Host "  -DeviceName supplied: searching your entire authenticated scope (rooted at $Script:PartnerName) for '$DeviceName' (bypasses partner selection)" -ForegroundColor Cyan
        Send-GetDevices $Script:PartnerId -RecurseSubPartners
    }else{
        $filter1 = "AT == 1 AND PN != 'Documents'"   ### Excludes M365 and Documents devices from lookup.
        Send-GetDevices $partnerId
    }

    $GridAction = if ($PSCmdlet.ParameterSetName -eq 'WipeGUIPW') { 'Wipe GUI Password' } else { 'Set GUI Password' }

    if ($DeviceName) {
        $Script:SelectedDevices = @($DeviceDetail | Select-Object PartnerId,PartnerName,Reference,@{Name="AccountID"; Expression={"$($_.AccountId)"}},DeviceName,ComputerName,DeviceAlias,GUIPassword,Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,SelectedGB,UsedGB,Location,OS,Notes,TempInfo)
        if ($Script:SelectedDevices.Count -eq 0) {
            Write-Host "  No devices matched '$DeviceName'" -ForegroundColor Yellow
            $Script:SelectedDevices = $null
        }elseif ($Script:SelectedDevices.Count -gt 1) {
            Write-Host "  $($Script:SelectedDevices.Count) devices match '$DeviceName' - select one or more:" -ForegroundColor Cyan
            $Script:SelectedDevices = $Script:SelectedDevices | Out-GridView -Title "$GridAction | matches for '$DeviceName' - select one or more devices" -OutputMode Multiple
        }else{
            Write-Host "  1 device matched '$DeviceName': $($Script:SelectedDevices.DeviceName)" -ForegroundColor Green
        }
    }elseif ($AllDevices) {
        $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,@{Name="AccountID"; Expression={"$($_.AccountId)"}},DeviceName,ComputerName,DeviceAlias,GUIPassword,Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,SelectedGB,UsedGB,Location,OS,Notes,TempInfo
        Write-Host "  $($SelectedDevices.AccountId.count) Devices Selected" -ForegroundColor Green
    }else{

        $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,@{Name="AccountID"; Expression={"$($_.AccountId)"}},DeviceName,ComputerName,DeviceAlias,GUIPassword,Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,SelectedGB,UsedGB,Location,OS,Notes,TempInfo | Out-GridView -Title "$GridAction | $Script:PartnerName - select one or more devices" -OutputMode Multiple
    }    

    if($null -eq $SelectedDevices) {
        # Cancel was pressed
        # Run cancel script
        Write-Host "  No Devices Selected" -ForegroundColor Yellow
        Break
    }
    else {
        # OK was pressed, $Selection contains what was chosen
        # Run OK script
        $script:SelectedDevices |  Select-Object PartnerId,PartnerName,Reference,@{Name="AccountID"; Expression={[int]$_.AccountId}},DeviceName,ComputerName,DeviceAlias,GUIPassword,Creation,TimeStamp | Sort-object AccountId | Format-Table

        if ($PSCmdlet.ParameterSetName -eq "SetGUIPW") {
            if ($GUIPassword) {
                $SecurePassword = if ($GUIPassword -is [SecureString]) { $GUIPassword } else { ConvertTo-SecureString -String ([string]$GUIPassword) -AsPlainText -Force }
                Write-Host "  Using GUI Password supplied via -GUIPassword parameter" -ForegroundColor DarkGray
            }else{
                $SecurePassword = Read-ConfirmedPassword "  Enter Backup Manager GUI Password to be applied to $($SelectedDevices.AccountId.count) Devices"
            }
            Write-Host "  Applying GUI Password to $($SelectedDevices.AccountId.count) Devices, please be patient." -ForegroundColor Cyan
            }else{
                Write-Host "  Wiping GUI Password from $($SelectedDevices.AccountId.count) Devices, please be patient." -ForegroundColor Cyan
            }

        $deviceCounter = 0
        foreach ($selecteddevice in $SelectedDevices) {

        $deviceCounter++
        Write-Progress -Id 2 -Activity "Bulk Set GUI Password" -Status "$($selecteddevice.DeviceName) ($deviceCounter of $($SelectedDevices.Count))" -PercentComplete ([int](100 * $deviceCounter / $SelectedDevices.Count))
        try {
            Send-RemoteCommand $selecteddevice
            #$result.result.result | Select-Object Id,@{Name="Status"; Expression={$_.Result.code}},@{Name="Message"; Expression={$_.Result.Message}} | Format-Table
            Write-Host " $($result.result.result.id) $($result.result.result.result.code)" -ForegroundColor DarkGray
        }catch{
            Write-Host "  $($selecteddevice.DeviceName) ($($selecteddevice.AccountID)): FAILED - $($_.Exception.Message)" -ForegroundColor Red
        }

        }
        Write-Progress -Id 2 -Activity "Bulk Set GUI Password" -Completed
        $SecurePassword = $null; $UnsecureGUIPassword = $null   ## drop the plaintext/SecureString references now that the loop is done

        Write-Host "  Done." -ForegroundColor Green

    }