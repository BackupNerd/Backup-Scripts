<# ----- About: ----
    # N-able | Cove Data Protection | Get Recovery Test Screenshot
    # Revision v01 - 2024-11-04 - Prototype
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com    
    # Reddit https://www.reddit.com/r/Nable/
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
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
    # For use with the Standalone edition of N-able | Cove Data Protection 
    # Some script elements may be developed, tested or documented using AI
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select devices
    # Optionally export to XLS/CSV
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -AllDevices switch parameter to skip GUI device selection
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -Export switch parameter to export statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm
# -----------------------------------------------------------#>  ## Behavior

<# ----- Example Execution: ----
    # This section provides examples of how to execute the script from a command prompt or PowerShell prompt with various parameters.
    #
    # Example 1: Run the script with default parameters
    # Command Prompt (cmd.exe):
    #    powershell -File .\GetRecoveryTestScreenshot.prototype.v01.ps1
    # PowerShell: 
    #    .\GetRecoveryTestScreenshot.prototype.v01.ps1
    #
    # Example 3: Run the script and Save the screenshot of a specific device
    # PowerShell:
    #    .\GetRecoveryTestScreenshot.prototype.v01.ps1 -DefaultPartner SEDEMO -deviceid 3453752
    #
    # Note: Ensure that the script is unblocked and the execution policy is set correctly before running the script.
# -----------------------------------------------------------#>  ## Example Execution

<# ----- Troubleshooting: ----
    # If you encounter issues running the script, ensure that the script is unblocked and the execution policy is set correctly.
    # You may need to login as an adminsitrator to perform these tasks.
    #
    # To unblock the script:
    # 1. Right-click the script file in File Explorer and select 'Properties'.
    # 2. In the 'General' tab, check the 'Unblock' checkbox if it is present.
    # 3. Click 'Apply' and then 'OK'.
    #
    # Alternatively, you can unblock the script using PowerShell:
    # 1. Open PowerShell as an administrator.
    # 2. Run the following command to unblock the script:
    #    Unblock-File -Path "C:\Path\To\Your\Script.ps1"
    #
    # To set the execution policy to allow scripts to run:
    # 1. Open PowerShell as an administrator.
    # 2. Run the following command to set the execution policy:
    #    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
    # 3. If prompted, confirm the change by typing 'Y' and pressing Enter.
    #
    # Note: Setting the execution policy to 'Unrestricted' allows all scripts to run, which is less secure but can be useful for troubleshooting.
    #       Alternatively, you can use 'Bypass' to completely bypass the execution policy for the current session:
    #       Run the following command to set the execution policy to 'Bypass':
    #       Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
    #       This setting is temporary and only applies to the current PowerShell session.
# -----------------------------------------------------------#>  ## Troubleshooting

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [Int]$Days = 30,                              ## Number of days to search for active devices
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 2000,                     ## Specify the maximum number of devices results to return
        [Parameter(Mandatory=$False)] [string]$DefaultPartner,                      ## Specify a default partner name (case sensitive)
        [Parameter(Mandatory=$False)] [Int]$deviceid,                               ## Specify a default partner name (case sensitive)
        [Parameter(Mandatory=$False)] [switch]$GridView = $true,                    ## Display output via Powershell Out-Gridview 
        [Parameter(Mandatory=$False)] [switch]$Export = $true,                      ## Save screenshot to file
        [Parameter(Mandatory=$False)] [switch]$Launch = $false,                      ## Launch screenshot
        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",        ## Export Path (Script location is the default)
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                     ## Remove stored API credentials at script run
    )   


#region ----- Environment, Variables, Names and Paths ----
Clear-Host
#Requires -Version 5.1
$ConsoleTitle   = "Cove Data Protection | Get Boot Screenshot"                ## Update with full script name
$ShortTitle     = "BootScreen"                                                      ## Update with short script name
$host.UI.RawUI.WindowTitle = $ConsoleTitle

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Push-Location $dir

Write-Output "  $ConsoleTitle`n`n$ScriptPath"
$Syntax = Get-Command $PSCommandPath -Syntax
Write-Output "`n  Script Parameter Syntax:`n`n  $Syntax"
   
$CurrentDate    = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
$ShortDate      = Get-Date -format "yyyy-MM-dd"

if ($ExportPath) {
    $ExportPath = Join-path -path $ExportPath -childpath "$($ShortTitle)_$($ShortDate)" 
}else{
    $ExportPath = Join-path -path $dir -childpath "$($ShortTitle)_$($ShortDate)"
}

If ($Export -and $ExportPath) {mkdir -force -path $ExportPath | Out-Null}

Write-Output "  Current Parameters:"
Write-Output "  -DefaultPartner = $DefaultPartner"
Write-Output "  -Days           = $Days"
Write-Output "  -DeviceCount    = $DeviceCount"
Write-Output "  -ColumnCode     = $ColumnCode"
Write-Output "  -CustomColumn   = $CustomColumn"
Write-Output "  -Gridview       = $GridView"
Write-Output "  -Export         = $Export"
Write-Output "  -Launch         = $Launch"
Write-Output "  -ExportPath     = $ExportPath"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------"
$urlJSON = 'https://api.backup.management/jsonapi'

$Script:MXB_Path = "C:\ProgramData\MXB\"
if (-not (Test-Path -Path $Script:MXB_Path)) {
    New-Item -ItemType Directory -Path $Script:MXB_Path -Force | Out-Null
}
$Script:APIcredfile = join-path -Path $Script:MXB_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
$Script:APIcredpath = Split-path -path $APIcredfile

#$filterDate = (Get-Date).AddDays(-$Days)
#$counter = 1

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----
#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
    Write-Output $Script:strLineSeparator     
    Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"

    DO{ $PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($partnerName.length -eq 0)
    
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

    $BackupCred.UserName | Out-file -append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
    
    Authenticate  ## Attempt API Authentication

} ## Set API credentials if not present

Function Get-APICredentials {

    $Script:MXB_Path = "C:\ProgramData\MXB\"
    $Script:APIcredfile = join-path -Path $Script:MXB_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-path -path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential File Cleared"
        Authenticate  ## Retry Authentication
        
    }else{ 
        Write-Output $Script:strLineSeparator 
        Write-Output "  Getting Backup API Credentials" 
    
        if (Test-Path $APIcredfile) {
            Write-Output    $Script:strLineSeparator        
            "  Backup API Credential File Present"
            $APIcredentials = get-content $APIcredfile
            
            $Script:cred0 = [string]$APIcredentials[0] 
            $Script:cred1 = [string]$APIcredentials[1]
            $Script:cred2 = $APIcredentials[2] | Convertto-SecureString 
            $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))

            Write-Output    $Script:strLineSeparator 
            Write-Output "  Stored Backup API Partner  = $Script:cred0"
            Write-Output "  Stored Backup API User     = $Script:cred1"
            Write-Output "  Stored Backup API Password = Encrypted"
            
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-Output "  Backup API Credential File Not Present"

            Set-APICredentials  ## Create API Credential File if Not Found
        }
    }
} ## Get API credentials if present

Function Authenticate {
    Get-APICredentials  ## Read API Credential File before Authentication

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    #$data.params.partner = $Script:cred0
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Authenticate = $webrequest | convertfrom-json

    #Debug Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($authenticate.visa) { 
        $Script:visa = $authenticate.visa
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }
} ## Use Backup.Management credentials to Authenticate

Function Get-VisaTime {
    if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
            Authenticate
        }
    }
} ## Check Visa Time and Authenticate again if needed

#endregion ----- Authentication ----


#region ----- Data Conversion ----
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
} ## Convert epoch time to date time 
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
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json

    if (($Partner.result.result.Level -ne "Root") -and ($Partner.result.result.Level -ne "Subroot") -and ($Partner.result.result.Level -ne "Distributor")) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$script:Level = $Partner.result.result.Level
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-Output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
    }else{
        Write-Output $Script:strLineSeparator
        Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
    }

    if ($partner.error) {
        Write-Output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername

    }

} ## get PartnerID and Partner Level

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
    $data.params.query.Filter = "AT==1 AND TS > $days.days().ago()"
    $data.params.query.Columns = @("AU","TS","AR","AN","MN","AL","CD","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
    $data.params.query.OrderBy = "AR ASC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = $DeviceCount
    $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")

    $webrequest = Invoke-WebRequest -Method POST `
    -ContentType 'application/json' `
    -Body (ConvertTo-Json $data -depth 6) `
    -Uri $url `
    -SessionVariable global:websession `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession

    #Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

    $Script:Devices = $webrequest | convertfrom-json
    $Script:visa = $Devices.visa

    if ($Partner.result.result.Uid) {
        [String]$Script:PartnerId = $Partner.result.result.Id
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-Output "  Searching  $PartnerName "
        Write-Output $Script:strLineSeparator
        }else{
        Write-Output $Script:strLineSeparator
        Write-Host "PartnerName or UID Not Found"
        Write-Output $Script:strLineSeparator
        }
         

    Write-Output "  Requesting boot screenshots for $($Devices.result.result.count) devices."
    Write-Output "  Please be patient, this could take some time."
    Write-Output $Script:strLineSeparator

    ForEach ( $DeviceResult in $Devices.result.result ) {

        Get-VisaTime   
        Write-output "Trying $($DeviceResult.AccountId)"
        Get-BootScreenshot $DeviceResult.AccountId
        }
    }   


Function Get-BootScreenshot ($deviceid) {
     $params1 = @{
          Uri         = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?filter[backup_cloud_device_id.eq]=$deviceid&filter[type.in]=RECOVERY_TESTING,SELF_HOSTED,AZURE_SELF_HOSTED,ESXI_SELF_HOSTED"
          Method      = 'GET'
          Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
          WebSession  = $websession
          ContentType = 'application/json'
     }   

     $a = Invoke-Restmethod @params1
     #$($a.data.attributes.last_boot_test_session_id)
     if ($a.data[-1].attributes.last_boot_test_screenshot_presented -eq $true  ) {   
          $ContinuityType = $a.data[-1].attributes.type
          $ContinuityBootStatus = $a.data[-1].attributes.last_boot_test_status
          $PartnerName = $a.data.attributes.backup_cloud_partner_name  
          $ComputerName = $a.data.attributes.backup_cloud_device_machine_name 
          if ($a.data[-1].attributes.last_boot_test_backup_session_timestamp) {
                $BackupTime = Convert-UnixTimeToDateTime $($a.data[-1].attributes.last_boot_test_backup_session_timestamp[0])
          }

        $params2 = @{
            Uri         = "https://api.backup.management/draas/actual-statistics/v1/sessions/$($a.data[-1].attributes.last_boot_test_session_id)/files/?filter[file_type.in]=screenshot"
            Method      = 'GET'
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $websession
            ContentType = 'application/json'
        }   

        $b = Invoke-Restmethod @params2
        #$($b.data.id)
        if ($b.data.id -and $a.data[-1].attributes.last_boot_test_session_id) {

            $params3 = @{
                Uri         = "https://api.backup.management/draas/actual-statistics/v1/sessions/$($a.data[-1].attributes.last_boot_test_session_id)/files/$($b.data.id)/get-temporary-url/"
                Method      = 'POST'
                Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
                WebSession  = $websession
                ContentType = 'application/json'
            }   

            $c = Invoke-Restmethod @params3
            #$($c.data.attributes.url)
            if ($c.data.attributes.url -ne $null) {
         
                $BackupTimestamp = Convert-UnixTimeToDateTime $($a.data[-1].attributes.last_boot_test_backup_session_timestamp[0])
                $FormattedTimestamp = $BackupTimestamp.ToString("yyyy-MM-dd_HH-mm-ss")
                $ScreenshotFileName = "$ExportPath\$Partnername $($ComputerName) $($a.data[-1].attributes.backup_cloud_device_id) $($ContinuityBootStatus) $FormattedTimestamp.png"
                Invoke-WebRequest -Uri $($c.data.attributes.url) -OutFile $ScreenshotFileName 
                if ($launch) { Start-Process $ScreenshotFileName }

            }

        }
    }else{Write-Output "No Boot Screenshot Available for $($DeviceResult.AccountId)"}
}

#endregion ----- Functions ----

#$deviceid = 3386098	

Authenticate

if ($DefaultPartner) {
    Send-GetPartnerInfo $DefaultPartner
}else{
    Send-GetPartnerInfo $cred0
}


if ($deviceid) {
    Get-BootScreenshot $deviceid
}else{  
    Send-GetDevices
}    

