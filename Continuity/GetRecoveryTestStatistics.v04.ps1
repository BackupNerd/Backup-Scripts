<# ----- About: ----
    # N-able | Cove Data Protection | Get Recovery Test Statistics 
    # Revision v04 - 2024-11-04
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
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
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
    #    powershell -File .\GetRecoveryTestStatistics.v04.ps1
    # PowerShell: 
    #    .\GetRecoveryTestStatistics.v04.ps1
    #
    # Example 3: Run the script and export the statistics to a specified path
    # PowerShell:
    #    .\GetRecoveryTestStatistics.v04.ps1 -Export -ExportPath "C:\Exports"
    #
    # Example 3: Run the script with both an export path and skip GUI partner / device selection
    # PowerShell:
    #    .\GetRecoveryTestStatistics.v04.ps1 -ExportPath "C:\Exports" -AllPartners -AllDevices
    #
    # Example 4: Run the script and launch the XLS/CSV file after completion
    # PowerShell:
    #    .\GetRecoveryTestStatistics.v04.ps1 -Launch
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
        [Parameter(Mandatory=$False)] [switch]$AllPartners,                         ## Skip partner selection
        [Parameter(Mandatory=$False)] [switch]$AllDevices,                          ## Skip device selection
        [Parameter(Mandatory=$False)] [switch]$DeviceName,                          ## Report on named Device
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 5000,                     ## Change Maximum Number of devices results to return
        [Parameter(Mandatory=$False)] [switch]$Export = $true,                      ## Generate CSV / XLS Output Files
        [Parameter(Mandatory=$False)] [switch]$Launch,                              ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                     ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",                ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                     ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Recovery Test Statistics "
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $Script:scriptpath = $MyInvocation.MyCommand.Path
    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $dir = Split-Path $scriptpath
    Push-Location $dir
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    if ($ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "DR_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "DR_$shortdate"}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urljson = "https://api.backup.management/jsonapi"

    Write-output "  Current Parameters:"
    Write-output "  -AllPartners     = $AllPartners"
    Write-output "  -AllDevices      = $AllDevices"
    Write-output "  -DeviceCount     = $DeviceCount"
    Write-output "  -Export          = $Export"
    Write-output "  -Launch          = $Launch"
    Write-output "  -ExportPath      = $ExportPath"
    Write-output "  -Delimiter       = $Delimiter"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove | Backup.Management API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

    $BackupCred.UserName | Out-file -append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
    
    Start-Sleep -milliseconds 300

    AuthenticateCDP  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    if (-not (Test-Path -Path $Script:True_path)) {
        New-Item -ItemType Directory -Path $Script:True_path -Force
    }
    $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-Path -Path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential File Cleared"
        AuthenticateCDP  ## Retry Authentication
        
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
                Write-output "  Stored Backup API Partner  = $Script:cred0"
                Write-output "  Stored Backup API User     = $Script:cred1"
                Write-output "  Stored Backup API Password = Encrypted`n"
                
            }else{
                Write-Output    $Script:strLineSeparator 
                Write-Output "  Backup API Credential File Not Present"

                Set-APICredentials  ## Create API Credential File if Not Found
                }
            }

}  ## Get API credentials if present

Function AuthenticateCDP {

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

    #Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($authenticate.visa) { 

        $Script:visa = $authenticate.visa
        $script:UserId = $authenticate.result.result.id

        if (($authenticate.result.result.flags -notcontains "securityofficer") -and 
            ($CleanupSharedBillable -or $CleanupUnlicensed -or $CleanupDeletedBillable -or $CleanupDeletedShared -or $CleanupDeletedBillableShared)
            ) { Write-warning "Security Office Role not found, Deletion of Historic Backup data not allowed!"}
            
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-Warning "`n  $($authenticate.error.message)"
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"

        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Visa-Check {
     if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){   
            AuthenticateCDP
        }
    }
}  ## Recheck remaining Visa time and reauthenticate

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

Function Save-CSVasExcel {
    param (
        [string]$CSVFile = $(Throw 'No file provided.')
    )
    
    BEGIN {
        function Resolve-FullPath ([string]$Path) {    
            if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) {
                # $Path = Join-Path (Get-Location) $Path
                $Path = "$PWD\$Path"
            }
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
        
        # can comment out this part if you don't care to have the columns autosized
        $ws = $wb.Worksheets.Item(1)
        $range = $ws.UsedRange
        [void]$range.AutoFilter()
        [void]$range.EntireColumn.Autofit()

        $num = 1
        $dir = Split-Path $xlOut
        $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
        $nextname = $xlOut
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }

        $wb.SaveAs($nextname, 51)
    }

    END {
        $xl.Quit()
    
        $null = $ws, $wb, $xl | ForEach-Object {Release-Ref $_}

        # del $CSVFile
    }
} ## Save as output XLS Routine

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
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Partner = $webrequest | convertfrom-json

        $RestrictedPartnerLevel = @("Root","Sub-root","Distributor")

        if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername
            }

        if ($partner.error) {
            write-output "  $($partner.error.message)"
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername

        }

    } ## get PartnerID and Partner Level    

    Function CallJSON($url,$object) {

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($object)
        $web = [System.Net.WebRequest]::Create($url)
        $web.Method = "POST"
        $web.ContentLength = $bytes.Length
        $web.ContentType = "application/json"
        $stream = $web.GetRequestStream()
        $stream.Write($bytes,0,$bytes.Length)
        $stream.close()
        $reader = New-Object System.IO.Streamreader -ArgumentList $web.GetResponse().GetResponseStream()
        return $reader.ReadToEnd()| ConvertFrom-Json
        $reader.Close()
    }

    Function Send-EnumeratePartners {
        # ----- Get Partners via EnumeratePartners -----
        
        # (Create the JSON object to call the EnumeratePartners function)
            $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
                Add-Member -PassThru NoteProperty visa $Script:visa |
                Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
                Add-Member -PassThru NoteProperty params @{
                                                            parentPartnerId = $PartnerId 
                                                            fetchRecursively = "true"
                                                            fields = (0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22) 
                                                            } |
                Add-Member -PassThru NoteProperty id '1')| ConvertTo-Json -Depth 5
        
        # (Call the JSON Web Request Function to get the EnumeratePartners Object)
                [array]$Script:EnumeratePartnersSession = CallJSON $urlJSON $objEnumeratePartners
        
                $Script:visa = $EnumeratePartnersSession.visa
        
                #Write-Output    $Script:strLineSeparator
                #Write-Output    "  Using Visa:" $Script:visa
                #Write-Output    $Script:strLineSeparator
        
        # (Added Delay in case command takes a bit to respond)
                Start-Sleep -Milliseconds 100
        
        # (Get Result Status of EnumerateAccountProfiles)
                $EnumeratePartnersSessionErrorCode = $EnumeratePartnersSession.error.code
                $EnumeratePartnersSessionErrorMsg = $EnumeratePartnersSession.error.message
        
        # (Check for Errors with EnumeratePartners - Check if ErrorCode has a value)
                if ($EnumeratePartnersSessionErrorCode) {
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  EnumeratePartnersSession Error Code:  $EnumeratePartnersSessionErrorCode"
                    Write-Output    "  EnumeratePartnersSession Message:  $EnumeratePartnersSessionErrorMsg"
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  Exiting Script"
        # (Exit Script if there is a problem)
        
                    #Break Script
                }
                    Else {
        # (No error)
        
                $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
                
                $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
                $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
                $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
            
                $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                        
                $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
                
                
                if ($AllPartners) {
                    $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  All Partners Selected"
                }else{
                    $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single
            
                    if($null -eq $Selection) {
                        # Cancel was pressed
                        # Run cancel script
                        Write-Output    $Script:strLineSeparator
                        Write-Output    "  No Partners Selected"
                        Break
                    }
                    else {
                        # OK was pressed, $Selection contains what was chosen
                        # Run OK script
                        [int]$script:PartnerId = $script:Selection.Id
                        [String]$script:PartnerName = $script:Selection.Name
                    }
                }

        }
        
    }  ## EnumeratePartners API Call

    Function Get-DRStatistics  {
        Param (
            [Parameter(Mandatory=$False)][Int]$PartnerId,
            [Parameter(Mandatory=$False)][string]$Devicename
        ) #end param

    $Script:fields = @(
        "agent_id",
        "backup_cloud_device_alias",
        "backup_cloud_device_id",
        "backup_cloud_device_machine_name",
        "backup_cloud_device_machine_os_type",
        "backup_cloud_device_name",
        "backup_cloud_device_status",
        "backup_cloud_partner_id",
        "backup_cloud_partner_name",
        "colorbar",
        "current_recovery_status",
        "data_sources",
        "device_boot_frequency",
        "device_recovery_frequency",
        "last_backup_session_timestamp",
        "last_boot_test_screenshot_presented",
        "last_boot_test_status",
        "last_recovery_duration_user",
        "last_recovery_errors_count",
        "last_recovery_restored_files_count",
        "last_recovery_restored_size",
        "last_recovery_selected_files_count",
        "last_recovery_selected_size",
        "last_recovery_session_id",
        "last_recovery_status",
        "last_recovery_timestamp",
        "plan_device_id",
        "plan_name",
        "recovery_agent_name",
        "recovery_agent_state",
        "recovery_session_progress",
        "recovery_target_azure_vm_size",
        "recovery_target_device_number_of_cpu",
        "recovery_target_device_ram_size_mb",
        "recovery_target_dns_server",
        "recovery_target_enable_replication_service",
        "recovery_target_gateway",
        "recovery_target_local_speed_vault",
        "recovery_target_lsv_path",
        "recovery_target_subnet_mask",
        "recovery_target_type",
        "recovery_target_vhd_path",
        "recovery_target_vm_address",
        "recovery_target_vm_virtual_switch",
        "recovery_target_vmware_host",
        "region_name",
        "type"
    )

        #$Script:url2 = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?offset=0&limit=200&fields=$($fields -join ",")&sort=last_recovery_timestamp&filter%5Btype%5D=RECOVERY_TESTING&filter%5Bpartner_materialized_path.contains%5D=/$($PartnerId)/"
        ## Recovey Testing Only

        #$Script:url2 = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?offset=0&limit=200&search=$DeviceName&fields=$($fields -join ",")&sort=last_recovery_timestamp&filter%5Bpartner_materialized_path.contains%5D=/$($PartnerId)/"
        ## Recovery Testing with a Device Name Filter

        $Script:url2 = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?offset=0&limit=200&search=$DeviceName&fields=$($fields -join ",")&sort=last_recovery_timestamp&filter%5Bpartner_materialized_path.contains%5D=/$($PartnerId)/"
        ## All Continuity Types

        $method = 'GET'

        $params = @{
            Uri         = $url2
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:DRStatisticsResponse = Invoke-RestMethod @params 

        Write-output  "$url2"

        $Script:DRStatistics = $Script:DRStatisticsResponse.data.attributes | Select-object $fields
        
        $Script:DRStatistics | foreach-object { $_.last_recovery_selected_size = [Math]::Round([Decimal]($($_.last_recovery_selected_size) /1GB),2) }
        $Script:DRStatistics | foreach-object { $_.last_recovery_restored_size = [Math]::Round([Decimal]($($_.last_recovery_restored_size) /1GB),2) }
        $Script:DRStatistics | foreach-object { $_.last_recovery_timestamp = Convert-UnixTimeToDateTime $($_.last_recovery_timestamp) }
        $Script:DRStatistics | foreach-object { $_.last_backup_session_timestamp = Convert-UnixTimeToDateTime $($_.last_backup_session_timestamp) }
        $Script:DRStatistics | foreach-object { 
            if ($_.device_boot_frequency -eq -1) {
            $_.device_boot_frequency = "N/A"
            } elseif ($_.device_boot_frequency -eq 0) {SEDEMPP
            $_.device_boot_frequency = "Each Session"
            } else {
            $days = [math]::Floor($_.device_boot_frequency / 86400)
            $_.device_boot_frequency = "$days Day"
            }
        }
        $Script:DRStatistics | foreach-object {
            $totalSeconds = [int]$_.last_recovery_duration_user
            $hours = [math]::Floor($totalSeconds / 3600)
            $minutes = [math]::Floor(($totalSeconds % 3600) / 60)
            $seconds = $totalSeconds % 60
            $_.last_recovery_duration_user = "{0}h {1}m {2}s" -f $hours, $minutes, $seconds
        }

    } 


#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $switch = $PSCmdlet.ParameterSetName

    AuthenticateCDP

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Send-GetPartnerInfo $Script:cred0

    if ($AllPartners) {}else{Send-EnumeratePartners}

    Get-DRStatistics $partnerId

    $selectFields = @(
        "agent_id",
        "backup_cloud_device_alias",
        "backup_cloud_device_id",
        "backup_cloud_device_machine_name",
        "backup_cloud_device_machine_os_type",
        "backup_cloud_device_name",
        "backup_cloud_device_status",
        "backup_cloud_partner_id",
        "backup_cloud_partner_name",
        "colorbar",
        "current_recovery_status",
        "data_sources",
        "device_boot_frequency",
        "device_recovery_frequency",
        "last_backup_session_timestamp",
        "last_boot_test_screenshot_presented",
        "last_boot_test_status",
        "last_recovery_duration_user",
        "last_recovery_errors_count",
        "last_recovery_restored_files_count",
        "last_recovery_restored_size",
        "last_recovery_selected_files_count",
        "last_recovery_selected_size",
        "last_recovery_session_id",
        "last_recovery_status",
        "last_recovery_timestamp",
        "plan_device_id",
        "plan_name",
        "recovery_agent_name",
        "recovery_agent_state",
        "recovery_session_progress",
        "recovery_target_azure_vm_size",
        "recovery_target_device_number_of_cpu",
        "recovery_target_device_ram_size_mb",
        "recovery_target_dns_server",
        "recovery_target_enable_replication_service",
        "recovery_target_gateway",
        "recovery_target_local_speed_vault",
        "recovery_target_lsv_path",
        "recovery_target_subnet_mask",
        "recovery_target_type",
        "recovery_target_vhd_path",
        "recovery_target_vm_address",
        "recovery_target_vm_virtual_switch",
        "recovery_target_vmware_host",
        "region_name",
        "type"
    )

    if ($AllDevices) {
        $script:SelectedDevices = $Script:DRStatistics | Select-object $selectFields
    }else{
        $script:SelectedDevices = $Script:DRStatistics | Select-object $selectFields | Out-GridView -title "Current Partner | $partnername" -OutputMode Multiple}

    if($null -eq $SelectedDevices) {
        # Cancel was pressed
        # Run cancel script
        Write-Output    $Script:strLineSeparator
        Write-Output    "  No Devices Selected"
        Break
    }else{
        # OK was pressed, $Selection contains what was chosen
        # Run OK script
        $script:SelectedDevices |  Select-Object $selectFields | Sort-object PartnerName,AccountId | format-table

        If ($Script:Export) {
            mkdir -force -path $ExportPath | Out-Null
            $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_ContinuityStats_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
            $Script:SelectedDevices | Select-object * | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8}
            
    }

    ## Generate XLS from CSV
    
    if ($csvoutputfile) {
        $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile
    }
    Write-output $Script:strLineSeparator

    ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
        
    if ($Launch) {
        If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
            Start-Process "$xlsoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
            }else{
            Start-Process "$csvoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening CSV file"
            Write-output $Script:strLineSeparator            
            }
        }
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $Script:csvoutputfile"
    Write-Output "  XLS Path = $Script:xlsoutputfile"
    Write-Output ""

    Start-Sleep -seconds 10