<# ----- About: ----
    # N-able | Cove Data Protection | Get Recovery Test Statistics for single device
    # Revision v04 - 2023-06-15
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
    # For use with the Standalone edition of N-able | Cove Data Protection 
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Authenticate to https://backup.management console
    #
    # The following input variable should be passed from your RMM tool.
    #
    # $AdminLoginPartnerName    ## Partner Name for user authenticating to the Cove Console
    # $AdminLoginUserName       ## User Name for user authenticating to the Cove Console
    # $PlainTextAdminPassword   ## Password for user authenticating to the Cove Console
    #
    # Lookup device name
    #
    # $DeviceName               ## Device name to lookup, if not specified at run time the device name will be pulled from a local file.
    #
    # Request Continuity Statistics -Excluding One-time Restores
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
# -----------------------------------------------------------#>  ## Behavior

#region ----- Environment, Variables, Names and Paths ----

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------"
$urljson = "https://api.backup.management/jsonapi"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function BasicAuthenticate {

if ($AdminLoginPartnerName -eq $null) {$AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"}
if ($AdminLoginUserName -eq $null) {$AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able | Cove Backup.Management API"}

if ($PlainTextAdminPassword -eq $null) {
    $AdminPassword = Read-Host -AsSecureString "  Enter Password for N-able | Cove Backup.Management API"
    $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword)) # (Convert SecureString Password to plain text)
    }

# (Show credentials for Debugging)
Write-Output "  Logging on with the following Credentials`n"
Write-Output "  PartnerName:  $AdminLoginPartnerName"
Write-Output "  UserName:     $AdminLoginUserName"
Write-Output "  Password:     It's secure..."

# (Create authentication JSON object using ConvertTo-JSON)
$objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
Add-Member -PassThru NoteProperty method 'Login' |
Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$AdminLoginUserName;password=$PlainTextAdminPassword}|
Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json

# (Call the JSON function with URL and authentication object)
$script:session = CallJSON $urlJSON $objAuth
Start-Sleep -Milliseconds 100

# (Variable to hold current visa and reused in following routines)
$script:visa = $session.visa
$script:PartnerId = [int]$session.result.result.PartnerId
    
# (Get Result Status of Authentication)
$AuthenticationErrorCode = $Session.error.code
$AuthenticationErrorMsg = $Session.error.message

# (Check if ErrorCode has a value)
If ($AuthenticationErrorCode) {
    Write-Output "Authentication Error Code:  $AuthenticationErrorCode"
    Write-Output "Authentication Error Message:  $AuthenticationErrorMsg"
    Pause
    Break Script
}   ## (Exit Script if there is a problem)
Else{

    } ## (Action if no error)

}  ## Use Backup.Management credentials to Authenticate

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

#region ----- Backup.Management JSON Calls ----

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

Function Get-DRStatistics ([Int]$PartnerId,[string]$Devicename ) {
 
    $Script:fields  = "backup_cloud_device_id,plan_device_id,backup_cloud_partner_id,last_recovery_session_id,current_recovery_status,last_recovery_status,last_recovery_timestamp,last_recovery_duration_user,plan_name,backup_cloud_partner_name,backup_cloud_device_name,backup_cloud_device_machine_name,region_name,type,recovery_target_type,recovery_agent_name,backup_cloud_device_status,agent_id,last_boot_test_session_id,last_boot_test_status,last_boot_test_recovery_session_timestamp,device_boot_frequency,backup_cloud_device_name,backup_cloud_device_machine_name,backup_cloud_partner_name,plan_name,colorbar,current_recovery_status,last_recovery_errors_count,last_recovery_timestamp,last_recovery_duration_user,last_boot_test_status,last_boot_test_screenshot_presented,device_boot_frequency,last_recovery_restored_files_count,last_recovery_selected_files_count,data_sources,last_recovery_status,last_recovery_restored_size,last_recovery_selected_size,backup_cloud_device_machine_os_type,recovery_session_progress,recovery_agent_state,recovery_target_type,recovery_target_vm_virtual_switch,recovery_target_vhd_path,recovery_target_local_speed_vault,recovery_target_lsv_path,recovery_target_enable_replication_service,recovery_target_vm_address,recovery_target_subnet_mask,recovery_target_gateway,recovery_target_dns_server,recovery_target_enable_machine_boot,recovery_agent_name,device_recovery_frequency,backup_cloud_device_status,backup_cloud_device_alias"
    $Script:url2 = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?offset=0&limit=20&search=$DeviceName&fields=$fields&sort=last_recovery_timestamp&filter%5Bpartner_materialized_path.contains%5D=/$($PartnerId)/"
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

    $Script:DRStatistics = $Script:DRStatisticsResponse.data.attributes | Select-object * 
    $Script:DRStatistics | foreach-object { $_.last_recovery_selected_size = [Math]::Round([Decimal]($($_.last_recovery_selected_size) /1GB),2) }
    $Script:DRStatistics | foreach-object { $_.last_recovery_restored_size = [Math]::Round([Decimal]($($_.last_recovery_restored_size) /1GB),2) }
    $Script:DRStatistics | foreach-object { $_.last_recovery_timestamp = Convert-UnixTimeToDateTime $($_.last_recovery_timestamp) }
    $Script:DRStatistics | foreach-object { $_.last_boot_test_recovery_session_timestamp = Convert-UnixTimeToDateTime $($_.last_boot_test_recovery_session_timestamp) }
    if ($Script:DRStatistics.colorbar) { 
        $Script:DRStatistics.colorbar | foreach-object { $_.backup_session_timestamp = Convert-UnixTimeToDateTime $($_.backup_session_timestamp ) }
        $Script:DRStatistics.colorbar | foreach-object { $_.recovery_session_timestamp = Convert-UnixTimeToDateTime $($_.recovery_session_timestamp ) }
    }
} 

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

#region ----- Body ----

BasicAuthenticate

if ($DeviceName) {
}else{
    $Script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 
    $DeviceName = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
}

Get-DRStatistics $partnerId $DeviceName

    $script:SelectedDevices = $Script:DRStatistics | Where-Object {($_.plan_name -ne "Azure On Demand (A specific backup session)") -and ($_.backup_cloud_device_name -eq $devicename) } | Select-object *

if($null -eq $SelectedDevices) {

    Write-Output    $Script:strLineSeparator
    Write-Warning   "No Continuity History Found for Device | $DeviceName"
    Break
}else{
    if     ($script:SelectedDevices.plan_name -eq "Recovery Testing (Biweekly)") { $script:SelectedDevices | Select-Object plan_name,backup_cloud_device_name,last_recovery_timestamp,* -ExcludeProperty *Azure*,*target*,*lsv*,colorbar -ea SilentlyContinue | format-list }
    elseif ($script:SelectedDevices.plan_name -eq "Recovery Testing (Monthly)") { $script:SelectedDevices | Select-Object plan_name,backup_cloud_device_name,last_recovery_timestamp,* -ExcludeProperty *Azure*,*target*,*lsv*,colorbar -ea SilentlyContinue | format-list }
    elseif ($script:SelectedDevices.plan_name -eq "Standby Image (Hyper-V)") { $script:SelectedDevices | Select-Object plan_name,backup_cloud_device_name,last_recovery_timestamp,* -ExcludeProperty *Azure*,colorbar -ea SilentlyContinue | format-list }
    elseif ($script:SelectedDevices.plan_name -eq "Standby Image (Azure)") { $script:SelectedDevices | Select-Object plan_name,backup_cloud_device_name,last_recovery_timestamp,* -excludeProperty *lsv*,colorbar -ea SilentlyContinue  | format-list }
    $script:SelectedDevices.colorbar | Sort-Object recovery_session_timestamp | Select-Object * -ExcludeProperty session_id -ErrorAction SilentlyContinue  
}

#endregion ----- Body ----
