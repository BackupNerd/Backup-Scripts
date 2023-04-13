<# ----- About: ----
    # N-able | Cove Data Protection Monitor 1 - Application
    # Revision v01 - 2022-03-28
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

<# ----- Behavior: ----
    #  Get Cover service status
    #  Get Cove installation \ Process status \ details
    #  Get Cove installed version
    #  Get Cove license type
    #  Get Cove integration status   
    #  Get Cove process uptime    
    #  Get Cove process CPU usage      
    #  Get Cove process RAM usage
    #  Get Cove LSV status \ sync
    #  Get Cove cloud status \ sync    
    #  Get Cove discovered data sources
    #  Get Cove Protected data sources        
    #  Get Cove unprotected data sources         
    #  Get Cove selected\used storage  
 # -----------------------------------------------------------#>

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [switch]$DebugDetail,              ## Set True and Disconnect NIC to test Debug Data scenarios,    
        [Parameter(Mandatory=$False)] [switch]$ThresholdValue = $true,   ## Set True to Output Threshold Values seperately for Debugging
        [Parameter(Mandatory=$False)] [switch]$GlobalThresholdValue = $true    ## Set True to Output Threshold Values seperately for Debugging
    )        
Clear-Host
#Requires -Version 5.1 -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection - Application Monitor"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

#region Functions
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
}  ## Convert epoch time to date time 
Function RND {
    Param(
    [Parameter(ValueFromPipeline,Position=3)]$Value,
    [Parameter(Position=0)][string]$unit = "MB",
    [Parameter(Position=1)][int]$decimal = 2
    )
    "$([math]::Round(($Value/"1$Unit"),$decimal)) $Unit"

<# Usage Examples

1.23123123123123123 | RND '' 6
1.231231 

RND KB 2 234234234.234234234234
228744.37 KB

234234234.234234234234 | RND KB 2
228744.37 KB

234234234.234234234234 | RND MB 4
223.3832 MB

234234234.234234234234 | RND GB 1
0.2 GB

234234234.234234234234 | RND 
0.22 GB

234234234.234234234234 | RND MB 0
223 MB

1223234234234.234234234234 | RND TB 2
1.11 TB

write-output "12312312312123.123123123" | RND KB
12023742492.31 KB

write-output "TEST $(12312312312123.12312312 | RND GB 2)" 
TEST 12023742492.31 KB
#>

} ## Rounding function for B,KB,MB,GB,TB

Function Get-BackupService {
 
    $Script:BackupService = get-service -name $BackupServiceName -ea SilentlyContinue
    Write-output "`n[Backup Service Controller]"
    if ($BackupService.status -eq "Running"){
        $Global:BackupServiceOutVal = 0
        $Global:BackupServiceOutTxt = $BackupService.status 
        Write-output  "Service Status             : $BackupServiceOutTxt"
    }elseif (($BackupService.status -ne "Running") -and ($null -ne $BackupService.status )){
        $Global:BackupServiceOutVal = 2
        $Global:BackupServiceOutTxt = $BackupService.status 
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }elseif ($null -eq $BackupService.status ){
        $Global:BackupServiceOutVal = 1
        $Global:BackupServiceOutTxt = "Not Installed"
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }
    
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Service Value" $Global:BackupServiceOutVal
        Write-Output "-- Service Text" $Global:BackupServiceOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }

} ## Get Backup Service Status

    if ($DebugDetail) {
        Write-output "`n####### Start Debug Detail #######"
        $BackupServiceName = "Incorrect Backup Service Controller Name";  Get-BackupService
        $BackupServiceName = "Backup Service Controller"; Stop-service -name "$BackupServiceName" -Force; start-sleep -Seconds 10; Get-BackupService
        $BackupServiceName = "Backup Service Controller"; Start-service -name "$BackupServiceName"; start-sleep -Seconds 10; Get-BackupService
        Write-output "`n####### End Debug Detail #######"
    }else{
        $BackupServiceName = "Backup Service Controller"
        Get-BackupService
    }  

Function Get-BackupProcess {

    $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"

    If ($BackupService.status -eq "Running") {
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
        if ($FunctionalProcess) {

            $Global:ProcessStateOutTXT = "Running"
            # Run Time
            $CurrentTime = Get-Date
            $Script:ProcessDuration = New-TimeSpan -Start $FunctionalProcess.StartTime -End $CurrentTime
            
            # RAM
            $CompObject = Get-WmiObject -Class WIN32_OperatingSystem     
            $SystemRAM = $CompObject.TotalVisibleMemorySize / 1mb
            $UsedRAM = (($CompObject.TotalVisibleMemorySize - $CompObject.FreePhysicalMemory)/ $CompObject.TotalVisibleMemorySize)

            # BackupFP Memory usage
            $ProcessMemoryUsage = Get-WmiObject WIN32_PROCESS | Sort-Object -Property ws -Descending | where-object {$_.Processname -eq "Backupfp.exe"} | Select-Object processname, @{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}}
    
            # CPU cores
            $Cores = (Get-WmiObject -class win32_processor -Property numberOfCores).NumberOfCores
            $LogicalCores = (Get-WmiObject -class Win32_processor -Property NumberOfLogicalProcessors).NumberOfLogicalProcessors
        
            # BackupFP CPU usage
            $SleepSeconds = 2
            $cpu1 = (get-process -name "BackupFP" | where-object {$_.path -eq $BackupFP}).cpu
            start-sleep -Seconds $SleepSeconds  
            $cpu2 = (get-process -name "BackupFP" | where-object {$_.path -eq $BackupFP}).cpu
            #$BackupFPpercentCPU = (($cpu2 - $cpu1)/($LogicalCores * $SleepSeconds)).ToString('P0')
            $BackupFPpercentCPU = (($cpu2 - $cpu1)/($LogicalCores * $SleepSeconds)).ToString()

            if ($FunctionalProcess.Responding -ne "True") {$Global:ProcessRespondingOutVal = 1;$Global:ProcessRespondingOutTxt = "False"
            }else{$Global:ProcessRespondingOutVal = 0;$Global:ProcessRespondingOutTxt = "True"}
            if ($FunctionalProcess.ProductVersion) {$Global:ClientVersionOutTxt = $FunctionalProcess.ProductVersion}
            if ($ProcessDuration) {$Global:ProcessDurationinHoursOutVal = $ProcessDuration.TotalHours | RND '' 0}
            if ($BackupFPpercentCPU) {$Global:ProcessCPUPecentOutVal = ($BackupFPpercentCPU | RND '' 2) }
            if ($ProcessMemoryUsage) {$Global:ProcessRAMUsageMBOutVal = $ProcessMemoryUsage.'Mem Usage(MB)'}

            Write-Output "`n[Backup Functional Process]"
            Write-Output "Client Version             : $Global:ClientVersionOuttxt"
            Write-Output "Process State              : $Global:ProcessStateOutVal" 
            Write-Output "Process Responding         : $Global:ProcessRespondingOutTxt"
            Write-Output "Process Start              : $($FunctionalProcess.StartTime)"
            Write-Output "Current Time               : $CurrentTime"
            Write-Output "Process Uptime (~Hours)    : $Global:ProcessDurationinHoursOutVal"
            Write-Output "Phys | Logical Cores       : $Cores|$LogicalCores"
            Write-Output "Process CPU %              : $Global:ProcessCPUPecentOutVal"
            Write-Output "System RAM GB              : $([math]::round($SystemRAM,2))"
            Write-Output "System RAM % Used          : $($UsedRAM.ToString('P0'))"
            Write-Output "Process RAM MB Used        : $Global:ProcessRAMUsageMBOutVal"


        }
    }else{
        $Global:ProcessStateOutTXT = "Not Running"
    }
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Process State" $Global:ProcessStateOutTXT
        Write-Output "-- Process Version" $Global:ClientVersionOutTxt  
        Write-Output "-- Process Responding" $Global:ProcessRespondingOutVal
        Write-Output "-- Process Responding" $Global:ProcessRespondingOutTxt
        Write-Output "-- Process Uptime (~HR)" ($Global:ProcessDurationinHoursOutVal | rnd '' 0)
        Write-Output "-- Process CPU %" $Global:ProcessCPUPecentOutVal
        ## Trigger Threshold if > 50
        Write-Output "-- Process RAM (MB)" $Global:ProcessRAMUsageMBOutVal
        ## Trigger Threshold if > 1024 MB
        Write-Output "`n####### End Threshold Values #######"
    }
} ## Get BackupFP Process Info

Get-BackupProcess

Function Get-InitError {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($debugdetail) {
        Restart-Service -Name "Backup Service Controller" -Force
        start-sleep -Seconds 10
    }

    if ($Null -eq ($FunctionalProcess)) { 
        Break
    }else{ 
        try {
            $ErrorActionPreference = 'Stop'; $initerror = & $clienttool control.initialization-error.get  | convertfrom-json
        }catch{
            Write-Warning "ERROR     : $_" 
        }
    }
    Write-output "`n[Cloud Initialization]"


        if ($initerror.code -ne "0") {
            $Global:CloudInitOutVal = $initerror.Code
            $Global:CloudInitOutTxt = $initerror.Message
            Write-Warning "`nCloud Init                 : $Global:CloudInitOutTxt`n"
        }elseif ($initerror.code -eq "0") { 
            $Global:CloudInitOutVal = $initerror.Code
            $Global:CloudInitOutTxt = "Ok"
            Write-output "Cloud Init                 : $Global:CloudInitOutTxt"
        }  

    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Cloud Connection" $Global:CloudInitOutVal 
        Write-Output "-- Cloud Connection"  $Global:CloudInitOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
} ## Get Backup Manager Initialization Errors from ClientTool.exe

Get-InitError

Function Get-ApplicationStatus {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $AppStatus = & $clienttool control.application-status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Manager]"

    if ($AppStatus) {
        $Global:AppStatusOutTxt = $AppStatus
        Write-output "Application Status         : $Global:AppStatusOutTxt"
    }
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Application Status" $Global:AppStatusOutTxt 
        Write-Output "`n####### End Threshold Values #######"
    }
} ## Get Backup Application Status from ClientTool.exe

Get-ApplicationStatus

Function Get-JobStatus {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $JobStatus = & $clienttool control.status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Job]"

    if ($JobStatus) {
        $Global:JobStatusOutTxt = $JobStatus
        Write-output "Job Status                 : $Global:JobStatusOutTxt"
    }
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Job Status" $Global:JobStatusOutTxt 
        Write-Output "`n####### End Threshold Values #######"
    }
} ## Get Backup Job Status from ClientTool.exe

Get-JobStatus

Function Get-BackupIntegrationStatus {

    $IntegratedBackupXML = "C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\MSPBackupManagerConfig.xml"     ## Ncentral XML for Integrated Backup
    
    if ((test-path -PathType Leaf -Path $IntegratedBackupXML) -eq $true) {

        [xml]$BackupAgent = get-content -path $IntegratedBackupXML 

        $IsIntegrated = $BackupAgent.MaxBackupManagerConfig.enabled
        $IsInstalled  = $BackupAgent.MaxBackupManagerConfig.installationStatus
        $CUID         = $BackupAgent.MaxBackupManagerConfig.partnerUuid
        $ProfileID    = $BackupAgent.MaxBackupManagerConfig.profileId

        Write-Output "`n[N-Central Integration Status]"
        if ($IsIntegrated -eq "False") {
            $Global:IsIntegratedOutTxt = $IsIntegrated
            $Global:IsIntegratedOutVal = 0
            Write-Output "Integrated Backup          : $IsIntegrated"
        }
        if ($IsIntegrated -eq "True") {
            $Global:IsIntegratedOutTxt = $IsIntegrated
            $Global:IsIntegratedOutVal = 1
            Write-Output "Integrated Backup          : $IsIntegrated"
            Write-Output "Backup Install Status      : $IsInstalled"
            Write-Output "Backup CUID                : $CUID"
            Write-Output "Backup ProfileId           : $ProfileID"
        }
    } else {Write-output "No prior N-central Integrated Backup settings found"}
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Integration" $Global:IsIntegratedOutTxt
        Write-Output "-- Integration" $Global:IsIntegratedOutVal
        Write-Output "`n####### End Threshold Values #######"
    }
}  ## Check N-Central for Backup Integration Status

Get-BackupIntegrationStatus

Function Get-Status ([int]$SynchThreshold) {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
    if ($Null -eq ($FunctionalProcess)) { 
        Write-Output "Backup Manager Not Running" 
    }else{ 
        Do {
            $BackupStatus = & $clienttool -machine-readable control.status.get
            $StatusValue = @("Suspended")
            if ($StatusValue -contains $BackupStatus) { 
            }else{   
                try { 
                    $ErrorActionPreference = 'Stop'
                    $Script:LSVSettings = & $clienttool control.setting.list 
                }
                catch{ # What to do with terminating errors
                    Write-Warning "ERROR     : $_" 
                }
            }
        }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 
        write-output "$lsvsettings"
        write-output $retrycounter $BackupStatus

    $Script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $DocumentsEnabled                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//AntiCryptoEnabled")."#text"
    $Device                                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
    $MachineName                                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text"
    $PartnerName                                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text" 
    $TimeStamp                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $TimeZone                                   = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeZone")."#text"
    $LocalSpeedVaultEnabled                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultEnabled")."#text"
    $BackupServerSynchronizationStatus          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//BackupServerSynchronizationStatus")."#text"
    $LocalSpeedVaultSynchronizationStatus       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultSynchronizationStatus")."#text"
    $SelectedSize                               = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginTotal-LastCompletedSessionSelectedSize")."#text"
    $UsedStorage                                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//UsedStorage")."#text"
    if (($LocalSpeedVaultEnabled -eq 1) -and ($LSVSettings)) {
        $LSVPath                                = ($LSVSettings | Where-Object { $_ -like "LocalSpeedVaultLocation *"}).replace("LocalSpeedVaultLocation ","")
        $LSVUser                                = ($LSVSettings | Where-Object { $_ -like "LocalSpeedVaultUser *"}).replace("LocalSpeedVaultUser     ","")
    }

    $Global:CoveDeviceNameOutTxt = $Device
    $Global:MachineNameOutTxt = $MachineName    
    $Global:PartnerNameOutTxt = $PartnerName    
    $Global:TotalSelectedGBOutTxt = ($SelectedSize | RND GB 2)
    $Global:TotalUsedGBOutTxt = ($UsedStorage | RND GB 2)
    $Global:LSVEnabledOutTxt = $LocalSpeedVaultEnabled.replace("1","True").replace("0","False")
    $Global:LSVEnabledOutVal = $LocalSpeedVaultEnabled
    $Global:CloudSyncStatusOutTxt = $BackupServerSynchronizationStatus
    if ($LocalSpeedVaultEnabled -eq 1) { 
         $Global:LSVSyncStatusOutTxt = $LocalSpeedVaultSynchronizationStatus
    }elseif ($LocalSpeedVaultEnabled -eq 0) {
         $Global:LSVSyncStatusOutTxt = "Not Enable"
    }
    Write-Output "`n[Settings]"
    Write-Output "Device                  : $Global:CoveDeviceNameOutTxt"
    Write-Output "Machine                 : $Global:MachineNameOutTxt"
    Write-Output "Partner Name            : $Global:PartnerNameOutTxt"
    Write-Output "TimeStamp (UTC)         : $(Convert-UnixTimeToDateTime $timestamp)"
    Write-Output "TimeZone                : $timezone"

    Write-Output "`n[LocalSpeedVault Status]"
    Write-Output "LSV Enabled             : $Global:LSVEnabledOutTxt"
    Write-Output "LSV Sync Status         : $Global:LSVSyncStatusOutTxt"
    Write-Output "Cloud Sync Status       : $Global:CloudSyncStatusOutTxt"
    Write-Output "Selected Size (GB)      : $Global:TotalSelectedGBOutTxt"        
    Write-Output "Used Storage (GB)       : $Global:TotalUsedGBOutTxt"
    Write-Output "Last LSV Path           : $LSVpath"
    Write-Output "Last LSV User           : $LSVUser"

    # Fail if LSV Enabled = True & ( LSV Sync = Failed or Cloud Sync = Failed )
    if (($LocalSpeedVaultEnabled -eq 1) -and (($BackupServerSynchronizationStatus -eq "Failed") -or ($LocalSpeedVaultSynchronizationStatus -eq "Failed"))) {Write-Warning "LSV Failed"}
            
    elseif (($LocalSpeedVaultEnabled -eq 1) -and ($BackupServerSynchronizationStatus -ne "synchronized")) { 
        if ( ($BackupServerSynchronizationStatus.replace("%","")/1 -lt $SynchThreshold )){Write-Warning "Cloud sync is below $SynchThreshold%"}
    } ## Warn if Sync % is below threshold
        
    elseif (($LocalSpeedVaultEnabled -eq 1) -and ($LocalSpeedVaultSynchronizationStatus -ne "synchronized")) { 
        if ( ($LocalSpeedVaultSynchronizationStatus.replace("%",""))/1 -lt $SynchThreshold ){Write-Warning "LSV sync is below $SynchThreshold%"}
    }  ## Warn if Sync % is below threshold

    # Warn if LSV Enabled = False & LSV Path != ""
    
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Device" $Global:CoveDeviceNameOutTxt
        Write-Output "-- Machine" $Global:MachineNameOutTxt
        Write-Output "-- Partner" $Global:PartnerNameOutTxt
        Write-Output "-- LSV Enabled" $Global:LSVEnabledOutTxt
        Write-Output "-- LSV Sync" $Global:LSVSyncStatusOutTxt
        Write-Output "-- Cloud Sync" $Global:CloudSyncStatusOutTxt
        Write-Output "-- Selected GB" $Global:TotalSelectedGBOutTxt
        Write-Output "-- Used GB" $Global:TotalUsedGBOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
}
    

} ## Get Status

Get-Status

Function Get-Datasources {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { 
        "Backup Manager Not Running" 
    }
    else { 
        try { # Command(s) to try
            $ErrorActionPreference = 'Stop'
            $script:Datasources = & $clienttool -machine-readable control.selection.list | ConvertFrom-String | Select-Object -skip 1 -property p1,p2 -unique | ForEach-Object {If ($_.P2 -eq "Inclusive") {Write-Output $_.P1}} 
        }
        catch{ # What to do with terminating errors
        }
    }
    $Global:DataSourcesOutTxt = $Datasources -join ", "
    Write-Output "`n[Configured Datasources]"
    Write-Output "Data Sources            : $Global:DataSourcesOutTxt"

    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- DataSources" $Global:DataSourcesOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
        

} ## Get Last Error per Active Data Source from ClientTool.exe

Get-Datasources

Function Get-BackupProductOsType {

    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $DateTime = $epoch.AddSeconds($inputUnixTime)
        return "{0:yyyy/MM/dd} {0:HH:mm:ss}" -f $DateTime
        }else{ return ""}
    }  ## Convert epoch time to date time 

    $FunctionalProcess = get-process "BackupFP" -ea SilentlyContinue
    $StatusReportxml   = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" ## StatusReport.XML to Parse
    $Configinipath     = "C:\Program Files\Backup Manager\Config.ini"  
    
    $ProductStatus = @{
        "$null" = '0'
        'Professional Workstation'  = '1'
        'Professional Server'       = '2'
        'Documents Workstation'     = '3'
        'Documents Server'          = '4'
    }

    if (($null -eq $FunctionalProcess) -and ((Test-Path $Configinipath -PathType leaf) -eq $false)) {
        Write-Warning "Active Backup Manager Installation Not Found"
    }else{
        if ((Test-Path $StatusReportxml -PathType leaf) -eq $false) {
            Write-Warning "No Backup Status History Found"
        }else{
            $OsVersion                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text"
            $Device                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
            $PartnerName                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
            $TimeStamp                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
            $PluginTotalLastSessionTimestamp  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginTotal-LastSessionTimestamp")."#text"
            
            $StoragePath  = "C:\ProgramData\MXB\Backup Manager\storage\$device*\config\$device.info"
            if ((Test-Path $StoragePath -PathType leaf) -eq $true) {
            $ProductInfo  = Get-Content -Raw -Path "c:\programdata\mxb\backup manager\storage\$device*\config\$device.info" | ConvertFrom-Json 
            }
                       
            if ($ProductInfo.ProductName -eq "Documents") { $ProductType = "Documents" }
            if (($ProductInfo.ProductName -ne "Documents") -and ($null -ne $ProductInfo.ProductName)) { $ProductType = "Professional" }
            if ($OsVersion -like "*Server*") { $OsType = "Server" }
            if ($OsVersion -notlike "*Server*") { $OsType = "Workstation" }
            if (($ProductType) -and ($OsType)) { 
                $ProductOsType = "$ProductType $OsType"
                $Global:ProductOSTypeOutTxt = $ProductOsType
            }
            if ($ProductOsType -ne "Documents Server") {
                Write-Output "`n[Backup Product/OS Type]"
                Write-Output "OS                      : $OsVersion"
                Write-Output "Customer                : $PartnerName"
                Write-Output "Device                  : $Device"
                Write-Output "TimeStamp (UTC)         : $(Convert-UnixTimeToDateTime $timestamp)"
                Write-Output "LastSession (UTC)       : $(Convert-UnixTimeToDateTime $PluginTotalLastSessionTimestamp)"
                Write-Output "Backup Product\ OS      : $ProductOSType"
                if ($ProductOSType -eq "Documents Server") {
                    Write-Warning "Configuration is Not Supported"
                }
            }
        }
    }
    $ProductStatusCode = $ProductStatus["$ProductOSType"]
    $ProductStatusCode
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Product OS\Type" $Global:ProductOSTypeOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
}  ## Get Installed Backup Product / Detected OS Type

Get-BackupProductOsType 