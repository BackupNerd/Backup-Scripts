<# ----- About: ----
    # N-able | Cove Data Protection | Monitor Application
    # Revision v07 - 2022-05-31
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
    #  Get Cove service status
    #  Get Cove installation/details
    #  Get Cove installed version
    #  Get Cove integration status   
    #  Get Cove process state/uptime    
    #  Get Cove process CPU/RAM usage      
    #  Get Cove LSV/Cloud sync status
    #  Get Cove selected data sources        
    #  Get Cove selected/used storage 
    #  Get Cove license type 
 # -----------------------------------------------------------#>

<#
[CmdletBinding()]
    Param (
               [Parameter(Mandatory=$False)] [switch]$ThresholdOutput = $true      ## Set True to Output Threshold Values at end of script
    )

#Clear-Host
#>

#Requires -Version 5.1 -RunAsAdministrator
[switch]$ThresholdOutput = $true 
#$ConsoleTitle = "Cove Data Protection - Application Monitor"
#$host.UI.RawUI.WindowTitle = $ConsoleTitle
#region Functions

Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return "{0:yyyy/MM/dd} {0:HH:mm:ss}" -f $epoch
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

$Script:BackupServiceOutVal = -1
$Script:BackupServiceOutTxt = "Not Installed"
$Script:ProcessStateOutTxt = "Undefined"
$Script:ClientVerOutTxt = "Undefined"

$Script:ProcessRespondingOutVal = 0
$Script:ProcessRespondingOutTxt = "Undefined"
$Script:ProcessDurationinHoursOutVal = 0
$Script:ProcessCPUPecentOutVal = 0
$Script:ProcessRAMUsageMBOutVal = 0
$Script:CloudInitOutVal = 0
$Script:CloudInitOutTxt = "Undefined"
$Script:AppStatusOutTxt = "Undefined"
$Script:JobStatusOutTxt = "Undefined"
$Script:IsIntegratedOutVal = 0
$Script:IsIntegratedOutTxt = "Undefined"
$Script:CoveDeviceNameOutTxt = "Undefined"
$Script:MachineNameOutTxt = "Undefined"
$Script:CustomerNameOutTxt = "Undefined"
$Script:ProfileNameOutTxt = "Undefined"
$Script:OsVersionOutTxt = "Undefined"
$Script:LSVEnabledOutVal = 0
$Script:LSVEnabledOutTxt = "Undefined"
$Script:LSVSyncStatusOutTxt = "Undefined"
$Script:CloudSyncStatusOutTxt = "Undefined"
$Script:TotalSelectedGBOutTxt = "Undefined"
$Script:TotalUsedGBOutTxt = "Undefined"
$Script:BRSizeOutVal = 0
$Script:DataSourcesOutTxt = "Undefined"
$Script:ProductOSTypeOutTxt = "Undefined"






Function Get-BackupState {
 
    $Script:BackupService = get-service -name $BackupServiceName -ea SilentlyContinue
    Write-output "`n[Backup Service Controller]"
    if ($BackupService.status -eq "Running"){
        $Script:BackupServiceOutVal = 1
        $Script:BackupServiceOutTxt = "$($BackupService.status)" 
        Write-output  "Service Status             : $BackupServiceOutTxt"
    }elseif (($BackupService.status -ne "Running") -and ($null -ne $BackupService.status )){
        $Script:BackupServiceOutVal = 0
        $Script:BackupServiceOutTxt = "$($BackupService.status)"
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }elseif ($null -eq $BackupService.status ){
        $Script:BackupServiceOutVal = -1
        $Script:BackupServiceOutTxt = "Not Installed"
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }

} ## Get Backup Service \ Process Status

$BackupServiceName = "Backup Service Controller"
Get-BackupState

Function Get-BackupProcess {

    $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"

    If ($BackupService.status -eq "Running") {
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
        if ($FunctionalProcess) {

            $Script:ProcessStateOutTXT = "Running"
            # Run Time
            $CurrentTime = "{0:yyyy/MM/dd} {0:HH:mm:ss}" -f $(Get-Date)
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

            if ($FunctionalProcess.Responding -ne "True") {$Script:ProcessRespondingOutVal = 0;$Script:ProcessRespondingOutTxt = "False"
            }else{$Script:ProcessRespondingOutVal = 1;$Script:ProcessRespondingOutTxt = "True"}
            if ($FunctionalProcess.ProductVersion) {$Script:ClientVerOutTxt = $FunctionalProcess.ProductVersion}
            if ($ProcessDuration) {$Script:ProcessDurationinHoursOutVal = $ProcessDuration.TotalHours | RND '' 0}
            if ($BackupFPpercentCPU) {$Script:ProcessCPUPecentOutVal = ($BackupFPpercentCPU | RND '' 2) }
            if ($ProcessMemoryUsage) {$Script:ProcessRAMUsageMBOutVal = $ProcessMemoryUsage.'Mem Usage(MB)'}
            $FunctionalProcessStart = "{0:yyyy/MM/dd} {0:HH:mm:ss}" -f $FunctionalProcess.StartTime

            Write-Output "`n[Backup Functional Process]"
            Write-Output "Client Version             : $Script:ClientVerOutTxt"
            Write-Output "Process State              : $Script:ProcessStateOutTxt" 
            Write-Output "Process Responding         : $Script:ProcessRespondingOutTxt"
            Write-Output "Process Start              : $FunctionalProcessStart"
            Write-Output "Current Time               : $CurrentTime"
            Write-Output "Process Uptime (~Hours)    : $Script:ProcessDurationinHoursOutVal"
            Write-Output "Phys | Logical Cores       : $Cores|$LogicalCores"
            Write-Output "Process CPU %              : $Script:ProcessCPUPecentOutVal"
            Write-Output "System RAM GB              : $([math]::round($SystemRAM,2))"
            Write-Output "System RAM % Used          : $($UsedRAM.ToString('P0'))"
            Write-Output "Process RAM MB Used        : $Script:ProcessRAMUsageMBOutVal"
        }
    }else{
        $Script:ProcessStateOutTXT = "Not Running"
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
        #Break
    }else{ 
        try {
            $ErrorActionPreference = 'Stop'; $initerror = & $clienttool control.initialization-error.get  | convertfrom-json

            Write-output "`n[Cloud Initialization]"

            if ($initerror.code -ne "0") {
                $Script:CloudInitOutVal = $initerror.Code
                $Script:CloudInitOutTxt = $initerror.Message
                Write-Warning "`nCloud Init                 : $Script:CloudInitOutTxt`n"
            }elseif ($initerror.code -eq "0") { 
                $Script:CloudInitOutVal = $initerror.Code
                $Script:CloudInitOutTxt = "Ok"
                Write-output "Cloud Init                 : $Script:CloudInitOutTxt"
            }  
        }catch{
            Write-Warning "ERROR     : $_" 
        }
    }

} ## Get Backup Manager Initialization Errors from ClientTool.exe

Get-InitError

Function Get-ApplicationStatus {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $AppStatus = & $clienttool control.application-status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Manager]"

    if ($AppStatus) {
        $Script:AppStatusOutTxt = $AppStatus
        Write-output "Application Status         : $Script:AppStatusOutTxt"
    }

} ## Get Backup Application Status from ClientTool.exe

Get-ApplicationStatus

Function Get-JobStatus {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $JobStatus = & $clienttool control.status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Job]"

    if ($JobStatus) {
        $Script:JobStatusOutTxt = $JobStatus
        Write-output "Job Status                 : $Script:JobStatusOutTxt"
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
            $Script:IsIntegratedOutTxt = $IsIntegrated
            $Script:IsIntegratedOutVal = 0
            Write-Output "Integrated Backup          : $IsIntegrated"
        }
        if ($IsIntegrated -eq "True") {
            $Script:IsIntegratedOutTxt = $IsIntegrated
            $Script:IsIntegratedOutVal = 1
            Write-Output "Integrated Backup          : $IsIntegrated"
            Write-Output "Backup Install Status      : $IsInstalled"
            Write-Output "Backup CUID                : $CUID"
            Write-Output "Backup ProfileId           : $ProfileID"
        }
    } else {Write-output "No prior N-central Integrated Backup settings found"}

}  ## Check N-Central for Backup Integration Status

Get-BackupIntegrationStatus

Function Get-Status ([int]$SynchThreshold) {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
    Do {
        $BackupStatus = & $clienttool -machine-readable control.status.get
        $StatusValue = @("Suspended")
        if ($StatusValue -contains $BackupStatus) { 
        }else{   
            try { 
                $ErrorActionPreference = 'SilentlyContinue'
                $Script:LSVSettings = & $clienttool control.setting.list 
            }
            catch{ # What to do with terminating errors
                Write-Warning "ERROR     : $_" 
            }
        }
    }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 
    #write-output "$lsvsettings"
    #write-output $retrycounter $BackupStatus

    $Script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $DocumentsEnabled                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//AntiCryptoEnabled")."#text"
    $Script:CoveDeviceNameOutTxt                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
    $Script:MachineNameOutTxt                   = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text"
    $Script:CustomerNameOutTxt                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
    Script:ProfileNameOutTxt                    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileName")."#text"
    if ($null -eq $Script:ProfileNameOutTxt ) { $Script:ProfileNameOutTxt = "Undefined"} 
    $ProfileId                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileId")."#text"
    $ProfileVersion                             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileVersion")."#text"
    $Script:OsVersionOutTxt                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text" 
    $Script:TimeStampOutVal                     = Convert-UnixTimeToDateTime ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
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

    #$Script:CoveDeviceNameOutTxt = $Device
    #$Script:MachineNameOutTxt = $MachineName    
    #$Script:PartnerNameOutTxt = $PartnerName
    $Script:ProfileIdOutVal = $ProfileId    
    #$Script:ProfileNameOutTxt = $ProfileName        
    $Script:TotalSelectedGBOutTxt = ($SelectedSize | RND GB 2)
    $Script:TotalUsedGBOutTxt = ($UsedStorage | RND GB 2)
    $Script:LSVEnabledOutTxt = $LocalSpeedVaultEnabled.replace("1","True").replace("0","False")
    $Script:LSVEnabledOutVal = $LocalSpeedVaultEnabled
    $Script:CloudSyncStatusOutTxt = $BackupServerSynchronizationStatus
    if ($LocalSpeedVaultEnabled -eq 1) { 
         $Script:LSVSyncStatusOutTxt = $LocalSpeedVaultSynchronizationStatus
    }elseif ($LocalSpeedVaultEnabled -eq 0) {
         $Script:LSVSyncStatusOutTxt = "Disabled"
    }
    Write-Output "`n[Settings]"
    Write-Output "Device                     : $Script:CoveDeviceNameOutTxt"
    Write-Output "Machine                    : $Script:MachineNameOutTxt"
    Write-Output "Customer Name              : $Script:CustomerNameOutTxt"
    Write-Output "Profile Name               : $Script:ProfileNameOutTxt"
    Write-Output "ProfileId                  : $ProfileId"
    Write-Output "ProfileVersion             : $ProfileVersion"
    Write-Output "TimeStamp (UTC)            : $Script:TimeStampOutVal"
    Write-Output "TimeZone                   : $timezone"

    Write-Output "`n[LocalSpeedVault Status]"
    Write-Output "LSV Enabled                : $Script:LSVEnabledOutTxt"
    Write-Output "LSV Sync Status            : $Script:LSVSyncStatusOutTxt"
    Write-Output "Cloud Sync Status          : $Script:CloudSyncStatusOutTxt"
    Write-Output "Selected Size (GB)         : $Script:TotalSelectedGBOutTxt"        
    Write-Output "Used Storage (GB)          : $Script:TotalUsedGBOutTxt"
    Write-Output "Last LSV Path              : $LSVpath"
    Write-Output "Last LSV User              : $LSVUser"

    # Fail if LSV Enabled = True & ( LSV Sync = Failed or Cloud Sync = Failed )
    if (($LocalSpeedVaultEnabled -eq 1) -and (($BackupServerSynchronizationStatus -eq "Failed") -or ($LocalSpeedVaultSynchronizationStatus -eq "Failed"))) {Write-Warning "LSV Failed"}
            
    elseif (($LocalSpeedVaultEnabled -eq 1) -and ($BackupServerSynchronizationStatus -ne "synchronized")) { 
        if ( ($BackupServerSynchronizationStatus.replace("%","")/1 -lt $SynchThreshold )){Write-Warning "Cloud sync is below $SynchThreshold%"}
    } ## Warn if Sync % is below threshold
        
    elseif (($LocalSpeedVaultEnabled -eq 1) -and ($LocalSpeedVaultSynchronizationStatus -ne "synchronized")) { 
        if ( ($LocalSpeedVaultSynchronizationStatus.replace("%",""))/1 -lt $SynchThreshold ){Write-Warning "LSV sync is below $SynchThreshold%"}
    }  ## Warn if Sync % is below threshold

    # Warn if LSV Enabled = False & LSV Path != ""
    
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
            $Script:DataSourcesOutTxt = $Datasources -join ", "
            Write-Output "`n[Configured Datasources]"
            Write-Output "Data Sources               : $Script:DataSourcesOutTxt" 
        }
        catch{ # What to do with terminating errors
        }
    }

} ## Get Last Error per Active Data Source from ClientTool.exe

Get-Datasources

Function Get-BRSize {

    $BRPath = "c:\programdata\mxb\backup manager\storage\*\br\br.db"

    if ((test-path -PathType Leaf -Path $BRPath) -eq $true) {

        $Script:BRSizeOutVal = (get-item $BRpath | Sort-Object Lastwritetime)[-1].length

        Write-Output "Local Backup Register Size : $($Script:BRSizeOutVal | RND GB 3)"
    }
}

Get-BRSize

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
            $Script:OsVersionOutTxt           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text"
            $Script:CoveDeviceNameOutTxt      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
            $Script:PartnerNameOutTxt         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
            $PluginTotalLastSessionTimestamp  = Convert-UnixTimeToDateTime ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginTotal-LastSessionTimestamp")."#text"
            
            $StoragePath  = "C:\ProgramData\MXB\Backup Manager\storage\$Script:CoveDeviceNameOutTxt*\config\$Script:CoveDeviceNameOutTxt.info"
            if ((Test-Path $StoragePath -PathType leaf) -eq $true) {
            $ProductInfo  = Get-Content -Raw -Path "c:\programdata\mxb\backup manager\storage\$Script:CoveDeviceNameOutTxt*\config\$Script:CoveDeviceNameOutTxt.info" | ConvertFrom-Json 
            }
                       
            if ($ProductInfo.ProductName -eq "Documents") { $ProductType = "Documents" }
            if (($ProductInfo.ProductName -ne "Documents") -and ($null -ne $ProductInfo.ProductName)) { $ProductType = "Professional" }
            if ($Script:OsVersionOutTxt -like "*Server*") { $OsType = "Server" }
            if ($Script:OsVersionOutTxt -notlike "*Server*") { $OsType = "Workstation" }
            if (($ProductType) -and ($OsType)) { 
                $Script:ProductOSTypeOutTxt = "$ProductType $OsType"
                #$Script:ProductOSTypeOutTxt = $ProductOsType
            }
            if ($Script:ProductOSTypeOutTxt -ne "Documents Server") {
                Write-Output "`n[Backup Product/OS Type]"
                Write-Output "OS                         : $Script:OsVersionOutTxt"
                Write-Output "Customer                   : $Script:PartnerNameOutTxt"
                Write-Output "Device                     : $Script:CoveDeviceNameOutTxt"
                Write-Output "TimeStamp (UTC)            : $Script:TimeStampOutVal"
                Write-Output "LastSession (UTC)          : $PluginTotalLastSessionTimestamp"
                Write-Output "Backup Product\ OS         : $Script:ProductOSTypeOutTxt"
                if ($Script:ProductOSTypeOutTxt -eq "Documents Server") {
                    Write-Warning "Configuration is Not Supported"
                }
            }
        }
    }
    $ProductStatusCode = $ProductStatus["$Script:ProductOSTypeOutTxt"]
    $ProductStatusCode

}  ## Get Installed Backup Product / Detected OS Type

Get-BackupProductOsType

if ($ThresholdOutput) {
    Write-Output "`n####### Start Threshold Output Values #######"
    Write-Output " Backup service value     | $Script:BackupServiceOutVal"
    Write-Output " Backup service state     | $Script:BackupServiceOutTxt"
    Write-Output " Device name              | $Script:CoveDeviceNameOutTxt"
    Write-Output " Machine name             | $Script:MachineNameOutTxt"
    Write-Output " Customer name            | $Script:CustomerNameOutTxt"

    Write-Output " Backup client ver        | $Script:ClientVerOutTxt"
    Write-Output " OS Version               | $Script:OsVersionOutTxt"

    Write-Output " Backup process state     | $Script:ProcessStateOutTxt"
    Write-Output " Process Responding Value | $Script:ProcessRespondingOutVal"
    Write-Output " Process Responding State | $Script:ProcessRespondingOutTxt"
    Write-Output " Process Uptime (HRS)     | $Script:ProcessDurationinHoursOutVal"
    Write-Output " Process CPU %            | $Script:ProcessCPUPecentOutVal"
    Write-Output " Process RAM (MB)         | $Script:ProcessRAMUsageMBOutVal"

    Write-Output " Cloud Connection         | $Script:CloudInitOutVal"
    Write-Output " Cloud Connection         | $Script:CloudInitOutTxt"
    Write-Output " Application Status       | $Script:AppStatusOutTxt"
    Write-Output " Job Status               | $Script:JobStatusOutTxt"

    Write-Output " Integration State        | $Script:IsIntegratedOutTxt"
    Write-Output " Integration Value        | $Script:IsIntegratedOutVal"

    Write-Output " Profile                  | $Script:ProfileNameOutTxt"

    Write-Output " LSV Enabled              | $Script:LSVEnabledOutTxt"
    Write-Output " LSV Sync                 | $Script:LSVSyncStatusOutTxt"
    Write-Output " Cloud Sync               | $Script:CloudSyncStatusOutTxt"
    Write-Output " Selected GB              | $Script:TotalSelectedGBOutTxt"
    Write-Output " Used GB                  | $Script:TotalUsedGBOutTxt"
    
    Write-Output " Backup Reg Size (GB)     | $($Script:BRSizeOutVal/1GB | rnd '' 3) "
    
    Write-Output " DataSources              | $Script:DataSourcesOutTxt"
    Write-Output " Product OS\Type          | $Script:ProductOSTypeOutTxt"
    Write-Output "`n####### End Threshold Output Values #######"
}

Function Reset-Full {
    Stop-service -Name 'Backup Service Controller' -Force
    Copy-Item 'C:\config.ini.full' 'C:\Program Files\Backup Manager\config.ini'
    Start-Service -Name 'Backup Service Controller'
} ## debug purposes only

Function Reset-Docs {
    Stop-service -Name 'Backup Service Controller' -Force
    Copy-Item 'C:\config.ini.docs' 'C:\Program Files\Backup Manager\config.ini'
    Start-Service -Name 'Backup Service Controller'
}  ## debug purposes only