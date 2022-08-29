[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [switch]$debugdetail   
    )        
Clear-host

#region ----- Environment, Variables, Names and Paths ----

$clientool = "C:\Program Files\Backup Manager\ClientTool.exe"
$config = "C:\Program Files\Backup Manager\config.ini" 

#endregion ----- Environment, Variables, Names and Paths ----

#region Functions
Function Hash-Value ($value) {
    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding
    $hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($value)))
    $hash
}

Function Get-TimeStamp {
    return "[{0:yyyy/MM/dd} {0:HH:mm:ss}]" -f (Get-Date)
}

Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $DateTime = $epoch.AddSeconds($inputUnixTime)
    return $DateTime
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

}

Function Get-BackupIntegrationStatus {

    $IntegratedBackupXML = "C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\MSPBackupManagerConfig.xml" ## Ncentral XML for Integrated Backup
    
    if ((test-path -PathType Leaf -Path $IntegratedBackupXML) -eq $true) {

        [xml]$BackupAgent = get-content -path $IntegratedBackupXML 

        $IsIntegrated = $BackupAgent.MaxBackupManagerConfig.enabled
        $IsInstalled  = $BackupAgent.MaxBackupManagerConfig.installationStatus
        $CUID         = $BackupAgent.MaxBackupManagerConfig.partnerUuid
        $ProfileID    = $BackupAgent.MaxBackupManagerConfig.profileId

        if ($IsIntegrated -eq "False") {
            Write-Output "`n[N-Central Integration Status]"
            Write-Output "Integrated Backup          : $IsIntegrated"
        }
        if ($IsIntegrated -eq "True") {
            Write-Output "`n[Integration Status]"
            Write-Output "Integrated Backup          : $IsIntegrated"
            Write-Output "Backup Install Status      : $IsInstalled"
            Write-Output "Backup CUID                : $CUID"
            Write-Output "Backup ProfileId           : $ProfileID"
        }
    } else {Write-output "No prior N-central Integrated Backup settings found"}
}

Function Get-BackupService {
    $Script:BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
    if ($BackupService.status -eq "Running"){ 
        Write-output "`n[Backup Service Controller]"
        Write-output "Service Status             : $($BackupService.status)"
    }elseif ($BackupService.status -ne "Running"){
        Write-warning "`n[Backup Service Controller]"
        Write-output "Service Status             : $($BackupService.status)"
    }
    
} ## Get Backup Service Status

Function Get-BackupProcess {
    $Script:FunctionalProcess = get-process "BackupFP" -ea SilentlyContinue
    if ($FunctionalProcess) {
        $CurrentTime = Get-Date
        $ProcessDuration = New-TimeSpan -Start $FunctionalProcess.StartTime -End $CurrentTime
        
        # RAM
        $CompObject = Get-WmiObject -Class WIN32_OperatingSystem     
        $SystemRAM = $CompObject.TotalVisibleMemorySize / 1mb
        $UsedRAM = (($CompObject.TotalVisibleMemorySize - $CompObject.FreePhysicalMemory)/ $CompObject.TotalVisibleMemorySize)

        # CPU cores
        $Cores = (Get-WmiObject -class win32_processor -Property numberOfCores).numberOfCores
        $LogicalCores = (Get-WmiObject -class Win32_processor -Property NumberOfLogicalProcessors).NumberOfLogicalProcessors
     
        # BackupFP CPU usage
        $SleepSeconds = 2
        $cpu1 = (get-process -name Backupfp).cpu
        start-sleep -Seconds $SleepSeconds  
        $cpu2 = (get-process -name Backupfp).cpu
        $BackupFPpercentCPU = (($cpu2 - $cpu1)/($LogicalCores*$sleepseconds)).ToString('P0')

        Write-Output "`n[Environment]"
        Write-Output "Current Time               : $CurrentTime"
        Write-Output "System RAM GB              : $([math]::round($SystemRAM,2))"
        Write-Output "System RAM % Used          : $($UsedRAM.ToString('P0'))"
        Write-Output "Physical Cores             : $Cores"
        Write-Output "Logical Cores              : $LogicalCores"

        Write-Output "`n[Backup Functional Process]"
        Write-Output "Process Version            : $($FunctionalProcess.ProductVersion)"
        Write-Output "Process Responding         : $($FunctionalProcess.Responding)"
        Write-Output "Process Start              : $($FunctionalProcess.StartTime)"
        Write-Output "Process Uptime             : $ProcessDuration"
        Write-Output "Process CPU %              : $BackupFPpercentCPU"
        #Write-Output "Physical Cores             : $Cores"
        #Write-Output "Logical Cores              : $LogicalCores"
      
        # BackupFP Memory usage
        $ProcessMemoryUsage = Get-WmiObject WIN32_PROCESS | Sort-Object -Property ws -Descending | where-object {$_.Processname -eq "Backupfp.exe"} | Select-Object processname, @{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}}
        write-output "Process RAM MB Used        : $($ProcessMemoryUsage.'Mem Usage(MB)')"

    }
} ## Get BackupFP Process Info

Function Get-SystemInfo {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; [array]$BackupSystemInfo = & $clientool system-info.get | convertfrom-json }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[System Info]"

    if ($BackupSystemInfo) {

        Write-output "RAM                        : $($BackupSystemInfo.Memory)"
        Write-output "OS                         : $($BackupSystemInfo.OsVersion)"
        Write-output "CPU                        : $($BackupSystemInfo.processor)"

    }

} ## Get Backup Job Status from ClientTool.exe

Function Get-InitError {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $initerror = & $clientool control.initialization-error.get  | convertfrom-json}catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Cloud Initilization]"

    if ($initerror) {

        if ($initerror.code -gt 0) {
            Write-Warning "Cloud Init                 : $($initerror.Message)`n"
        }
        else{ 
            Write-output "`Cloud Init                 : Ok"}

    }
} ## Get Backup Manager Initialization Errors from ClientTool.exe

Function Get-ApplicationStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $AppStatus = & $clientool control.application-status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Manager]"

    if ($AppStatus) {

        Write-output "Application Status         : $appstatus"
    }
} ## Get Backup Application Status from ClientTool.exe

Function Get-BackupFPVisa {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupFPVisa = & $clientool in-agent-authentication-token.get -config-path "$config" | convertfrom-json }catch{ Write-Warning "ERROR     : $_" }} 
    Write-output "`n[Backup Manager]"
    
    if ($BackupFPVisa) {
    
        Write-output "Client Visa                : $($BackupFPVisa.InAgentAuthenticationToken)"
    }
    
} ## Get BackupFP Authentication Visa from ClientTool.exe



Function Get-JobStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $JobStatus = & $clientool control.status.get }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Job]"

    if ($JobStatus) {

        Write-output "Job Status                 : $jobStatus"
    }

} ## Get Backup Job Status from ClientTool.exe

Function Get-VssStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $VssStatus = & $clientool vss.check }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[VSS Status]"
    #($VssStatus -split '\r?\n').Trim()
    ($VssStatus | where-object {($_ -match 'error') -or ($_ -notmatch 'ok')  })
} ## Check VSS Status from ClientTool.exe

Function Get-StorageStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop' ; $StorageStatus = & $clientool storage.test }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Storage Status]"
    #($StorageStatus -split '\r?\n').Trim()
    #($StorageStatus | where-object {($_ -match '<*>') -or ($_ -match 'node') })
    if ($StorageStatus.length -eq 20)  {
        #$StorageStatus.length
        ($StorageStatus | where-object {($_ -match '<*>') -or ($_ -match 'node') -or ($_ -notmatch '... ok') })
    }else{
       $($StorageStatus | where-object {($_ -match '<*>') -or ($_ -match 'node') -or ($_ -notmatch '... ok') })
    }
} ## Get Backup Storage Status from ClientTool.exe

Function Get-BackupSettings {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSettings = & $clientool control.setting.list }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Settings]"

    if ($BackupSettings) {

        $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
        $Hash = "BackupSettings.hash"
        $log = "BackupSettings.log"

        $prior = test-path -PathType Leaf -Path $BackupMonitoringPath\$hash

        if ($prior -eq $false) {
            mkdir -Path $BackupMonitoringPath -ea SilentlyContinue
            new-item -path $BackupMonitoringPath\$hash -ea SilentlyContinue
            
            Write-output "`n$(get-timestamp)`nNo Prior Values" | out-file -FilePath $BackupMonitoringPath\$log -Append
            Write-output "No Prior Hash" | out-file -FilePath $BackupMonitoringPath\$hash
            
        }
        $BackupSettingsHash = @{}
        $BackupSettingsHash.prior = Get-Content -Path $BackupMonitoringPath\$hash
        $BackupSettingsHash.current = Hash-Value $BackupSettings

        if ($BackupSettingsHash.prior -eq $BackupSettingsHash.current ) {
            if ($debugdetail) {
                Write-Output "Prior Hash                  : $($BackupSettingsHash.prior)"
                Write-Output "Current Hash                : $($BackupSettingsHash.current)"
            }
            Write-Output "Values Match"
            #$BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            #$BackupSettings | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Output $BackupSettings
        }elseif ($BackupSettingsHash.prior -ne $BackupSettingsHash.current ){
            Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
            Write-Output "Prior Hash                  : $($BackupSettingsHash.prior)"
            Write-Output "Current Hash                : $($BackupSettingsHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
            $BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$hash
            #$BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            $BackupSettings | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        }
    }
} ## Get Backup Settings from ClientTool.exe

Function Get-BackupSelections {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSelections = & $clientool control.selection.list }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Selections]"

    if ($BackupSelections) {

        $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
        $Hash = "BackupSelections.hash"
        $log = "BackupSelections.log"

        $prior = test-path -PathType Leaf -Path $BackupMonitoringPath\$hash

        if ($prior -eq $false) {
            mkdir -Path $BackupMonitoringPath -ea SilentlyContinue
            new-item -path $BackupMonitoringPath\$hash -ea SilentlyContinue
            
            Write-output "`n$(get-timestamp)`nNo Prior Value" | out-file -FilePath $BackupMonitoringPath\$log -Append
            Write-output "No Prior Hash" | out-file -FilePath $BackupMonitoringPath\$hash
            
        }
        $BackupSelectionsHash = @{}
        $BackupSelectionsHash.prior = Get-Content -Path $BackupMonitoringPath\$hash
        $BackupSelectionsHash.current = Hash-Value $BackupSelections

        if ($BackupSelectionsHash.prior -eq $BackupSelectionsHash.current ) {
            if ($debugdetail) {
                Write-Output "Prior Hash                  : $($BackupSelectionsHash.prior)"
                Write-Output "Current Hash                : $($BackupSelectionsHash.current)"
            }
            Write-Output "Values Match"
            #$BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            #$BackupSelections | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Output $BackupSelections
        }elseif ($BackupSelectionsHash.prior -ne $BackupSelectionsHash.current ){
            Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
            Write-Output "Prior Hash                  : $($BackupSelectionsHash.prior)"
            Write-Output "Current Hash                : $($BackupSelectionsHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
            $BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$hash
            #$BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            $BackupSelections | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        }
    }
} ## Get Backup Selections from ClientTool.exe

Function Get-BackupErrors {
    if ($Null -eq ($FunctionalProcess)) { 
        "Backup Manager Not Running" 
    }
    else { 
        try { # Command(s) to try
            $ErrorActionPreference = 'Stop'
            $script:Datasources = & $clientool -machine-readable control.selection.list | ConvertFrom-String | Select-Object -skip 1 -property p1,p2 -unique | ForEach-Object {If ($_.P2 -eq "Inclusive") {Write-Output $_.P1}} 
        }
        catch{ # What to do with terminating errors

        }
    }
    Write-output "`n[DataSource Errors]"
    $Script:SessionErrors = @{}
    foreach ($datasource in  $Datasources ) {
        if ($Null -eq ($FunctionalProcess)) {
             "Backup Manager Not Running" 
            }
            else {  
                try { # Command(s) to try
                    $ErrorActionPreference = 'Stop'
                    $sessionerrors.$datasource = & $clientool control.session.error.list -datasource $datasource
                    
                    If ($sessionerrors.$datasource -ne "No session errors found.") {
                        Write-Warning "[$datasource] Errors Found"
                        $sessionerrors.$datasource
                    }
                    else {
                        "`n[$datasource] $($sessionerrors.$datasource)`n"
                    } 
                }
                catch{ # What to do with terminating errors

                }
            }
        }
        

} ## Get Last Error per Active Data Source from ClientTool.exe

Function Get-BackupFilters {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupFilters = & $clientool control.filter.list }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Filters]"

    if ($BackupFilters) {

        $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
        $Hash = "BackupFilters.hash"
        $log = "BackupFilters.log"

        $prior = test-path -PathType Leaf -Path $BackupMonitoringPath\$hash

        if ($prior -eq $false) {
            mkdir -Path $BackupMonitoringPath -ea SilentlyContinue
            new-item -path $BackupMonitoringPath\$hash -ea SilentlyContinue
            
            Write-output "`n$(get-timestamp)`nNo Prior Value" | out-file -FilePath $BackupMonitoringPath\$log -Append
            Write-output "No Prior Hash" | out-file -FilePath $BackupMonitoringPath\$hash
            
        }
        $BackupFiltersHash = @{}
        $BackupFiltersHash.prior = Get-Content -Path $BackupMonitoringPath\$hash
        $BackupFiltersHash.current = Hash-Value $BackupFilters

        if ($BackupFiltersHash.prior -eq $BackupFiltersHash.current ) {
            if ($debugdetail) {
                Write-Output "Prior Hash                  : $($BackupFiltersHash.prior)"
                Write-Output "Current Hash                : $($BackupFiltersHash.current)"
            }
            Write-Output "Values Match"
            #$BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            #$BackupFilters | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Output $BackupFilters
        }elseif ($BackupFiltersHash.prior -ne $BackupFiltersHash.current ){
            Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
            Write-Output "Prior Hash                  : $($BackupFiltersHash.prior)"
            Write-Output "Current Hash                : $($BackupFiltersHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
            $BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$hash
            #$BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            $BackupFilters | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        }
    }
} ## Get Backup Filters from ClientTool.exe

Function Get-BackupSchedules {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSchedules = & $clientool control.schedule.list }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Backup Schedules]"

    if ($BackupSchedules) {

        $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
        $Hash = "BackupSchedules.hash"
        $log = "BackupSchedules.log"

        $prior = test-path -PathType Leaf -Path $BackupMonitoringPath\$hash

        if ($prior -eq $false) {
            mkdir -Path $BackupMonitoringPath -ea SilentlyContinue
            new-item -path $BackupMonitoringPath\$hash -ea SilentlyContinue
            
            Write-output "`n$(get-timestamp)`nNo Prior Value" | out-file -FilePath $BackupMonitoringPath\$log -Append
            Write-output "No Prior Hash" | out-file -FilePath $BackupMonitoringPath\$hash
            
        }
        $BackupSchedulesHash = @{}
        $BackupSchedulesHash.prior = Get-Content -Path $BackupMonitoringPath\$hash
        $BackupSchedulesHash.current = Hash-Value $BackupSchedules

        if ($BackupSchedulesHash.prior -eq $BackupSchedulesHash.current ) {
            if ($debugdetail) {
                Write-Output "Prior Hash                  : $($BackupSchedulesHash.prior)"
                Write-Output "Current Hash                : $($BackupSchedulesHash.current)"
            }
            Write-Output "Values Match"
            #$BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            #$BackupSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Output $BackupSchedules
        }elseif ($BackupSchedulesHash.prior -ne $BackupSchedulesHash.current ){
            Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
            Write-Output "Prior Hash                  : $($BackupSchedulesHash.prior)"
            Write-Output "Current Hash                : $($BackupSchedulesHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
            $BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$hash
            #$BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            $BackupSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        }
    }
} ## Get Backup Schedules from ClientTool.exe

Function Get-BackupSuccess {
    param (
    [Parameter(Mandatory = $false)] [Int]$SessionDays = 7, ## Day Count for Sessions
    [Parameter(Mandatory = $false)] [Int]$DataSourceDays = 30 ## Day Count for Data Sources
    )

    $Ignore = @("InProcess","Completed")  ## Successful / running job history to ignore

    [xml]$sessions = get-content -path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"  ## Session Report to parse

    [datetime]$Start = (get-date).AddDays($DataSourceDays/-1)

    [array]$History = $sessions.SessionStatistics.session | sort-object -Descending StartTimeUTC | Where-Object { ($_.type -eq "Backup") -and ($Start -lt [datetime]$_.starttimeutc) -and ($_.Status -notin "InProcess")}

    $History | ForEach-Object { if ($_.plugin -match 'BackupPlugin') {$_.plugin = $_.plugin -replace("BackupPlugin","") } }
    $History | ForEach-Object { if ($_.plugin -eq 'Fs') {$_.plugin = 'FileSystem'} }
    #$History | ForEach-Object {$_.StartTimeUTC = [datetime]$_.StartTimeUTC } 
    #$History | ForEach-Object {$_.StartTimeUTC = ((Get-Date $_.StartTimeUTC).tostring("yyyy/MM/dd H:mm:ss K"))} 
    #$History | ForEach-Object {$_.EndTimeUTC = ((Get-Date $_.EndTimeUTC).tostring())} 

      $plugins = ($History | select-object Plugin -Unique ).plugin 

    [datetime]$Start = (get-date).AddDays($SessionDays/-1)

    foreach ($plugin in $plugins) {

        $TotalHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.type -eq "Backup") -and ($_.Status -notin "InProcess") -and ($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        $FailHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.type -eq "Backup") -and ($_.Status -notin $Ignore) -and ($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        Write-Output "`n[$plugin Backup Success]"
        $TotalHistoryCount = @($TotalHistory).Count 
        Write-Output "Total Sessions              : $TotalHistoryCount"
        $FailHistoryCount = @($FailHistory).Count
        Write-Output "Unsuccessful Sessions       : $FailHistoryCount"

        $success = (1-$FailHistory.count/$totalHistory.count).ToString('P0') 
        $unsuccess = ($FailHistory.count/$totalHistory.count).ToString('P0')
        Write-Output "Unsuccessful %              : $unsuccess"
        Write-Output "Successful %                : $success"

        $debugdetail=$true
        if ($debugdetail) {
            if ($FailHistory) {$FailHistory | sort-object -Descending Starttimeutc | select-object Type,Plugin,starttimeutc,status,@{l='Errors';e={$_.errorscount}},@{l='Sel GB';e={[Math]::Round([Decimal](($_.SelectedSize) /1GB),3)}},SelectedCount,* -ea SilentlyContinue | Format-Table }else{ Write-Output " No Backup Session Failures Found in Last $MinDays Days"}
        }

    } 

} ## Get Backup Success from SessionReport.xml

Function Get-ArchiveSchedules {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $ArchiveSchedules = & $clientool control.archiving.list }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Archive Schedules]"

    if ($ArchiveSchedules) {

        $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
        $Hash = "ArchiveSchedules.hash"
        $log = "ArchiveSchedules.log"

        $prior = test-path -PathType Leaf -Path $BackupMonitoringPath\$hash

        if ($prior -eq $false) {
            mkdir -Path $BackupMonitoringPath -ea SilentlyContinue
            new-item -path $BackupMonitoringPath\$hash -ea SilentlyContinue
            
            Write-output "`n$(get-timestamp)`nNo Prior Value" | out-file -FilePath $BackupMonitoringPath\$log -Append
            Write-output "No Prior Hash" | out-file -FilePath $BackupMonitoringPath\$hash
            
        }
        $ArchiveSchedulesHash = @{}
        $ArchiveSchedulesHash.prior = Get-Content -Path $BackupMonitoringPath\$hash
        $ArchiveSchedulesHash.current = Hash-Value $ArchiveSchedules

        if ($ArchiveSchedulesHash.prior -eq $ArchiveSchedulesHash.current ) {
            if ($debugdetail) {
                Write-Output "Prior Hash                  : $($ArchiveSchedulesHash.prior)"
                Write-Output "Current Hash                : $($ArchiveSchedulesHash.current)"
            }
            Write-Output "Values Match"
            #$ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            #$ArchiveSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Output $ArchiveSchedules
        }elseif ($ArchiveSchedulesHash.prior -ne $ArchiveSchedulesHash.current ){
            Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
            Write-Output "Prior Hash                  : $($ArchiveSchedulesHash.prior)"
            Write-Output "Current Hash                : $($ArchiveSchedulesHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
            Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
            $ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$hash
            #$ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
            $ArchiveSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        }
    }
} ## Get Archive Schedules from ClientTool.exe

Function Get-ArchiveSessions {

    $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
    $AllSessions = "AllSessions.tsv"

    & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > $BackupMonitoringPath\$AllSessions
    $ArchiveSessions = Import-Csv -Delimiter "`t" -Path $BackupMonitoringPath\$AllSessions

    $ArchiveSessions | ForEach-Object {$_.START = [datetime]$_.START} 
    $ArchiveSessions | ForEach-Object {if ($_.End -ne "-") {$_.End = [datetime]$_.END}} -ErrorAction SilentlyContinue

    if ($debugdetail) {
        Write-Output "All Session History from Clienttool.exe"    
        $ArchiveSessions | format-table
    }
    
    $plugins = ($ArchiveSessions |  Where-Object{ ($_.TYPE -eq "Backup") } | select-object DSRC -Unique).DSRC

    foreach ($plugin in $plugins) {

        # Clienttool Sessions - Filtered
        Write-Output "`n[$plugin Archive Sessions]" 

        $Inventory = $ArchiveSessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TYPE -eq "Backup") -and ($_.DSRC -eq "$plugin") } | Select-Object * | Sort-Object START

        if ($Inventory){
            $ArchiveCount = @($inventory).Count
            $CompletedCount =    @($inventory | Where-Object{ ($_.STATE -eq'Completed') }).count
            $CompletedErrorsCount =    @($inventory | Where-Object{ ($_.STATE -eq'CompletedWithErrors') }).count
            Write-output "Total Attempted Archive Sessions               : $ArchiveCount"          
            Write-output "Total Completed Archive Sessions               : $CompletedCount"
            Write-output "Total Completed w/Error Archive Sessions       : $CompletedErrorsCount"

            if ($CompletedErrorsCount){
                $oldestCompletedErrors = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'CompletedWithErrors') })[0].start
                $latestCompletedErrors = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'CompletedWithErrors') })[-1].start
                Write-output "Oldest Completed w/Error Archive Session UTC   : $($oldestCompletedErrors.ToUniversalTime())"
                Write-output "Latest Completed w/Error Archive Session UTC   : $($latestCompletedErrors.ToUniversalTime())"
                $since = new-timespan -start (get-date -date $latestCompletedErrors)
                Write-output "Days Since Completed w/Error Archive Session   : $($since.days)"
            }

            if ($CompletedCount){
                $oldestCompleted = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[0].start
                $latestCompleted = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[-1].start
                Write-output "Oldest Completed Archive Session UTC           : $($oldestCompleted.ToUniversalTime())"
                Write-output "Latest Completed Archive Session UTC           : $($latestCompleted.ToUniversalTime())"
                $since = new-timespan -start (get-date -date $latestCompleted)
                Write-output "Days Since Completed Archive Session           : $($since.days)"
            }

            if ($debugdetail) {
                Write-output "`n[$plugin Archive Debug]"
                Write-output "`nOldest archive session"   
                ($inventory | sort-object START)[0] | format-table
                Write-output "Latest archive session"   
                ($inventory | sort-object START)[-1] | format-table

                if ($inventory.state -contains "CompletedwithErrors") {
                    Write-output "Oldest Completed w/Error archive session"   
                    ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'CompletedWithErrors') })[0] | format-table
                    Write-output "Latest Completed w/Error archive session"   
                    ($inventory |sort-object START | Where-Object{ ($_.STATE -eq'CompletedWithErrors') })[-1] | format-table
                }
                if ($inventory.state -contains "Completed") {
                    Write-output "Oldest Completed archive session"   
                    ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[0] | format-table
                    Write-output "Latest Completed archive session"   
                    ($inventory |sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[-1] | format-table
                }
                Write-output "10 Oldest archive sessions"
                ($inventory | sort-object START)[0..9] | format-table 
                Write-output "10 Latest archive sessions"   
                ($inventory | sort-object START)[-10..-1] | format-table
            } ## Output if $debugdetail = $true

        } 

    }             

}  ## Get Archive session times via Clienttool.exe

Function Get-StatusReport{

    [xml]$StatusReport = get-content -path 'C:\ProgramData\MXB\Backup Manager\StatusReport.xml' ## StatusReport.XML to Parse

    $data = @{}
    $data = $statusreport.Statistics | Foreach-Object {
        $_ | Select-Object ($_.psobject.Properties | Where-Object Value).Name -ExcludeProperty *xml,*text,*child*,*node*,Schema*,*sibling,Attrib*,owner*
    }

    $data.Account
    $data.InstallationType
    $data.ProfileId
    $data.ProfileVersion
    $data | Select-Object *Colorbar | Format-List
    $data | Select-Object *Retention* | Format-List
    $data | Select-Object *Vault*,*Sync* -ErrorAction SilentlyContinue | Format-List 


    [array]$colorbar = $data | Select-Object *Colorbar
    $colorbar."PluginFileSystem-ColorBar".Replace('8','!').Replace('7','!').Replace('6','?').Replace('5','+').replace("0","X").replace("1",">").replace("2","-")

# -replace("8","!") -replace("7","!") -replace("6","?") -replace("5","+") -replace("2","-") -replace("1",">") -replace("0","X")}  



} ## Get Data from Status.xml

Function Get-StatusReport2{
    $StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $status = @{
        '0' = 'NoBackup (o)'
        '1' = 'InProcess (>)'
        '2' = 'Failed (-)'
        '3' = 'Aborted (x)'
        '4' = 'Unknown (?)'
        '5' = 'Completed (+)'
        '6' = 'Interrupted (&)'
        '7' = 'NotStarted (!)'
        '8' = 'CompletedWithErrors (#)'
        '9' = 'InProgressWithFaults (%) '
        '10' = 'OverQuota ($)'
        '11' = 'NoSelection (0)'
        '12' = 'Restarted (*)'
    }

    $ReplaceArray = @(
        @('9','%'),
        @('8','#'),
        @('7','!'),
        @('6','&'),
        @('5','+'),
        @('4','?'),
        @('3','x'),
        @('2','-'),
        @('1','>'),
        @('0','o')
    )

    $Account                                            = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text" 
    $InstallationKey                                    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//InstallationKey")."#text"
    $TimeStamp                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $StatusCode                                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
    $PartnerName                                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
    $UsedStorage                                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//UsedStorage")."#text"
    $OsVersion                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text"
    $ClientVersion                                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ClientVersion")."#text"
    $MachineName                                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text"
    $IpAddress                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//IpAddress")."#text"
    $DashboardFrequency                                 = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//DashboardFrequency")."#text"
    $DashboardLanguage                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//DashboardLanguage")."#text"
    $TimeZone                                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeZone")."#text"
    $ActivePlugins                                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ActivePlugins")."#text"
    $ActivePlugins_NewNotation                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ActivePlugins_NewNotation")."#text"
    $VirtualUsedStorage                                 = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//VirtualUsedStorage")."#text"
    $SimpleStatus                                       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//SimpleStatus")."#text"
    $AntiCryptoEnabled                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//AntiCryptoEnabled")."#text"
    $LocalSpeedVaultEnabled                             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultEnabled")."#text"
    $BackupServerSynchronizationStatus                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//BackupServerSynchronizationStatus")."#text"
    $LocalSpeedVaultSynchronizationStatus               = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultSynchronizationStatus")."#text"
    $TotalArchivedSize                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TotalArchivedSize")."#text"
    $RetentionUnits                                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//RetentionUnits")."#text"
    $ProfileId                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileId")."#text"
    $ProfileVersion                                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileVersion")."#text"
    $ProxyType                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProxyType")."#text"
  
    "`n[Device]"
    "Device Name             : $Account"
    "Machine Name            : $MachineName"
    "Partner Name            : $PartnerName"
    "TimeStamp UTC           : $(Convert-UnixTimeToDateTime $timestamp)"
    "Timezone                : $TimeZone"  
    #"Installation Key        : $InstallationKey"
    "OS Version              : $OsVersion"
    "Client Version          : $ClientVersion"
    "IP Address              : $IpAddress"
    "Dash Frequency          : $DashboardFrequency"
    "Dash Language           : $DashboardLanguage"
    "Active Plugins          : $ActivePlugins "
    "Active Plugins          : $ActivePlugins_NewNotation"
    #"VirtualUsedStorage      : $($VirtualUsedStorage | RND GB 5)"
    "Used Storage            : $($UsedStorage | RND GB 5)"
    "Archived Size           : $($TotalArchivedSize | RND GB 5)"   
    "Simple Status           : $SimpleStatus"
    "Status Code             : $StatusCode"
    "Documents               : $AntiCryptoEnabled"
    "ProfileId               : $ProfileId"
    "ProfileVersion          : $ProfileVersion"
    "ProxyType               : $ProxyType"
    "`n[Sync]"
    "LSV                     : $LocalSpeedVaultEnabled"
    "LSV Sync Status         : $LocalSpeedVaultSynchronizationStatus"
    "Cloud Sync Status       : $BackupServerSynchronizationStatus"

    #PluginFileSystem
    $TimeStamp                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $StatusCode                                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
    $PluginFileSystemColorBar                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-ColorBar")."#text"

    if ($PluginFileSystemColorBar) {
        $replaceArray | ForEach-Object { $PluginFileSystemColorBar = $PluginFileSystemColorBar -replace $_[0],$_[1]}
        $PluginFileSystemLastSessionStatus                  = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionStatus")."#text"]
        $PluginFileSystemLastSessionTimestamp               = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionTimestamp")."#text")
        $PluginFileSystemLastSuccessfulSessionStatus        = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSuccessfulSessionStatus")."#text"]
        $PluginFileSystemLastSuccessfulSessionTimestamp     = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSuccessfulSessionTimestamp")."#text")
        $PluginFileSystemLastCompletedSessionStatus         = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastCompletedSessionStatus")."#text"]
        $PluginFileSystemLastCompletedSessionTimestamp      = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastCompletedSessionTimestamp")."#text")
        $PluginFileSystemPreRecentSessionSelectedSize       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-PreRecentSessionSelectedSize")."#text"
        $PluginFileSystemLastSessionSelectedSize            = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionSelectedSize")."#text"
        $PluginFileSystemLastSessionProcessedSize           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionProcessedSize")."#text"
        $PluginFileSystemLastSessionSentSize                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionSentSize")."#text"
        $PluginFileSystemPreRecentSessionSelectedCount      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-PreRecentSessionSelectedCount")."#text"
        $PluginFileSystemLastSessionSelectedCount           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionSelectedCount")."#text"
        $PluginFileSystemLastSessionProcessedCount          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionProcessedCount")."#text"
        $PluginFileSystemLastSessionErrorsCount             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionErrorsCount")."#text"
        #$PluginFileSystemProtectedSize                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-ProtectedSize")."#text"
        $PluginFileSystemSessionDuration                    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-SessionDuration")."#text"
        #$PluginFileSystemLastSessionLicenceItemsCount       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionLicenceItemsCount")."#text"
        $PluginFileSystemRetention                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-Retention")."#text"

        "`n[FileSystem Session]"
        "TimeStamp UTC           : $(Convert-UnixTimeToDateTime $timestamp)"
        "Status Code             : $StatusCode"
        "28-Day Status           : $PluginFileSystemColorBar"
        "Retention               : $PluginFileSystemRetention $RetentionUnits"   
        "Session Duration        : $($PluginFileSystemSessionDuration/60) Mins"                                           
        "Last Session            : $PluginFileSystemLastSessionTimestamp $PluginFileSystemLastSessionStatus "
        "Last Success            : $PluginFileSystemLastSuccessfulSessionTimestamp $PluginFileSystemLastSuccessfulSessionStatus"
        "Last Complete           : $PluginFileSystemLastCompletedSessionTimestamp $PluginFileSystemLastCompletedSessionStatus"
        "Last Error Count        : $PluginFileSystemLastSessionErrorsCount"

        if($statuscode -notlike "Backup, Files and folders*") {
            "`n[Size]"
            "Last Sel Size           : $($PluginFileSystemLastSessionSelectedSize | RND GB 5)"
            "Last Proc Size          : $($PluginFileSystemLastSessionProcessedSize | RND GB 5)"
            "Last Sent Size          : $($PluginFileSystemLastSessionSentSize | RND MB 5)" 
            "Last Proc/Sel Size%     : $(($PluginFileSystemLastSessionProcessedSize/$PluginFileSystemLastSessionSelectedSize)*100)"   
            "Last Sent/Sel Size%     : $(($PluginFileSystemLastSessionSentSize/$PluginFileSystemLastSessionSelectedSize)*100)"    
            "`n[Size Change]"
            "Prior Sel Size          : $($PluginFileSystemPreRecentSessionSelectedSize | RND GB 5)"
            "Sel Change Size         : $(($PluginFileSystemLastSessionSelectedSize-$PluginFileSystemPreRecentSessionSelectedSize) | RND MB 5)"
            "Sel Change Size %       : $(([Math]::Abs((($PluginFileSystemLastSessionSelectedSize-$PluginFileSystemPreRecentSessionSelectedSize)/$PluginFileSystemPreRecentSessionSelectedSize)*100)) | rnd '' 4)"  
            "`n[Count]"
            "Last Sel Count          : $PluginFileSystemLastSessionSelectedCount"
            "Last Proc Count         : $PluginFileSystemLastSessionProcessedCount"
            "`n[Count Change]"
            "Prior Sel Count         : $PluginFileSystemPreRecentSessionSelectedCount"
            "Sel Change Count        : $($PluginFileSystemLastSessionSelectedCount-$PluginFileSystemPreRecentSessionSelectedCount)" 
            "Sel Change Count %      : $(([Math]::Abs((($PluginFileSystemLastSessionSelectedCount-$PluginFileSystemPreRecentSessionSelectedCount)/$PluginFileSystemPreRecentSessionSelectedCount)*100)) | rnd '' 4) "     
            #"Protected Size          : $($PluginFileSystemProtectedSize | RND GB 5)"                 
            #"License Count           : $PluginFileSystemLastSessionLicenceItemsCount"
        }
    }

    #PluginVssSystemState
    $TimeStamp                                              = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $StatusCode                                             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
    $PluginVssSystemStateColorBar                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
    if ($PluginVssSystemStateColorBar) {
        $replaceArray | ForEach-Object { $PluginVssSystemStateColorBar  = $PluginVssSystemStateColorBar  -replace $_[0],$_[1]}
        $PluginVssSystemStateLastSessionStatus                  = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionStatus")."#text"]
        $PluginVssSystemStateLastSessionTimestamp               = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionTimestamp")."#text")
        $PluginVssSystemStateLastSuccessfulSessionStatus        = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionStatus")."#text"]
        $PluginVssSystemStateLastSuccessfulSessionTimestamp     = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionTimestamp")."#text")
        $PluginVssSystemStateLastCompletedSessionStatus         = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastCompletedSessionStatus")."#text"]
        $PluginVssSystemStateLastCompletedSessionTimestamp      = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastCompletedSessionTimestamp")."#text")
        $PluginVssSystemStatePreRecentSessionSelectedSize       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-PreRecentSessionSelectedSize")."#text"
        $PluginVssSystemStateLastSessionSelectedSize            = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSelectedSize")."#text"
        $PluginVssSystemStateLastSessionProcessedSize           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionProcessedSize")."#text"
        $PluginVssSystemStateLastSessionSentSize                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSentSize")."#text"
        $PluginVssSystemStatePreRecentSessionSelectedCount      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-PreRecentSessionSelectedCount")."#text"
        $PluginVssSystemStateLastSessionSelectedCount           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSelectedCount")."#text"
        $PluginVssSystemStateLastSessionProcessedCount          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionProcessedCount")."#text"
        $PluginVssSystemStateLastSessionErrorsCount             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionErrorsCount")."#text"
        #$PluginVssSystemStateProtectedSize                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ProtectedSize")."#text"
        $PluginVssSystemStateSessionDuration                    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-SessionDuration")."#text"
        #$PluginVssSystemStateLastSessionLicenceItemsCount       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionLicenceItemsCount")."#text"
        $PluginVssSystemStateRetention                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-Retention")."#text"

        "`n[SystemState Session]"
        "TimeStamp UTC           : $(Convert-UnixTimeToDateTime $timestamp)"
        "Status Code             : $StatusCode"
        "28-Day Status           : $PluginVssSystemStateColorBar"
        "Retention               : $PluginVssSystemStateRetention $RetentionUnits"   
        "Session Duration        : $($PluginVssSystemStateSessionDuration/60) Mins"                                           
        "Last Session            : $PluginVssSystemStateLastSessionTimestamp $PluginVssSystemStateLastSessionStatus "
        "Last Success            : $PluginVssSystemStateLastSuccessfulSessionTimestamp $PluginVssSystemStateLastSuccessfulSessionStatus"
        "Last Complete           : $PluginVssSystemStateLastCompletedSessionTimestamp $PluginVssSystemStateLastCompletedSessionStatus"
        "Last Error Count        : $PluginVssSystemStateLastSessionErrorsCount"
        
        if($statuscode -notlike "Backup, System state (VSS)*") {
            "`n[Size]"
            "Last Sel Size           : $($PluginVssSystemStateLastSessionSelectedSize | RND GB 5)"
            "Last Proc Size          : $($PluginVssSystemStateLastSessionProcessedSize | RND GB 5)"
            "Last Sent Size          : $($PluginVssSystemStateLastSessionSentSize | RND MB 5)" 
            "Last Proc/Sel Size%     : $(($PluginVssSystemStateLastSessionProcessedSize/$PluginVssSystemStateLastSessionSelectedSize)*100)"
            "Last Sent/Sel Size%     : $(($PluginVssSystemStateLastSessionSentSize/$PluginVssSystemStateLastSessionSelectedSize)*100)"    
            "`n[Size Change]"
            "Prior Sel Size          : $($PluginVssSystemStatePreRecentSessionSelectedSize | RND GB 5)"
            "Sel Change Size         : $(($PluginVssSystemStateLastSessionSelectedSize-$PluginVssSystemStatePreRecentSessionSelectedSize) | RND MB 5)"
            "Sel Change Size %       : $(([Math]::Abs((($PluginVssSystemStateLastSessionSelectedSize-$PluginVssSystemStatePreRecentSessionSelectedSize)/$PluginVssSystemStatePreRecentSessionSelectedSize)*100)) | rnd '' 4)"  
            "`n[Count]"
            "Last Sel Count          : $PluginVssSystemStateLastSessionSelectedCount"
            "Last Proc Count         : $PluginVssSystemStateLastSessionProcessedCount"
            "`n[Count Change]"
            "Prior Sel Count         : $PluginVssSystemStatePreRecentSessionSelectedCount"
            "Sel Change Count        : $($PluginVssSystemStateLastSessionSelectedCount-$PluginVssSystemStatePreRecentSessionSelectedCount)" 
            "Sel Change Count %      : $(([Math]::Abs((($PluginVssSystemStateLastSessionSelectedCount-$PluginVssSystemStatePreRecentSessionSelectedCount)/$PluginVssSystemStatePreRecentSessionSelectedCount)*100)) | rnd '' 4)"     
            #"Protected Size          : $($PluginVssSystemStateProtectedSize | RND GB 5)"                 
            #"License Count           : $PluginVssSystemStateLastSessionLicenceItemsCount"
        }
    }


}






#endregion Functions






<#


$PluginVssSystemStateLastSessionStatus                 = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionStatus")."#text"]
$PluginVssSystemStateLastSessionSelectedCount          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSelectedCount")."#text"
$PluginVssSystemStateLastSessionProcessedCount         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionProcessedCount")."#text"
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateLastSessionSelectedSize           = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSelectedSize")."#text"/1GB)
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateLastSessionProcessedSize          = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionProcessedSize")."#text"/1GB)
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateLastSessionSentSize               = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionSentSize")."#text"/1GB)
$PluginVssSystemStateLastSessionErrorsCount            = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionErrorsCount")."#text"
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateProtectedSize                     = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ProtectedSize")."#text"/1GB)
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateLastSuccessfulSessionTimestamp    = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionTimestamp")."#text")
$PluginVssSystemStatePreRecentSessionSelectedCount     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-PreRecentSessionSelectedCount")."#text"
$PluginVssSystemStatePreRecentSessionSelectedSize      = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-PreRecentSessionSelectedSize")."#text"/1GB)
$PluginVssSystemStateColorBar                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"
$PluginVssSystemStateSessionDuration                   = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-SessionDuration")."#text"
$PluginVssSystemStateLastSessionLicenceItemsCount      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionLicenceItemsCount")."#text"
$PluginVssSystemStateRetention                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-Retention")."#text"
$PluginVssSystemStateLastSessionTimestamp              = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionTimestamp")."#text")
$PluginVssSystemStateLastSuccessfulSessionStatus       = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionStatus")."#text"]
$PluginVssSystemStateLastCompletedSessionStatus        = $status[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastCompletedSessionStatus")."#text"]
$PluginVssSystemStateLastCompletedSessionTimestamp     = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastCompletedSessionTimestamp")."#text")




$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
$ = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"

<#

$lastBackup = ([Xml] (get-content "C:\Users\badams\Scripts\backuplog.xml")).SelectSingleNode("//LastBackup")."#text" -as [DateTime]

#>

clear-host

Get-BackupIntegrationStatus



Get-BackupService;Get-BackupProcess;Get-BackupFPVisa;Get-SystemInfo

Get-InitError;Get-ApplicationStatus;Get-JobStatus;Get-VssStatus;Get-StorageStatus

Get-BackupSettings;Get-BackupSelections;Get-BackupFilters;Get-BackupSchedules;Get-ArchiveSchedules

Get-BackupErrors;Get-BackupSuccess;Get-ArchiveSessions

get-statusreport2





Function Read-config {
$storage = "c:\programdata\MXB\Backup Manager\storage\"  

$MXB = gci -path $storage -File "$account.info" -Recurse    

$a = Get-Content $mxb.FullName | convertFrom-Json

$MXBcfg = gci -path $storage -File "$account.cfg" -Recurse    

$b = Get-Content $mxbcfg.FullName | convertFrom-Json
$b = $b.Features.split("`n")

$b = $b -notlike "*=0"
$a
$b
($b -like "LegalityStatus=*").split("=")[1]

}