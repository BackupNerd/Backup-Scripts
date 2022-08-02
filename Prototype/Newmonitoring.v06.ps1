Clear-host

#region ----- Environment, Variables, Names and Paths ----

$debugdetail = $true
#$debugdetail = $false
$clientool = "C:\Program Files\Backup Manager\ClientTool.exe"
$config = "C:\Program Files\Backup Manager\config.ini" 

#endregion ----- Environment, Variables, Names and Paths ----

#region Functions
-Function Hash-Value ($value) {
    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding
    $hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($value)))
    $hash
}

Function Get-TimeStamp {
    return "[{0:yyyy/MM/dd} {0:HH:mm:ss}]" -f (Get-Date)
}

Function Get-BackupService {
    $BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
    if ($BackupService.status -eq "Running"){ 
        Write-output "[Backup Service Controller] `nService Status        : $($BackupService.status)"
    }elseif ($BackupService.status -ne "Running"){
        Write-warning "[Backup Service Controller] 'nService Status        : $($BackupService.status)"
    }
} ## Get Backup Service Status

Function Get-BackupProcess {
    $Script:FunctionalProcess = get-process "BackupFP" -ea SilentlyContinue
    if ($FunctionalProcess) {
        $CurrentTime = Get-Date
        $ProcessDuration = New-TimeSpan -Start $FunctionalProcess.StartTime -End $CurrentTime
        Write-Output "`n[Backup Functional Process] `nProcess Responding    : $($FunctionalProcess.Responding)"
        Write-Output "Process Version       : $($FunctionalProcess.ProductVersion)"
        Write-Output "Process Start         : $($FunctionalProcess.StartTime)"
        Write-Output "Process Uptime        : $ProcessDuration"

        $SleepSeconds = 5
        $Cores = (Get-WmiObject -class win32_processor -Property numberOfCores).numberOfCores;
        $LogicalCores = (Get-WmiObject -class Win32_processor -Property NumberOfLogicalProcessors).NumberOfLogicalProcessors;
         
        $cpu1 = (get-process -name Backupfp).cpu
        start-sleep -Seconds $SleepSeconds  
        $cpu2 = (get-process -name Backupfp).cpu
        $percent = (($cpu2 - $cpu1)/($LogicalCores*$sleepseconds)).ToString('P0')
        Write-Output "Process CPU %         : $percent"
        Write-Output "Physical Cores        : $Cores"
        Write-Output "Logical Cores         : $LogicalCores"

        $CompObject =  Get-WmiObject -Class WIN32_OperatingSystem
        $Memory = (($CompObject.TotalVisibleMemorySize - $CompObject.FreePhysicalMemory)/ $CompObject.TotalVisibleMemorySize)
        $TotalMemory = $CompObject.TotalVisibleMemorySize / 1mb
       
        Write-Output "System RAM GB Used    : $([math]::round($totalMemory,2))"
        Write-Output "System RAM % Used     : $($Memory.ToString('P0'))"
        
        # BackupFP Memory usage
        $processMemoryUsage = Get-WmiObject WIN32_PROCESS | Sort-Object -Property ws -Descending | where-object {$_.Processname -eq "Backupfp.exe"} | Select-Object processname, @{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}}
       
        write-output "BackupFP RAM MB Used  : $($processMemoryUsage.'Mem Usage(MB)')"

    }
} ## Get BackupFP Process Info

Function Get-InitError {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $initerror = & $clientool control.initialization-error.get  | convertfrom-json}catch{ Write-Warning "ERROR: $_`n" }}
    if ($initerror.code -gt 0) {"`n[Cloud Initilization] `nCloud Init            : $($initerror.Message)`n"}else{ "`n[Cloud Initilization] `nCloud Init            : Ok"}
} ## Get Backup Manager Initialization Errors from ClientTool.exr

Function Get-ApplicationStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $AppStatus = & $clientool control.application-status.get }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Backup Manager] `nApplication Status    : $appstatus"
} ## Get Backup Application Status from ClientTool.exe

Function Get-BackupFPVisa {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupFPVisa = & $clientool in-agent-authentication-token.get -config-path "c:\program files\backup manager\config.ini" | convertfrom-json }catch{ Write-Warning "ERROR: $_" }} 
    Write-output "`n[Backup Manager] `nClient Visa           : $($BackupFPVisa.InAgentAuthenticationToken)"
} ## Get BackupFP Authentication Visa from ClientTool.exe

Function Get-JobStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $JobStatus = & $clientool control.status.get }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Backup Job] `nJob Status            : $jobStatus"
} ## Get Backup Job Status from ClientTool.exe

Function Get-VssStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $VssStatus = & $clientool vss.check }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[VSS Status]"
    #($VssStatus -split '\r?\n').Trim()
    ($VssStatus | where-object {($_ -match 'error') -or ($_ -notmatch 'ok')  })
} ## Check VSS Status from ClientTool.exe

Function Get-StorageStatus {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop' ; $StorageStatus = & $clientool storage.test }catch{ Write-Warning "Oops: $_" }}
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
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSettings = & $clientool control.setting.list }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Backup Settings]"

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
            Write-Output "Prior Hash             : $($BackupSettingsHash.prior)"
            Write-Output "Current Hash           : $($BackupSettingsHash.current)"
        }
        Write-Output "Values Match"
        #$BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        #$BackupSettings | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Output $BackupSettings
    }elseif ($BackupSettingsHash.prior -ne $BackupSettingsHash.current ){
        Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
        Write-Output "Prior Hash             : $($BackupSettingsHash.prior)"
        Write-Output "Current Hash           : $($BackupSettingsHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
        $BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$hash
        #$BackupSettingsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        $BackupSettings | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
    }
} ## Get Backup Settings from ClientTool.exe

Function Get-BackupSelections {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSelections = & $clientool control.selection.list }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Backup Selections]"

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
            Write-Output "Prior Hash             : $($BackupSelectionsHash.prior)"
            Write-Output "Current Hash           : $($BackupSelectionsHash.current)"
        }
        Write-Output "Values Match"
        #$BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        #$BackupSelections | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Output $BackupSelections
    }elseif ($BackupSelectionsHash.prior -ne $BackupSelectionsHash.current ){
        Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
        Write-Output "Prior Hash             : $($BackupSelectionsHash.prior)"
        Write-Output "Current Hash           : $($BackupSelectionsHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
        $BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$hash
        #$BackupSelectionsHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        $BackupSelections | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
    }
} ## Get Backup Selections from ClientTool.exe

Function Get-DataSourceErrors {
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
    Write-output "`n[DataSources]"
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
                        Write-Warning "`n[$datasource] Errors Found"
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
        

} ## Get Backup Selections from ClientTool.exe


Function Get-BackupFilters {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupFilters = & $clientool control.filter.list }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Backup Filters]"

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
            Write-Output "Prior Hash             : $($BackupFiltersHash.prior)"
            Write-Output "Current Hash           : $($BackupFiltersHash.current)"
        }
        Write-Output "Values Match"
        #$BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        #$BackupFilters | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Output $BackupFilters
    }elseif ($BackupFiltersHash.prior -ne $BackupFiltersHash.current ){
        Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
        Write-Output "Prior Hash             : $($BackupFiltersHash.prior)"
        Write-Output "Current Hash           : $($BackupFiltersHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
        $BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$hash
        #$BackupFiltersHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        $BackupFilters | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
    }
} ## Get Backup Filters from ClientTool.exe

Function Get-BackupSchedules {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSchedules = & $clientool control.schedule.list }catch{ Write-Warning "ERROR: $_" }}

    #$BackupSchedules = $BackupSchedules -replace('nesday','') -replace('urday','') -replace('day','') -replace('NetworkShares','NWShares') -replace('Vss','')

    Write-output "`n[Backup Schedules]"

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
            Write-Output "Prior Hash             : $($BackupSchedulesHash.prior)"
            Write-Output "Current Hash           : $($BackupSchedulesHash.current)"
        }
        Write-Output "Values Match"
        #$BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        #$BackupSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Output $BackupSchedules
    }elseif ($BackupSchedulesHash.prior -ne $BackupSchedulesHash.current ){
        Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
        Write-Output "Prior Hash             : $($BackupSchedulesHash.prior)"
        Write-Output "Current Hash           : $($BackupSchedulesHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
        $BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$hash
        #$BackupSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        $BackupSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
    }
} ## Get Backup Schedules from ClientTool.exe

Function Get-ArchiveSchedules {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $ArchiveSchedules = & $clientool control.archiving.list }catch{ Write-Warning "ERROR: $_" }}
    Write-output "`n[Archive Schedules]"

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
            Write-Output "Prior Hash             : $($ArchiveSchedulesHash.prior)"
            Write-Output "Current Hash           : $($ArchiveSchedulesHash.current)"
        }
        Write-Output "Values Match"
        #$ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        #$ArchiveSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Output $ArchiveSchedules
    }elseif ($ArchiveSchedulesHash.prior -ne $ArchiveSchedulesHash.current ){
        Write-Output "`n$(get-timestamp)" | Out-File -FilePath  $BackupMonitoringPath\$log -Append
        Write-Output "Prior Hash             : $($ArchiveSchedulesHash.prior)"
        Write-Output "Current Hash           : $($ArchiveSchedulesHash.current)" | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
        Write-Warning "Value Mismatch `nTo view prior values check the Backup.Management device audit `nor access $BackupMonitoringPath\$log on the local device"
        $ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$hash
        #$ArchiveSchedulesHash.current | out-file -FilePath $BackupMonitoringPath\$log -Append
        $ArchiveSchedules | Tee-Object -FilePath $BackupMonitoringPath\$log -Append
    }
} ## Get Archive Schedules from ClientTool.exe



Function get-success {
    param (
    [Parameter(Mandatory = $false)] [Int]$SessionDays = 7, ## Day Count for Sessions
    [Parameter(Mandatory = $false)] [Int]$DataSourceDays = 30 ## Day Count for Data Sources
    )

    $Ignore = @("InProcess","Completed")  ## Successful / running job history to ignore

    [xml]$sessions = get-content -path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"  ## Session Report to parse

    [datetime]$Start = (get-date).AddDays($DataSourceDays/-1)

    $History = $sessions.SessionStatistics.session | sort-object -Descending Starttimeutc | Where-Object { ($Start -lt $_.starttimeutc) -and ($_.Status -notin "InProcess")}

    $History | ForEach-Object { if ($_.plugin -match 'BackupPlugin') {$_.plugin = $_.plugin -replace("BackupPlugin","") } }

    $plugins = ($History | select-object Plugin -Unique).plugin

    [datetime]$Start = (get-date).AddDays($SessionDays/-1)

    foreach ($plugin in $plugins) {

        $TotalHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.Status -notin "InProcess") -and($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        $FailHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.Status -notin $Ignore) -and ($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        Write-Output "Data Source            : $($plugin)"
        Write-Output "Total Sessions         : $($totalHistory.count)"
        Write-Output "Unsuccessful Sessions  : @($FailHistory).count)" 
        $success = (1-$FailHistory.count/$totalHistory.count).ToString('P0') 
        $unsuccess = ($FailHistory.count/$totalHistory.count).ToString('P0')
        Write-Output "Unsuccessful %         : $unsuccess"
        Write-Output "Successful %           : $success"

        if ($debugdetail) {
            if ($FailHistory) {$FailHistory | select-object Type,Plugin,starttimeutc,status,@{l='Errors';e={$_.errorscount}},@{l='Sel GB';e={[Math]::Round([Decimal](($_.SelectedSize) /1GB),3)}},SelectedCount,* -ea SilentlyContinue | Format-Table }else{ Write-Output " No Backup Session Failures Found in Last $MinDays Days"}
        }

    } 



}

Function get-archives {

    $Ignore = @("InProcess","Completed")  ## Successful / running job history to ignore

    [xml]$sessions = get-content -path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"  ## Session Report to parse

    [datetime]$Start = (get-date).AddDays(-1)

    $History = $sessions.SessionStatistics.session | sort-object -Descending Starttimeutc | Where-Object {($_.Status -notin "InProcess")}   

    $History | ForEach-Object { if ($_.plugin -match 'BackupPlugin') {$_.plugin = $_.plugin -replace("BackupPlugin","") } }

    $archived = $History | sort-object -Descending Starttimeutc | Where-Object {($_.Archived -eq 1)}   

    $plugins = ($History | select-object Plugin -Unique).plugin

    [datetime]$Start = (get-date).AddDays($SessionDays/-1)

    foreach ($plugin in $plugins) {

        $TotalHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.Status -notin "InProcess") -and($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        $FailHistory = $History | sort-object -Descending Starttimeutc | Where-Object { ($_.Status -notin $Ignore) -and ($start -lt $_.starttimeutc) -and ($_.plugin -eq $plugin)}

        Write-Output "Data Source            : $($plugin)"
        $totalHistorycount = @($totalHistory).Count
        Write-Output "Total Sessions         : $totalHistorycount"
        $failHistorycount = @($FailHistory).Count
        Write-Output "Unsuccessful Sessions  : $failHistorycount" 
        $success = (1-$FailHistory.count/$totalHistory.count).ToString('P0') 
        $unsuccess = ($FailHistory.count/$totalHistory.count).ToString('P0')
        Write-Output "Unsuccessful %         : $unsuccess"
        Write-Output "Successful %           : $success"

        if ($FailHistory) {$FailHistory | select-object Type,Plugin,starttimeutc,status,@{l='Errors';e={$_.errorscount}},@{l='Sel GB';e={[Math]::Round([Decimal](($_.SelectedSize) /1GB),3)}},SelectedCount,* -ea SilentlyContinue | Format-Table }else{ Write-Output " No Backup Session Failures Found in Last $MinDays Days"}

    } 



}


Function ClienttoolGetArchiveSessions {

    $BackupMonitoringPath = "C:\programdata\mxb\Monitoring"
    $AllSessions = "AllSessions.tsv"

    & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > $BackupMonitoringPath\$AllSessions
    $ArchiveSessions = Import-Csv -Delimiter "`t" -Path $BackupMonitoringPath\$AllSessions

    $ArchiveSessions | ForEach-Object {$_.START = [datetime]$_.START} 
    $ArchiveSessions | ForEach-Object {$_.End = [datetime]$_.END} -ErrorAction SilentlyContinue

    if ($debugdetail) {
        Write-Output "  All Session History from Clienttool.exe"    
        $ArchiveSessions | format-table
    }
    
    $plugins = ($ArchiveSessions |  Where-Object{ ($_.TYPE -eq "Backup") } | select-object DSRC -Unique).DSRC

    foreach ($plugin in $plugins) {

        # Clienttool Sessions - Filtered
         Write-Output "`n[$plugin]" 

        $Inventory = $ArchiveSessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TYPE -eq "Backup") -and ($_.DSRC -eq "$plugin") } | Select-Object * | Sort-Object START



        $count = @($inventory).Count
        $SuccessCount =    @($inventory | Where-Object{ ($_.STATE -eq'Completed')}).count
        Write-output "Total Attempted Archive Sessions  : $count"          
        Write-output "Total Successful Archive Sessions : $SuccessCount"

        $oldest = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[0].start
        $latest = ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[-1].start
        Write-output "Oldest Successful Archive Session : $oldest"
        Write-output "Latest Successful Archive Session : $latest"
        $since = new-timespan -end (get-date -date $latest)
        Write-output "Days Since Last Successful Archive Session : $($since.days)"

        if ($debugdetail) {
            Write-output "10 oldest archive sessions"
            ($inventory | sort-object START)[0..9] | format-table 
            Write-output "10 newest archive sessions"   
            ($inventory | sort-object START)[-10..-1] | format-table
            Write-output "Oldest archive session"   
            ($inventory | sort-object START)[0] | format-table
            Write-output "Oldest successful archive session"   
            ($inventory | sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[0] | format-table
            Write-output "Latest archive session"   
            ($inventory | sort-object START)[-1] | format-table
            Write-output "Latest successful archive session"   
            ($inventory |sort-object START | Where-Object{ ($_.STATE -eq'Completed') })[-1] | format-table
        } ## Output if $debugdetail = $true



    }             

    #$Script:CleanSessions = $ArchiveSessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TyPE -eq "Backup") -and ([datetime]$_.START -lt $filterDate) } | Select-Object -Property START,DSRC,STATE,TYPE,FLAGS,SELS | Sort-Object START 



    #Write-Output "  All Archive Session History from Clienttool.exe older than $filterdate [datetime]"
    #$Script:CleanSessions |Select-Object -Property DSRC,START,STATE,TYPE,FLAGS,SELS | Sort-Object START | format-table
}  ## Get Archive session times via Clienttool.exe


Function Is-integrated {

    test-path -PathType Leaf -Path "C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\MSPBackupManagerConfig.xml"

}

#endregion Functions


clear-host;Get-BackupService;Get-BackupProcess;Get-InitError;Get-ApplicationStatus;Get-JobStatus;Get-VssStatus;Get-StorageStatus

get-backupsettings;Get-BackupSelections;Get-BackupFilters;Get-BackupSchedules;Get-ArchiveSchedules

Get-DataSourceErrors;get-success

ClienttoolGetArchiveSessions
