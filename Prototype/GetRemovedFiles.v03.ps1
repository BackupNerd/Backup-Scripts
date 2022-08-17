[CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)] [Int]$SessionDays = 7, ## Day Count for Sessions
        [Parameter(Mandatory = $False)] [switch]$debugdetail   
    )        

Clear-host

#region ----- Environment, Variables, Names and Paths ----

$clientool = "C:\Program Files\Backup Manager\ClientTool.exe"
$config = "C:\Program Files\Backup Manager\config.ini" 
[datetime]$Start = (get-date).AddDays($DataSourceDays/-1)
$CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"

#endregion ----- Environment, Variables, Names and Paths ----




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

Function Get-BackupSessions {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $SessionList = & $clientool -machine-readable control.session.list -delimiter ',' }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Getting Backup Sessions]"

    if ($SessionList) {
        $SessionList | out-file SessionList.csv
        $BackupSessionList = Import-Csv SessionList.csv
        $Script:SLIST = $BackupSessionList | Sort-Object start | Where-Object { ($_.type -eq "Backup") -and ($_.dsrc -eq "FileSystem") -and ($_.remc -gt 0) } | Select-Object Start,Remc,* -ExcludeProperty Flags,end -ea SilentlyContinue    
        $Script:SLIST | Format-Table      
       
    }
} ## Get Backup Sessions from ClientTool.exe

Function Get-RemovedFiles ($starttime) {
    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $FileList = & $clientool -machine-readable control.session.node.list -removed -datasource FileSystem -delimiter ',' -time $starttime  }catch{ Write-Warning "ERROR     : $_" }}
    Write-output "`n[Getting Removed Files from $starttime]"

    if ($FileList) {
        $FileList | out-file TempFileList.csv 
        $script:RemovedFileList = Import-Csv TempFileList.csv
        $script:RemovedFileList | foreach-object {$_.start = $starttime} 
        $script:RemovedFileList = $script:RemovedFileList | Select-Object START,@{Name = 'Present';Expression = { test-path $_.Path }},@{Name = 'LastWrite';Expression = { if ($_.present -eq "true") {(get-item $_.Path).lastwritetime.tostring()} }},PATH | Export-Csv -Path .RemovedFileList_$currentdate.csv -NoTypeInformation -Append
        
        #$script:RemovedFileList | out-gridview

      
    }
} ## Get REmoved from ClientTool.exe

Get-BackupService
Get-BackupProcess
Get-BackupSessions

foreach ($session in $slist) {
    Get-RemovedFiles $session.start

}


Function Get-ModificationTime {

    (get-item $_.Path).lastwritetime.tostring()

}
