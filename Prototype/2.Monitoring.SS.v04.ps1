<# ----- About: ----
    # N-able | Cove Data Protection Monitor 2 - SystemState
    # Revision v02 - 2022-03-28
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
        [Parameter(Mandatory=$False)] [switch]$ThresholdValue = $true    ## Set True to Output Threshold Values seperately for Debugging
    )        
cls
#Requires -Version 5.1 -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection - SystemState Monitor"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

#region Functions

Clear-Host

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

Function Get-BackupState {
 
    $Script:BackupService = get-service -name $BackupServiceName -ea SilentlyContinue
    Write-output "`n[Backup Service Controller]"
    if ($BackupService.status -eq "Running"){
        $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
    }
} ## Get Backup Service \ Process Status

Get-BackupState

Function Get-LSVstatus ([int]$SynchThreshold) {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $Script:LSVSettings = & $clienttool control.setting.list }catch{ Write-Warning "ERROR     : $_" }}

    $Script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $DocumentsEnabled                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//AntiCryptoEnabled")."#text"
    $Device                                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
    $MachineName                                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text" 
    $TimeStamp                                  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $LocalSpeedVaultEnabled                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultEnabled")."#text"
    $BackupServerSynchronizationStatus          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//BackupServerSynchronizationStatus")."#text"
    $LocalSpeedVaultSynchronizationStatus       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//LocalSpeedVaultSynchronizationStatus")."#text"
    $SelectedSize                               = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginTotal-LastSessionSelectedSize")."#text"
    $UsedStorage                                = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//UsedStorage")."#text"
    if ($LocalSpeedVaultEnabled -eq 1) {
    $LSVPath                                    = ($LSVSettings | Where-Object { $_ -like "LocalSpeedVaultLocation *"}).replace("LocalSpeedVaultLocation ","")
    $LSVUser                                    = ($LSVSettings | Where-Object { $_ -like "LocalSpeedVaultUser *"}).replace("LocalSpeedVaultUser     ","")
    }


    if ($DocumentsEnabled -eq 1) { 
        Write-Warning "LocalSpeedVault Not Supported on Documents Devices" 
    }else{
        $Global:LSVEnabledOutTxt = $LocalSpeedVaultEnabled.replace("1","True").replace("0","False")
        $Global:LSVEnabledOutVal = $LocalSpeedVaultEnabled
        $Global:LSVSyncStatusOutTxt = $LocalSpeedVaultSynchronizationStatus
        $Global:CloudSyncStatusOutTxt = $BackupServerSynchronizationStatus
        $Global:TotalSelectedGBTxt = ($SelectedSize | RND GB 2)
        $Global:TotalUsedGBTxt = ($UsedStorage | RND GB 2)

        Write-Output "`n[LocalSpeedVault Status]"
        Write-Output "Device                  : $Device"
        Write-Output "Machine                 : $MachineName"
        Write-Output "TimeStamp (UTC)         : $(Convert-UnixTimeToDateTime $timestamp)"
        Write-Output "LSV Enabled             : $Global:LSVEnabledOutTxt"
        Write-Output "LSV Sync Status         : $Global:LSVSyncStatusOutTxt"
        Write-Output "Cloud Sync Status       : $Global:CloudSyncStatusOutTxt"
        Write-Output "Total Selected Size     : $Global:TotalSelectedGBTxt"        
        Write-Output "Est Space Required      : $Global:TotalUsedGBTxt"
        Write-Output "Last LSV Path           : $LSVpath"
        Write-Output "Last LSV User           : $LSVUser"

        # Fail if LSV Enabled = True & ( LSV Sync = Failed or Cloud Sync = Failed )
        if (($LocalSpeedVaultEnabled -eq 1) -and (($BackupServerSynchronizationStatus -eq "Failed") -or ($LocalSpeedVaultSynchronizationStatus -eq "Failed"))) {Write-Warning "LSV Failed"}
                
        if (($LocalSpeedVaultEnabled -eq 1) -and ($BackupServerSynchronizationStatus -ne "synchronized")) { 
            if ( ($BackupServerSynchronizationStatus.replace("%","")/1 -lt $SynchThreshold )){Write-Warning "Cloud sync is below $SynchThreshold%"}
        } ## Warn if Sync % is below threshold
            
        if (($LocalSpeedVaultEnabled -eq 1) -and ($LocalSpeedVaultSynchronizationStatus -ne "synchronized")) { 
            if ( ($LocalSpeedVaultSynchronizationStatus.replace("%",""))/1 -lt $SynchThreshold ){Write-Warning "LSV sync is below $SynchThreshold%"}
        }  ## Warn if Sync % is below threshold

        # Warn if LSV Enabled = False & LSV Path != ""
        
        if ($ThresholdValue) {
            Write-Output "`n####### Start Threshold Values #######"
            Write-Output "-- LSV Enabled" $Global:LSVEnabledOutTxt
            Write-Output "-- LSV Sync" $Global:LSVSyncStatusOutTxt
            Write-Output "-- Cloud Sync" $Global:CloudSyncStatusOutTxt
            Write-Output "-- Selected GB" $Global:TotalSelectedGBTxt
            Write-Output "-- Used GB" $Global:TotalUsedGBTxt
            Write-Output "`n####### End Threshold Values #######"
        }
    }

} ## Get LocalSpeedVaultStatus

#Get-LSVstatus

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
    $Global:GlobalDataSourcesOutTxt = $Datasources -join ","
    Write-Output "`n[Configured Datasources]"
    Write-Output "Data Sources            : $Global:GlobalDataSourcesOutTxt"

    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- DataSources" $Global:GlobalDataSourcesOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
        

} ## Get Last Error per Active Data Source from ClientTool.exe

#Get-Datasources

Function Get-StatusReport ($SuccessHours) {

    $script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $ReplaceStatus = @{
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
    } ## Replace numeric status codes with description and Ascii char.  
    
    $ReplaceArray = @(
        @('0','o'),
        @('1','>'),
        @('2','-'),
        @('3','x'),
        @('4','?'),
        @('5','+'),
        @('6','&'),
        @('7','!'),
        @('8','#'),
        @('9','%')
    ) ## Replace numeric status codes in ColorBar with Ascii char.  

    $Device                                             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text" 
    $PartnerName                                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
    $OsVersion                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text"
    $ClientVersion                                      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ClientVersion")."#text"
    $MachineName                                        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text"
    $TimeStamp                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $TimeZone                                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeZone")."#text"
    $ProfileId                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileId")."#text"
    $ProfileVersion                                     = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileVersion")."#text"
    $StatusCode                                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
  
    $Global:CoveDeviceNameOutTxt = $Device
    $Global:MachineNameOutTxt = $MachineName    
    $Global:PartnerNameOutTxt = $PartnerName    

    Write-Output "`n[Device]"
    Write-Output "Device Name             : $Device"
    Write-Output "Machine Name            : $MachineName"
    Write-Output "TimeStamp (UTC)         : $(Convert-UnixTimeToDateTime $timestamp)"
    Write-Output "Timezone                : $TimeZone"  
    Write-Output "ProfileId               : $ProfileId"
    Write-Output "ProfileVersion          : $ProfileVersion"
    #"Status Code             : $StatusCode"
   

    #PluginVssSystemState
    $Global:PluginDataSourceOutTxt = "SystemState"
    $TimeStamp                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $StatusCode                                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
    $PluginVssSystemStateColorBar                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-ColorBar")."#text"

    if ($PluginVssSystemStateColorBar) {
        $ReplaceArray | ForEach-Object {$PluginVssSystemStateColorBar = $PluginVssSystemStateColorBar -replace $_[0],$_[1]}
        $PluginVssSystemStatecolorbar = ($PluginVssSystemStateColorBar[-1..-($PluginVssSystemStatecolorbar.length)] -join "")   #Reverse colorbar string

        $PluginVssSystemStateLastSessionStatus                  = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionStatus")."#text"]
        $PluginVssSystemStateLastSessionTimestamp               = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSessionTimestamp")."#text")

        $PluginVssSystemStateLastSuccessfulSessionStatus        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionStatus")."#text"
        if ($null -eq $PluginVssSystemStateLastSuccessfulSessionStatus ) {$PluginVssSystemStateLastSuccessfulSessionStatus = "Never"} else {
        $PluginVssSystemStateLastSuccessfulSessionStatus        = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionStatus")."#text"] }

        $script:PluginVssSystemStateLastSuccessfulSessionTimestamp     = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastSuccessfulSessionTimestamp")."#text")
        $PluginVssSystemStateLastCompletedSessionStatus         = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginVssSystemState-LastCompletedSessionStatus")."#text"]
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
        
        Write-Output "`n[SystemState Session]"
        $Global:TimeStampUTCOutTxt = (Convert-UnixTimeToDateTime $timestamp)
        Write-Output "TimeStamp (UTC)         : $Global:TimeStampUTCOutTxt"
        Write-Output "Status Code             : $StatusCode"
        Write-Output "28-Day Status           : $PluginVssSystemStateColorBar"
        Write-Output "Retention               : $PluginVssSystemStateRetention $RetentionUnits"   
        Write-Output "Session Duration        : $($PluginVssSystemStateSessionDuration/60) Mins"                                           
        Write-Output "Last Session (UTC)      : $PluginVssSystemStateLastSessionTimestamp $PluginVssSystemStateLastSessionStatus "
        Write-Output "Last Success (UTC)      : $PluginVssSystemStateLastSuccessfulSessionTimestamp $PluginVssSystemStateLastSuccessfulSessionStatus"
        Write-Output "Last Complete (UTC)     : $PluginVssSystemStateLastCompletedSessionTimestamp $PluginVssSystemStateLastCompletedSessionStatus"
        Write-Output "Last Error Count        : $PluginVssSystemStateLastSessionErrorsCount"

        if ($PluginVssSystemStateLastSuccessfulSessionTimestamp) { 
            $CurrentTime = Get-Date
            $TimeSinceLast = New-TimeSpan -Start $PluginVssSystemStateLastSuccessfulSessionTimestamp -End $CurrentTime.ToUniversalTime()
            $Global:HoursSinceLast = $TimeSinceLast.Totalhours | rnd '' 2
            if ($Global:HoursSinceLast -le $SuccessHours) {
                Write-Output "Last Success (HRS) Ago  : $Global:HoursSinceLast"
            }elseif ($Global:HoursSinceLast -gt $SuccessHours) {
                Write-Warning "Last Success (HRS) Ago : $Global:HoursSinceLast"
            }
        }
            
        if($statuscode -notlike "Backup, System state*") {
            Write-Output "`n[SystemState Size]"
            Write-Output "Last Sel Size           : $($PluginVssSystemStateLastSessionSelectedSize | RND GB 5)"
            Write-Output "Last Proc Size          : $($PluginVssSystemStateLastSessionProcessedSize | RND GB 5)"
            Write-Output "Last Sent Size          : $($PluginVssSystemStateLastSessionSentSize | RND MB 5)" 
            if ($PluginVssSystemStateLastSessionSelectedSize -gt 0) {
                Write-Output "Last Proc/Sel Size%     : $(($PluginVssSystemStateLastSessionProcessedSize/$PluginVssSystemStateLastSessionSelectedSize)*100)"   
                Write-Output "Last Sent/Sel Size%     : $(($PluginVssSystemStateLastSessionSentSize/$PluginVssSystemStateLastSessionSelectedSize)*100)"
            }
            
            if ($PluginVssSystemStatePreRecentSessionSelectedSize -gt 0) {
                Write-Output "`n[SystemState Size Change]"
                Write-Output "Prior Sel Size          : $($PluginVssSystemStatePreRecentSessionSelectedSize | RND GB 5)"
                Write-Output "Sel Change Size         : $(($PluginVssSystemStateLastSessionSelectedSize-$PluginVssSystemStatePreRecentSessionSelectedSize) | RND MB 5)"
                Write-Output "Sel Change Size %       : $(([Math]::Abs((($PluginVssSystemStateLastSessionSelectedSize-$PluginVssSystemStatePreRecentSessionSelectedSize)/$PluginVssSystemStatePreRecentSessionSelectedSize)*100)) | rnd '' 4)"  
            }
              
            Write-Output "`n[SystemState Count]"
            Write-Output "Last Sel Count          : $PluginVssSystemStateLastSessionSelectedCount"
            Write-Output "Last Proc Count         : $PluginVssSystemStateLastSessionProcessedCount"

                if ($PluginVssSystemStatePreRecentSessionSelectedCount) {  
                    Write-Output "`n[SystemState Count Change]"
                    Write-Output "Prior Sel Count         : $PluginVssSystemStatePreRecentSessionSelectedCount"
                    Write-Output "Sel Change Count        : $($PluginVssSystemStateLastSessionSelectedCount-$PluginVssSystemStatePreRecentSessionSelectedCount)" 
                    Write-Output "Sel Change Count %      : $(([Math]::Abs((($PluginVssSystemStateLastSessionSelectedCount-$PluginVssSystemStatePreRecentSessionSelectedCount)/$PluginVssSystemStatePreRecentSessionSelectedCount)*100)) | rnd '' 4) "     }
            }
            #Write-Output "Protected Size          : $($PluginVssSystemStateProtectedSize | RND GB 5)"                 
            #Write-Output "License Count           : $PluginVssSystemStateLastSessionLicenceItemsCount"
        }
    
        if ($ThresholdValue) {
            Write-Output "`n####### Start Threshold Values #######"
            Write-Output "-- Device" $Global:CoveDeviceNameOutTxt
            Write-Output "-- Machine" $Global:MachineNameOutTxt
            Write-Output "-- Partner" $Global:PartnerNameOutTxt 
            Write-Output "-- DataSource" $Global:PluginDataSourceOutTxt
            Write-Output "-- Timestamp (UTC)`n$Global:TimeStampUTCOutTxt"
            Write-Output "-- Hours since last success" $Global:HoursSinceLast
            Write-Output "-- Error Path" $Global:ErrorPathOutTxt
            Write-Output "`n####### End Threshold Values #######"
        }
    
}

get-statusReport -SuccessHours 24

Function Get-BackupErrors ([string]$datasource="SystemState",[int]$Limit=5){

    # Datasource Possible values are BareMetalRestore, Exchange, SystemState, MySql,NetworkShares, Oracle, SystemState, VMware, VirtualDisasterRecovery, VssHyperV, VssMsSql or VssSharePoint

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

        if ($Null -eq ($FunctionalProcess)) {
             "Backup Manager Not Running" 
            }
            else {  
                try { # Command(s) to try
                    $ErrorActionPreference = 'Stop'

                    & $clienttool -machine-readable control.session.list -datasource $DataSource > "C:\ProgramData\MXB\Backup Manager\$DataSource.Sessions.tsv"
                    $sessions = Import-Csv -Delimiter "`t" -Path "C:\ProgramData\MXB\Backup Manager\$DataSource.Sessions.tsv"

                    #$lastsession = ($sessions | Where-object {($_.type -eq "Backup") -and ($_.State -ne "Skipped") -and ($_.State -ne "InProcess")})[0] ## Inprogress does not clear last error
                    $lastsession = ($sessions | Where-object {($_.type -eq "Backup") -and ($_.State -ne "Skipped")})[0] ## Inprogress does clear last error

                    & $clienttool -machine-readable control.session.error.list -datasource $DataSource -limit $limit -time $lastsession.start > "C:\ProgramData\MXB\Backup Manager\$DataSource.Errors.tsv"
                    
                    $Script:sessionerrors = Import-Csv -Delimiter "`t" -Path "C:\ProgramData\MXB\Backup Manager\$DataSource.Errors.tsv"

                    If ($Script:sessionerrors -ne $null) {
                        $Global:ErrorDateTimeOutTxt = $Script:sessionerrors[-1].datetime
                        $Global:ErrorContentOutTxt = $Script:sessionerrors[-1].content 
                        $Global:ErrorPathOutTxt = $Script:sessionerrors[-1].path

                        Write-Warning "[$datasource] Errors Found"
                        $Script:sessionerrors | Select-Object Datetime,Content,Path | Sort-Object -Descending Datetime
                    }elseif ($Script:sessionerrors -eq $null) {
                        $Global:ErrorDateTimeOutTxt = "None"
                        $Global:ErrorContentOutTxt = "None"
                        $Global:ErrorPathOutTxt = "None" 

                        Write-Output "`n[$datasource Errors] $Global:ErrorContentOutTxt`n"
                    } 
                        
                }
                catch{ # What to do with terminating errors
                }
                if ($ThresholdValue) {
                    Write-Output "`n####### Start Threshold Values #######"
                    Write-Output "-- Error Date" $Global:ErrorDateTimeOutTxt
                    Write-Output "-- Error Msg"  $Global:ErrorContentOutTxt
                    Write-Output "-- Error Path" $Global:ErrorPathOutTxt
                    Write-Output "`n####### End Threshold Values #######"
                }
            }


} ## Get Last Error per Active Data Source from ClientTool.exe

Get-BackupErrors SystemState 5




