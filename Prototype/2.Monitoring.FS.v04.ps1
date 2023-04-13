<# ----- About: ----
    # N-able | Cove Data Protection Monitor 2 - FileSystem
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
Clear-Host
#Requires -Version 5.1 -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection - FileSystem Monitor"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

#region Functions
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
} ## Convert epoch time to date time 
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
        $Global:BackupServiceOutVal = 0
        $Global:BackupServiceOutTxt = $BackupService.status 
        Write-output  "Service Status          : $BackupServiceOutTxt"
        $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
    }elseif (($BackupService.status -ne "Running") -and ($null -ne $BackupService.status )){
        $Global:BackupServiceOutVal = 2
        $Global:BackupServiceOutTxt = $BackupService.status
        Write-warning "Service Status    : $BackupServiceOutTxt"
        break
    }elseif ($null -eq $BackupService.status ){
        $Global:BackupServiceOutVal = 1
        $Global:BackupServiceOutTxt = "Not Installed"
        Write-warning "Service Status    : $BackupServiceOutTxt"
        break
    }

} ## Get Backup Service \ Process Status

$BackupServiceName = "Backup Service Controller"
Get-BackupState
Function Get-StatusReport ($InputSuccessHours) {

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
    Write-Output "Partner Name            : $Global:PartnerNameOutTxt"
    Write-Output "TimeStamp (UTC)         : $(Convert-UnixTimeToDateTime $timestamp)"
    Write-Output "Timezone                : $TimeZone"  
    Write-Output "ProfileId               : $ProfileId"
    Write-Output "ProfileVersion          : $ProfileVersion"
    #"Status Code             : $StatusCode"
   

    #PluginFileSystem
    $Global:PluginDataSourceOutTxt = "FileSystem"
    $TimeStamp                                          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
    $StatusCode                                         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
    $PluginFileSystemColorBar                           = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-ColorBar")."#text"

    if ($PluginFileSystemColorBar) {
        $ReplaceArray | ForEach-Object {$PluginFileSystemColorBar = $PluginFileSystemColorBar -replace $_[0],$_[1]}
        $pluginfilesystemcolorbar = ($PluginFileSystemColorBar[-1..-($pluginfilesystemcolorbar.length)] -join "")   #Reverse colorbar string

        $PluginFileSystemLastSessionStatus                  = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionStatus")."#text"]
        $PluginFileSystemLastSessionTimestamp               = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSessionTimestamp")."#text")

        $PluginFileSystemLastSuccessfulSessionStatus        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSuccessfulSessionStatus")."#text"
        if ($null -eq $PluginFileSystemLastSuccessfulSessionStatus ) {$PluginFileSystemLastSuccessfulSessionStatus = "Never"} else {
        $PluginFileSystemLastSuccessfulSessionStatus        = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSuccessfulSessionStatus")."#text"] }

        $script:PluginFileSystemLastSuccessfulSessionTimestamp     = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastSuccessfulSessionTimestamp")."#text")
        $PluginFileSystemLastCompletedSessionStatus         = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PluginFileSystem-LastCompletedSessionStatus")."#text"]
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
        
        Write-Output "`n[FileSystem Session]"
        $Global:TimeStampUTCOutTxt = (Convert-UnixTimeToDateTime $timestamp)
        Write-Output "TimeStamp (UTC)         : $Global:TimeStampUTCOutTxt"
        Write-Output "Status Code             : $StatusCode"
        Write-Output "28-Day Status           : $PluginFileSystemColorBar"
        Write-Output "Retention               : $PluginFileSystemRetention $RetentionUnits"   
        Write-Output "Session Duration        : $($PluginFileSystemSessionDuration/60) Mins"                                           
        Write-Output "Last Session (UTC)      : $PluginFileSystemLastSessionTimestamp $PluginFileSystemLastSessionStatus "
        Write-Output "Last Success (UTC)      : $PluginFileSystemLastSuccessfulSessionTimestamp $PluginFileSystemLastSuccessfulSessionStatus"
        Write-Output "Last Complete (UTC)     : $PluginFileSystemLastCompletedSessionTimestamp $PluginFileSystemLastCompletedSessionStatus"
        Write-Output "Last Error Count        : $PluginFileSystemLastSessionErrorsCount"

        if ($PluginFileSystemLastSuccessfulSessionTimestamp) { 
            $CurrentTime = Get-Date
            $TimeSinceLast = New-TimeSpan -Start $PluginFileSystemLastSuccessfulSessionTimestamp -End $CurrentTime.ToUniversalTime()
            $Global:HoursSinceLastOutVal = $TimeSinceLast.Totalhours | rnd '' 2
            if ($Global:HoursSinceLastOutVal -le $InputSuccessHours) {
                Write-Output "Last Success (HRS) Ago  : $Global:HoursSinceLastOutVal"
            }elseif ($Global:HoursSinceLastOutVal -gt $SuccessHours) {
                Write-Warning "Last Success (HRS) Ago : $Global:HoursSinceLastOutVal"
            }
        }
            
        if($statuscode -notlike "Backup, Files and folders*") { 
            Write-Output "`n[FileSystem Size]"
            Write-Output "Last Sel Size           : $($PluginFileSystemLastSessionSelectedSize | RND GB 5)"
            Write-Output "Last Proc Size          : $($PluginFileSystemLastSessionProcessedSize | RND GB 5)"
            Write-Output "Last Sent Size          : $($PluginFileSystemLastSessionSentSize | RND GB 5)"

            if ($PluginFileSystemLastSessionSelectedSize -gt 0) {
                Write-Output "Last Proc/Sel Size%     : $((($PluginFileSystemLastSessionProcessedSize/$PluginFileSystemLastSessionSelectedSize)*100) | rnd '' 4)"   
                Write-Output "Last Sent/Sel Size%     : $((($PluginFileSystemLastSessionSentSize/$PluginFileSystemLastSessionSelectedSize)*100) | rnd '' 4)"
            } ## If statement prevents divide by zero errors if there is no Last Session data yet
            
            if ($PluginFileSystemPreRecentSessionSelectedSize -gt 0) {
                Write-Output "`n[FileSystem Size Change]"
                Write-Output "Prior Sel Size          : $($PluginFileSystemPreRecentSessionSelectedSize | RND GB 5)"
                Write-Output "Last Sel Size           : $($PluginFileSystemLastSessionSelectedSize | RND GB 5)"
                Write-Output "Sel Change Size         : $(($PluginFileSystemLastSessionSelectedSize-$PluginFileSystemPreRecentSessionSelectedSize) | RND GB 5)"
                if (($PluginFileSystemLastSessionSelectedSize-$PluginFileSystemPreRecentSessionSelectedSize) -gt 0) {$Global:SizeChangeOutTxt = "% Increase (+)" }else{$Global:SizeChangOutTxt = "% Decrease (-)" }
                Write-Output "Sel Change Size %       : $(([Math]::Abs((($PluginFileSystemLastSessionSelectedSize-$PluginFileSystemPreRecentSessionSelectedSize)/$PluginFileSystemPreRecentSessionSelectedSize)*100)) | rnd '' 4)$Global:SizeChangeOutTxt"  
            } ## If statement prevents divide by zero errors if there is no Pre Recent Session data yet
              
            Write-Output "`n[FileSystem Count]"
            Write-Output "Last Sel Count          : $PluginFileSystemLastSessionSelectedCount"
            Write-Output "Last Proc Count         : $PluginFileSystemLastSessionProcessedCount"
            Write-Output "Last Proc/Sel Count %   : $(($PluginFileSystemLastSessionProcessedCount/$PluginFileSystemLastSessionSelectedCount)*100)"

            if ($PluginFileSystemPreRecentSessionSelectedCount) {  
                Write-Output "`n[FileSystem Count Change]"
                Write-Output "Prior Sel Count         : $PluginFileSystemPreRecentSessionSelectedCount"
                Write-Output "Last Sel Count          : $PluginFileSystemLastSessionSelectedCount"
                Write-Output "Sel Change Count #      : $($PluginFileSystemLastSessionSelectedCount-$PluginFileSystemPreRecentSessionSelectedCount)" 
                if (($PluginFileSystemLastSessionSelectedCount-$PluginFileSystemPreRecentSessionSelectedCount) -gt 0) {$Global:CountChangeTXT = "% Increase (+)" }else{$Global:CountChangeTXT = "% Decrease (-)" }
                Write-Output "Sel Change Count %      : $(([Math]::Abs((($PluginFileSystemLastSessionSelectedCount-$PluginFileSystemPreRecentSessionSelectedCount)/$PluginFileSystemPreRecentSessionSelectedCount)*100)) | rnd '' 4)$Global:CountChangeTXT"

            }
        } ## If statement prevents update in the middle of a running job
    }
    if ($ThresholdValue) {
        Write-Output "`n####### Start Threshold Values #######"
        Write-Output "-- Device" $Global:CoveDeviceNameOutTxt
        Write-Output "-- Machine" $Global:MachineNameOutTxt
        Write-Output "-- Partner" $Global:PartnerNameOutTxt 
        Write-Output "-- DataSource" $Global:PluginDataSourceOutTxt
        Write-Output "-- Timestamp (UTC)`n$Global:TimeStampUTCOutTxt"
        Write-Output "-- Hours since last success" $Global:HoursSinceLastOutVal
        Write-Output "-- Error Path" $Global:ErrorPathOutTxt
        Write-Output "`n####### End Threshold Values #######"
    }
    
} ## Get Data from StatusReport.xml

Get-StatusReport -InputSuccessHours 24
Function Get-BackupErrors ([string]$datasource="FileSystem",[int]$Limit=5){

    # Datasource Possible values are BareMetalRestore, Exchange, FileSystem, MySql,NetworkShares, Oracle, SystemState, VMware, VirtualDisasterRecovery, VssHyperV, VssMsSql or VssSharePoint

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
            if ($Null -eq ($FunctionalProcess)) {
                Write-Output "Backup Manager Not Running" 
            }else{  
                Do {
                    $BackupStatus = & $clienttool -machine-readable control.status.get
                    $StatusValue = @("Suspended")
                    if ($StatusValue -contains $BackupStatus) { 
                        Start-Sleep -seconds 30
                        $retrycounter ++
                    }else{   
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
                }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 
                write-output $retrycounter $BackupStatus
            }


} ## Get Last Error per Active Data Source from ClientTool.exe

Get-BackupErrors FileSystem 5

#endregion Functions

if ($ThresholdValue) {
    Write-Output "`n####### Start Threshold Values #######"
    Write-Output "-- Device" $Global:CoveDeviceNameOutTxt
    Write-Output "-- Machine" $Global:MachineNameOutTxt
    Write-Output "-- Partner" $Global:PartnerNameOutTxt 
    Write-Output "-- DataSource" $Global:PluginDataSourceOutTxt
    Write-Output "-- Timestamp (UTC)`n$Global:TimeStampUTCOutTxt"
    Write-Output "-- Hours since last success" $Global:HoursSinceLastOutVal
    
    Write-Output "-- Error Path" $Global:ErrorPathOutTxt
    Write-Output "-- Error Date" $Global:ErrorDateTimeOutTxt
    Write-Output "-- Error Msg"  $Global:ErrorContentOutTxt
    Write-Output "-- Error Path" $Global:ErrorPathOutTxt
    
    Write-Output "`n####### End Threshold Values #######"
}
