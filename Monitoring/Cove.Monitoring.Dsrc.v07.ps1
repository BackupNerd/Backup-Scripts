<# ----- About: ----
    # N-able | Cove Data Protection | Monitor DataSource
    # Revision v07 - 2023-05-10
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
    #  Get Cove data sources / statistics
    #  Get Cove data source errors
     






 # -----------------------------------------------------------#>

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False, Position=0)] 
        [ValidateSet("FileSystem","SystemState","VMware","VssHyperV","VssMsSql","Exchange","MySql","NetworkShares","VssSharePoint","Oracle","Total")] [String]$datasource = "FileSystem",
        [Parameter(Mandatory=$False)] [decimal]$SuccessHoursInVal = 24,     ## Fail check after this number of hours since last success,           
        [Parameter(Mandatory=$False)] [switch]$DebugDetail,                 ## Set True and Disconnect NIC to test Debug Data scenarios,    
        [Parameter(Mandatory=$False)] [switch]$ThresholdOutput = $true      ## Set True to Output Threshold Values at end of script
    )        
Clear-Host
#Requires -Version 5.1 -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection - $datasource Monitor"
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

$Script:BackupServiceOutVal = -1
$Script:BackupServiceOutTxt = "Not Installed"
$Script:CoveDeviceNameOutTxt = "Undefined"
$Script:MachineNameOutTxt = "Undefined"
$Script:CustomerNameOutTxt = "Undefined"
$Script:OsVersionOutTxt = "Undefined"
$Script:ClientVerOutTxt = "Undefined"
$Script:PluginDataSourceOutTxt = "Undefined"
$Script:TimeStampUTCOutTxt = "Undefined"
$Script:TimeStampOutVal = 0 
$Script:TimeZoneOutVal = 0 
$Script:ProfileNameOutTxt = "Undefined"
$Script:ProfileIdOutVal = 0
$Script:ProfileVersionOutVal = 0
$Script:StatusCodeOutTxt = "Undefined"
$Script:RetentionUnitsOutTxt = "Undefined"
$Script:PluginRetentionOutVal = 0
$Script:PluginColorBarOutTxt = "Undefined"
$Script:LastSessionTimeUTCOutTxt = "Never"
$Script:LastSessionStatusOutTxt = "Never"
$Script:LastSuccessTimeUTCOutTxt = "Never"
$Script:LastSuccessStatusOutTxt = "Never"
$Script:LastCompleteTimeUTCOutTxt = "Never"
$Script:LastCompleteStatusOutTxt = "Never"
$Script:PluginSessionDurationHrsOutVal = -1
$Script:SuccessHoursOutVal = $SuccessHoursInVal
$Script:HoursSinceLastOutVal = -1
$Script:LastErrorCountOutVal = 0
$Script:ErrorDateTimeOutTxt = "None"
$Script:ErrorPathOutTxt = "None"
$Script:ErrorContentOutTxt = "None"
$Script:PriorSessionSelSizeOutVal = 0
$Script:LastSessionSelSizeOutVal = 0
$Script:LastSessionProcSizeOutVal = 0
$Script:LastSessionSentSizeOutVal = 0
$Script:SentPercentOutVal = 0
$Script:ProcPercentOutVal = 0
$Script:PriorSessionSelCountOutVal = 0
$Script:LastSessionSelCountOutVal = 0
$Script:SelChangeCountOutVal = 0
$Script:SelChangeCountPercentOutVal = 0
$Script:LastSessionProcCountOutVal = 0
$Script:CountChangeOutTxt = "None"
$Script:SelChangeSizeOutVal = 0
$Script:SelChangeSizePercentOutVal = 0
$Script:SizeChangeOutTxt = "None"

Function Get-FixedVolumes {
    $Volumes = Get-Volume

    $FixedVolumes = $Volumes | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.DriveLetter)}

    $FixedVolumes1 = $FixedVolumes | Select-Object @{l='Letter';e={$_.DriveLetter + ":"}},@{l='Type';e={$_.DriveType}},@{l='TotalGB';e={([Math]::Round(($_.Size/1GB),2))}},@{l='FreeGB';e={([Math]::Round(($_.SizeRemaining/1GB),2))}},@{l='UsedGB';e={([Math]::Round((($_.Size - $_.SizeRemaining)/1GB),2))}}
    
    $FixedVolumes2 = $FixedVolumes1 | Select-Object Letter,Type,UsedGB,@{l='String';e={$_.Letter + ' ' + $_.UsedGB +' GB'}}

    $FixedVolumeString = $FixedVolumes2.String -join " | "
    Write-output "`n[Detected DataSources]"
    Write-output "Fixed Volumes             : $FixedVolumeString"
}

Function Get-Filters {    
    
    & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.filter.list | out-file C:\programdata\mxb\filters.csv

    $Filters = import-csv -path C:\programdata\mxb\filters.csv -Header value

    if ($Filters) {
             
        $FilterString = $Filters.value.replace("\","\\") -join " | "
        Write-output "`n[$Datasource Filters]"
        Write-output "Runtime Exclusion Filters : $($FilterString.replace("\\","\"))"
    }

}

Function Get-BackupState {

    $Script:BackupService = get-service -name $BackupServiceName -ea SilentlyContinue
    Write-output "`n[Backup Service Controller]"
    if ($BackupService.status -eq "Running"){
        $Script:BackupServiceOutVal = 1
        $Script:BackupServiceOutTxt = $($BackupService.status)
        Write-output  "Service Status            : $BackupServiceOutTxt"
        $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
    }elseif (($BackupService.status -ne "Running") -and ($null -ne $BackupService.status )){
        $Script:BackupServiceOutVal = 0
        $Script:BackupServiceOutTxt = $($BackupService.status)
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }elseif ($null -eq $BackupService.status ){
        $Script:BackupServiceOutVal = -1
        $Script:BackupServiceOutTxt = "Not Installed"
        Write-warning "Service Status    : $BackupServiceOutTxt"
    }

} ## Get Backup Service \ Process Status

$BackupServiceName = "Backup Service Controller"


Function Get-BackupErrors ($datasource,[int]$Limit=5){

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
        
                            If ($null -ne $Script:sessionerrors) {
                                $Script:ErrorDateTimeOutTxt = $Script:sessionerrors[-1].datetime
                                $Script:ErrorContentOutTxt = $Script:sessionerrors[-1].content 
                                $Script:ErrorPathOutTxt = $Script:sessionerrors[-1].path
        
                                Write-Warning "[$datasource] Errors Found"
                                $Script:sessionerrors | Select-Object Datetime,Content,Path | Sort-Object -Descending Datetime
                            }elseif ($null -eq $Script:sessionerrors) {
                                $Script:ErrorDateTimeOutTxt = "None"
                                $Script:ErrorContentOutTxt = "None"
                                $Script:ErrorPathOutTxt = "None" 
        
                                Write-Output "`n[$datasource Errors] $Script:ErrorContentOutTxt`n"
                            } 
                        }
                        catch{ # What to do with terminating errors
                        }
                        if ($ThresholdValue) {
                            Write-Output "`n####### Start Threshold Values #######"
                            Write-Output " Last Error Count       | $Script:LastErrorCountOutVal"
                            Write-Output " Error Date             | $Script:ErrorDateTimeOutTxt"
                            Write-Output " Error Msg              | $Script:ErrorContentOutTxt"
                            Write-Output " Error Path             | $Script:ErrorPathOutTxt"
                            Write-Output "`n####### End Threshold Values #######"
                        }
                    }
                }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 
                if ($DebugDetail) {write-output $retrycounter $BackupStatus}
            }

} ## Get Last Error per Active Data Source from ClientTool.exe
Function Get-StatusReport {

    $script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    if (Test-Path $script:StatusReportxml ) {    

        $ReplaceDatasourceforStatus = @{
            'FileSystem'    = 'Backup, Files and folders*'
            'SystemState'   = 'Backup, System state*'
            'VMware'        = 'Backup, VMware*'
            'VssHyperV'     = 'Backup, HyperV*'
            'VssMsSql'      = 'Backup, MS SQL*'
            'Exchange'      = 'Backup, Exchange*'
            'MySql'         = 'Backup, MySQL*'
            'NetworkShares' = 'Backup, Network*'
            'VssSharePoint' = 'Backup, SharePoint*'
            'Total'         = 'Backup,*'
        } ## Replace DataSource with Plugin names

        $ReplaceDatasourceforXMLLookup = @{
            'FileSystem'    = 'FileSystem'
            'SystemState'   = 'VssSystemState'
            'VMware'        = 'VMWare'
            'VssHyperV'     = 'VssHyperV' 
            'VssMsSql'      = 'VssMsSql' 
            'Exchange'      = 'Exchange'  
            'MySql'         = 'MySql'  
            'NetworkShares' = 'NetworkShares'
            'VssSharePoint' = 'VssSharePoint'
            'Total'         = 'Total'
        } ## Replace DataSource with Plugin names for Lookup

        $ReplaceStatus = @{
            'Never' = 'NoBackup (o)'
            '' = 'NoBackup (o)'
            '0' = 'NoBackup (o)'
            '1' = 'InProcess (>)'
            '2' = 'Failed (-)'
            '3' = 'Aborted (x)'
            '4' = 'Unknown (?)'
            '5' = 'Completed (+)'
            '6' = 'Interrupted (&)'
            '7' = 'NotStarted (!)'
            '8' = 'CompletedWithErrors (#)'
            '9' = 'InProgressWithFaults (%)'
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
        
        $Script:PluginDataSourceOutTxt = $datasource
        $XMLDataSource = $ReplaceDatasourceforXMLLookup[$datasource]

        # Get Data from XML
        $Script:CoveDeviceNameOutTxt    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Account")."#text"
        $Script:MachineNameOutTxt       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//MachineName")."#text" 
        $Script:CustomerNameOutTxt      = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
        $Script:OsVersionOutTxt         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//OsVersion")."#text"
        $Script:ClientVerOutTxt         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ClientVersion")."#text"

        $Script:TimeZoneOutVal          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeZone")."#text"
        $Script:TimeStampOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//TimeStamp")."#text"
        $Script:TimeStampUTCOutTxt      = (Convert-UnixTimeToDateTime $Script:TimeStampOutVal)
        $Script:ProfileNameOutTxt       = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileName")."#text"
        if ($null -eq $Script:ProfileNameOutTxt) { $Script:ProfileNameOutTxt = "Undefined"} 
        $Script:ProfileIdOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileId")."#text"
        $Script:ProfileVersionOutVal    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//ProfileVersion")."#text"
        $Script:StatusCodeOutTxt        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//StatusCode")."#text"
        $Script:RetentionUnitsOutTxt    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//RetentionUnits")."#text"

        Write-Output "`n[Device]"
        Write-Output "Device Name               : $Script:CoveDeviceNameOutTxt"
        Write-Output "Machine Name              : $Script:MachineNameOutTxt"
        Write-Output "Customer Name             : $Script:CustomerNameOutTxt"
        Write-Output "TimeZone                  : $Script:TimeZoneOutVal"
        Write-Output "TimeStamp(UTC)            : $Script:TimeStampUTCOutTxt"
        Write-Output "Profile Name              : $Script:ProfileNameOutTxt"  
        Write-Output "Profile Id                : $Script:ProfileIdOutVal"
        Write-Output "Profile Ver               : $Script:ProfileVersionOutVal"
 
        $PluginColorBar                 = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-ColorBar")."#text"

        if ($PluginColorBar) {
            $ReplaceArray | ForEach-Object {$PluginColorBar = $PluginColorBar -replace $_[0],$_[1]}
            $Script:PluginColorBarOutTxt = ($PluginColorBar[-1..-($plugincolorbar.length)] -join "")   #Reverse colorbar string order

            $Script:LastSessionStatusOutTxt           = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionStatus")."#text"]
            $Script:LastSessionTimeUTCOutTxt          = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionTimestamp")."#text")

            $PluginLastSuccessfulSessionStatus        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSuccessfulSessionStatus")."#text"
            if ($null -eq $PluginLastSuccessfulSessionStatus ) {$Script:LastSuccessStatusOutTxt = "Never"} else {

                $Script:LastSuccessStatusOutTxt           = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSuccessfulSessionStatus")."#text"]
                $Script:LastSuccessTimeUTCOutTxt          = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSuccessfulSessionTimestamp")."#text")
                $Script:LastCompleteStatusOutTxt          = $ReplaceStatus[([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastCompletedSessionStatus")."#text"]
                $Script:LastCompleteTimeUTCOutTxt         = Convert-UnixTimeToDateTime(([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastCompletedSessionTimestamp")."#text")



                $Script:LastErrorCountOutVal              = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionErrorsCount")."#text"
                $Script:PluginSessionDurationHrsOutVal    = (([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-SessionDuration")."#text"/60/60)
                $Script:PluginRetentionOutVal             = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-Retention")."#text"
                
                Write-Output "`n[$Datasource Session]"
                Write-Output "JobStatus                 : $Script:StatusCodeOutTxt"
                Write-Output "28-Day Status             : $Script:PluginColorBarOutTxt"
                Write-Output "Retention                 : $Script:PluginRetentionOutVal $Script:RetentionUnitsOutTxt "   
                
                Write-Output "Last Success (UTC)        : $Script:LastSuccessTimeUTCOutTxt $Script:LastSuccessStatusOutTxt"
                Write-Output "Last Complete (UTC)       : $Script:LastCompleteTimeUTCOutTxt $Script:LastCompleteStatusOutTxt"      
                Write-Output "Last Session (UTC)        : $Script:LastSessionTimeUTCOutTxt $Script:LastSessionStatusOutTxt"
                Write-Output "Session Duration (HRS)    : $Script:PluginSessionDurationHrsOutVal"
                Write-Output "Success Check (HRS)       : $SuccessHoursInVal"

                if ($Script:LastSuccessTimeUTCOutTxt -gt 0) { 
                    $CurrentTime = Get-Date
                    $TimeSinceLast = New-TimeSpan -Start $Script:LastSuccessTimeUTCOutTxt -End $CurrentTime.ToUniversalTime()
                    [decimal]$Script:HoursSinceLastOutVal = $TimeSinceLast.Totalhours | rnd '' 2
                    if ($Script:HoursSinceLastOutVal -le $SuccessHoursInVal) {
                        Write-Output "Last Success              : $Script:HoursSinceLastOutVal(HRS) Ago"
                    }elseif ($Script:HoursSinceLastOutVal -gt $SuccessHoursInVal) {
                        Write-Warning "Last Success     : $Script:HoursSinceLastOutVal(HRS) Ago"
                    }
                }
            }
            Write-Output "Last Error Count          : $Script:LastErrorCountOutVal"

            if ($Script:LastErrorCountOutVal -ge 1) {Get-BackupErrors $DataSource 5}
            
            if (($Script:StatusCodeOutTxt -notlike $ReplaceDatasourceforStatus[$datasource]) -and ($Script:StatusCodeOutTxt -notlike "Scanning")) { 

                $Script:PriorSessionSelSizeOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-PreRecentSessionSelectedSize")."#text"
                $Script:LastSessionSelSizeOutVal          = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionSelectedSize")."#text"
                $Script:LastSessionProcSizeOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionProcessedSize")."#text"
                $Script:LastSessionSentSizeOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionSentSize")."#text"
                $Script:PriorSessionSelCountOutVal        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-PreRecentSessionSelectedCount")."#text"
                if ($Script:PriorSessionSelCountOutVal -eq $null) {$Script:PriorSessionSelCountOutVal = 0}
                $Script:LastSessionSelCountOutVal         = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionSelectedCount")."#text"
                $Script:LastSessionProcCountOutVal        = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-LastSessionProcessedCount")."#text"

                Write-Output "`n[$datasource Size]"
                Write-Output "Last Sel Size             : $($Script:LastSessionSelSizeOutVal | RND GB 3)"
                Write-Output "Last Proc Size            : $($Script:LastSessionProcSizeOutVal | RND MB 3)"
                Write-Output "Last Sent Size            : $($Script:LastSessionSentSizeOutVal | RND MB 3)"

                if ($Script:LastSessionSelSizeOutVal -ne 0) {
                    $Script:ProcPercentOutVal = $((($Script:LastSessionProcSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)
                    Write-Output "Last Proc/Sel Size%       : $($Script:ProcPercentOutVal)"
                    #Write-Output "Last Proc/Sel Size%       : $((($Script:LastSessionProcSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)"

                    $Script:SentPercentOutVal = $((($Script:LastSessionSentSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)
                    Write-Output "Last Sent/Sel Size%       : $($Script:SentPercentOutVal)"      
                    #Write-Output "Last Sent/Sel Size%       : $((($Script:LastSessionSentSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)"
                } ## If statement prevents divide by zero errors if there is no Last Session data yet
                
                if ($Script:PriorSessionSelSizeOutVal -ne 0) {
                    Write-Output "`n[$datasource Size Change]"
                    Write-Output "Prior Sel Size            : $($Script:PriorSessionSelSizeOutVal | RND GB 3)"
                    Write-Output "Last Sel Size             : $($Script:LastSessionSelSizeOutVal | RND GB 3)"

                    $Script:SelChangeSizeOutVal = $(($Script:LastSessionSelSizeOutVal-$Script:PriorSessionSelSizeOutVal) | RND MB 3)
                    #Write-Output "Sel Change Size           : $(($Script:LastSessionSelSizeOutVal-$Script:PriorSessionSelSizeOutVal) | RND MB 3)"
                    Write-Output "Sel Change Size           : $Script:SelChangeSizeOutVal"

                    if (($Script:LastSessionSelSizeOutVal-$Script:PriorSessionSelSizeOutVal) -ge 0) {$Script:SizeChangeOutTxt = "% Increase (+)" }else{$Script:SizeChangeOutTxt = "% Decrease (-)" }
                    $Script:SelChangeSizePercentOutVal = $(([Math]::Abs((($Script:LastSessionSelSizeOutVal-$Script:PriorSessionSelSizeOutVal)/$Script:PriorSessionSelSizeOutVal)*100)) | RND '' 3)
                    Write-Output "Sel Change Size %         : $Script:SelChangeSizePercentOutVal$Script:SizeChangeOutTxt"
                } ## If statement prevents divide by zero errors if there is no Pre Recent Session data yet
                
                Write-Output "`n[$datasource Count]"
                Write-Output "Last Sel Count            : $Script:LastSessionSelCountOutVal"
                Write-Output "Last Proc Count           : $Script:LastSessionProcCountOutVal"
                
                if ($Script:LastSessionSelCountOutVal -gt 0) {
                Write-Output "Last Proc/Sel Count %     : $((($Script:LastSessionProcCountOutVal/$Script:LastSessionSelCountOutVal)*100) | RND '' 3)"
                }

                if ($Script:PriorSessionSelCountOutVal -gt 0) {  
                    Write-Output "`n[$datasource Count Change]"
                    Write-Output "Prior Sel Count           : $Script:PriorSessionSelCountOutVal"
                    Write-Output "Last Sel Count            : $Script:LastSessionSelCountOutVal"
                    
                    $Script:SelChangeCountOutVal = ($Script:LastSessionSelCountOutVal-$Script:PriorSessionSelCountOutVal)
                    Write-Output "Sel Change Count #        : $Script:SelChangeCountOutVal"

                    if (($Script:SelChangeCountOutVal) -ge 0) {$Script:CountChangeOutTxt = "% Increase (+)" }else{$Script:CountChangeOutTxt = "% Decrease (-)" }

                    $Script:SelChangeCountPercentOutVal = ((([Math]::Abs((($Script:LastSessionSelCountOutVal-$Script:PriorSessionSelCountOutVal)/$Script:PriorSessionSelCountOutVal)*100))) | rnd '' 3)
                    Write-Output "Sel Change Count %        : $Script:SelChangeCountPercentOutVal$Script:CountChangeOutTxt"
    
                }
            }else{Write-Warning "A Backup session is in progress, some stats are delayed"} ## If statement to prevent update in the middle of a running job
        }else{Write-Warning "No Prior $Datasource Backup has completed"} ## Prevent null values

        }  ## Test if StatusReport.xml file exists
    else{Write-Warning "[$script:StatusReportxml] not found" }
} ## Get Data from StatusReport.xml

#Get-StatusReport -InputSuccessHours 24


Function Get-BackupSelections {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $BackupSelections = & $clienttool -machine-readable control.selection.list | ConvertFrom-String | Select-Object -skip 1 -property @{N='DataSource'; E={$_.P1}},@{N='Type'; E={$_.P2.replace("Inclusive","Inc (+)").replace("Exclusive","Exc (-)")}},@{N='Path'; E={$_.P4}} | where-object {$_.DataSource -eq $datasource}}catch{ Write-Warning "ERROR     : $_" }}

    Write-output "`n[$datasource Backup Selections]"

    if ($BackupSelections) {
        $BackupSelections | Select-object Type,Path | format-table
    }else { Write-Warning "No data selections exist for $DataSource."}
} ## Get Backup Selections from ClientTool.exe

#endregion Functions

Get-BackupState

Switch ($script:datasource) { ## "FileSystem","SystemState","VMware","VssHyperV","VssMsSql","Exchange","MySql","NetworkShares","VssSharePoint","Total"
    'FileSystem' {
        Get-FixedVolumes
        Get-Filters
        }
    'NetworkShares' {
        Get-Filters
        }        
    'MySql' {
        $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Oracle*") -and ($_.name -like "MySQL*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
        Write-output "`n[Detected DataSources]"
        if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
        else{Write-Warning "No $Datasource installation detected"}

        }
    'Oracle' {
        $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Oracle*") -and ($_.name -like "*Oracle*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
        Write-output "`n[Detected DataSources]"
        if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
        else{Write-Warning "No $Datasource installation detected"}

        }
    'VssMsSql' {
        $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "*SQL Server*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
        Write-output "`n[Detected DataSources]"
        if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
        else{Write-Warning "No $Datasource installation detected"}

        }
    'Exchange' {
        $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "*Exchange*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
        Write-output "`n[Detected DataSources]"
        if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
        else{Write-Warning "No $Datasource installation detected"}

        }
    'VssSharePoint' {
        $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "*Sharepoint*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
        Write-output "`n[Detected DataSources]"
       if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
        else{Write-Warning "No $Datasource installation detected"}

        }
    }

    Get-BackupSelections
    Get-StatusReport

if ($ThresholdOutput) {
    Write-Output "`n####### Start Threshold Output Values #######"

    Write-Output " Backup service value     | $Script:BackupServiceOutVal"
    Write-Output " Backup service state     | $Script:BackupServiceOutTxt"
    Write-Output " Device name              | $Script:CoveDeviceNameOutTxt"
    Write-Output " Machine name             | $Script:MachineNameOutTxt"
    Write-Output " Customer name            | $Script:CustomerNameOutTxt"

    Write-Output " Backup client Ver        | $Script:ClientVerOutTxt"
    Write-Output " OS Version               | $Script:OsVersionOutTxt"

    Write-Output " Profile Name             | $Script:ProfileNameOutTxt"
    Write-Output " Profile Id               | $Script:ProfileIdOutVal"
    Write-Output " Profile Ver              | $Script:ProfileVersionOutVal"

    Write-Output " DataSource               | $Script:PluginDataSourceOutTxt"
    Write-Output " Timezone                 | $Script:TimeZoneOutVal"
    Write-Output " Timestamp (UTC)          | $Script:TimeStampUTCOutTxt"
  
    Write-Output " Job Status               | $Script:StatusCodeOutTxt"
    Write-Output " 28-Day Status            | $Script:PluginColorBarOutTxt"
    Write-Output " Retention Units          | $Script:RetentionUnitsOutTxt"
    Write-Output " RetenTion Value          | $Script:PluginRetentionOutVal"
    
    Write-Output " Last Success Status      | $Script:LastSuccessStatusOutTxt"
    Write-Output " Last Success (UTC)       | $Script:LastSuccessTimeUTCOutTxt"
    
    Write-Output " Last Complete Status     | $Script:LastCompleteStatusOutTxt"
    Write-Output " Last Complete (UTC)      | $Script:LastCompleteTimeUTCOutTxt"

    Write-Output " Last Session Status      | $Script:LastSessionStatusOutTxt"
    Write-Output " Last Session (UTC)       | $Script:LastSessionTimeUTCOutTxt"

    Write-Output " Session Duration (HRS)   | $Script:PluginSessionDurationHrsOutVal"
    Write-Output " Success Check (HRS)      | $Script:SuccessHoursOutVal"    
    Write-Output " Hrs Since Last Success   | $Script:HoursSinceLastOutVal"
    
    Write-Output " Last Error Count #       | $Script:LastErrorCountOutVal"
    Write-Output " Error Date               | $Script:ErrorDateTimeOutTxt"
    Write-Output " Error Msg                | $Script:ErrorContentOutTxt"
    Write-Output " Error Path               | $Script:ErrorPathOutTxt"
  
    Write-Output " Last Sel Size            | $($Script:LastSessionSelSizeOutVal | RND GB 3)"
    Write-Output " Last Proc Size           | $($Script:LastSessionProcSizeOutVal | RND MB 3)"
    Write-Output " Last Sent Size           | $($Script:LastSessionSentSizeOutVal | RND MB 3)"
    #Write-Output " Last Proc/Sel Size%      | $((($Script:LastSessionProcSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)"
    Write-Output " Last Proc/Sel Size%      | $Script:ProcPercentOutVal"

    #Write-Output " Last Sent/Sel Size%      | $((($Script:LastSessionSentSizeOutVal/$Script:LastSessionSelSizeOutVal)*100) | RND '' 3)"
    Write-Output " Last Sent/Sel Size%      | $Script:SentPercentOutVal"
    Write-Output " Prior Sel Size           | $($Script:PriorSessionSelSizeOutVal | RND GB 3)"
    Write-Output " Last Sel Size            | $($Script:LastSessionSelSizeOutVal | RND GB 3)"
    Write-Output " Sel Change Size          | $Script:SelChangeSizeOutVal"
    Write-Output " Sel Change Size %        | $Script:SelChangeSizePercentOutVal"
    Write-Output " Sel Change Size Text     | $Script:SizeChangeOutTxt"
    Write-Output " Last Proc Count #        | $Script:LastSessionProcCountOutVal"
    Write-Output " Prior Sel Count #        | $Script:PriorSessionSelCountOutVal"
    Write-Output " Last Sel Count #         | $Script:LastSessionSelCountOutVal"
    Write-Output " Sel Change Count #       | $Script:SelChangeCountOutVal"
    Write-Output " Sel Change Count %       | $Script:SelChangeCountPercentOutVal"
    Write-Output " Sel Change Count Text    | $Script:CountChangeOutTxt"

    Write-Output "`n####### End Threshold Output Values #######"
}
