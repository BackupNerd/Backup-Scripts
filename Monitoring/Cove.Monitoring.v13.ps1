<# ----- About: ----
    # N-able | Cove Data Protection | Monitor
    # Revision v13 - 2023-09-20
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
    #  Get Cove service status / Settings / LSV
    #  Get Cove data sources / statistics / selections
    #  Get Cove data source errors
    #
    #  Suggested use with Connectwise or other Powershell Script Monitor
    #  Triggering on Script Output 'Contains' Keyword 'Warning:'
    #  Warnings force an 'Exit 1001' Code for user with N-able N-sight RMM
    #
 # -----------------------------------------------------------#>

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)][ValidateSet("FileSystem","SystemState","NetworkShares","VssMsSql","Exchange","MySql","VMware","VssHyperV","VssSharePoint","Oracle")] $Datasource,
        [Parameter(Mandatory=$False)] [int]$ServerSuccessHoursInVal = 24,   ## Throw warning if > than this number of hours since last success - Server
        [Parameter(Mandatory=$False)] [int]$WrkStnSuccessHoursInVal = 72,   ## Throw warning if > than this number of hours since last success - Wrkstn
        [Parameter(Mandatory=$False)] [int]$SynchThreshold = 90             ## Throw warning if LSV Sync is < this percentage
        
    )        
Clear-Host      
#Requires -Version 5.1 -RunAsAdministrator
#$ConsoleTitle = "Cove Data Protection - Monitor"
#$host.UI.RawUI.WindowTitle = $ConsoleTitle
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
1.23123123123123123 | RND '' 6                      ## output = 1.231231 
RND KB 2 234234234.234234234234                     ## output = 228744.37 KB
234234234.234234234234 | RND KB 2                   ## output = 228744.37 KB
234234234.234234234234 | RND MB 4                   ## output = 223.3832 MB
234234234.234234234234 | RND GB 1                   ## output = 0.2 GB
234234234.234234234234 | RND                        ## output = 0.22 GB
234234234.234234234234 | RND MB 0                   ## output = 223 MB
1223234234234.234234234234 | RND TB 2               ## output = 1.11 TB
write-output "12312312313.123123123" | RND KB       ## output = 12023742.49 KB
write-output "TEST $(1231231231.2312 | RND GB 2)"   ## output = TEST 1.15 GB
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

Function Get-BackupState {
    
    $script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 
     
    $Script:BackupService = get-service -name "Backup Service Controller" -ea SilentlyContinue


    Write-output "`n[Backup Service Controller]"
    if ($BackupService.status -eq "Running"){
        $Script:BackupServiceOutVal = 1
        $Script:BackupServiceOutTxt = $($BackupService.status)
        Write-output  "Service Status            | $BackupServiceOutTxt"
        $BackupFP = "C:\Program Files\Backup Manager\BackupFP.exe"
        $Script:FunctionalProcess = get-process -name "BackupFP" -ea SilentlyContinue | where-object {$_.path -eq $BackupFP}
    }elseif (($BackupService.status -ne "Running") -and ($null -ne $BackupService.status )){
        $Script:BackupServiceOutVal = 0
        $Script:BackupServiceOutTxt = $($BackupService.status)
        Write-warning "Service Status   | $BackupServiceOutTxt"
        $global:failed = 1
        break
    }elseif (($null -eq $BackupService.status ) -and (Test-Path $script:StatusReportxml)) {
        $Script:BackupServiceOutVal = -2
        $Script:BackupServiceOutTxt = "Previously Installed"
        Write-warning "Service Status    | $BackupServiceOutTxt"
        $global:failed = 1
        Break
    }elseif ($null -eq $BackupService.status ){
        $Script:BackupServiceOutVal = -1
        $Script:BackupServiceOutTxt = "Not Installed"
        Write-output "Service Status    | $BackupServiceOutTxt"
        Break
    
    }

} ## Get Backup Service \ Process Status

Function Get-Datasources {
    $retrycounter = 0
    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
    if ($Null -eq ($FunctionalProcess)) { 
        Write-Warning "Backup Manager Not Running"
        $global:failed = 1 
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
                        $script:Datasources = & $clienttool -machine-readable control.selection.list | ConvertFrom-String | Select-Object -skip 1 -property p1,p2 -unique | ForEach-Object {If ($_.P2 -eq "Inclusive") {Write-Output $_.P1}}
                        $Script:DataSourcesOutTxt = $Datasources -join ", "
                        Write-Output "`n[Configured Datasources]"
                        Write-Output "Data Sources              | $Script:DataSourcesOutTxt" 
                    }
                    catch{ # What to do with terminating errors
                        
                    }
                }
            }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5))
        }

} ## Get Active Data Source from ClientTool.exe

Function Get-FixedVolumes {
    $Volumes = Get-Volume
    $FixedVolumes = $Volumes | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.DriveLetter)}
    $FixedVolumes1 = $FixedVolumes | Select-Object @{l='Letter';e={$_.DriveLetter + ":"}},@{l='Type';e={$_.DriveType}},@{l='TotalGB';e={([Math]::Round(($_.Size/1GB),2))}},@{l='FreeGB';e={([Math]::Round(($_.SizeRemaining/1GB),2))}},@{l='UsedGB';e={([Math]::Round((($_.Size - $_.SizeRemaining)/1GB),2))}}
    
    $FixedVolumes2 = $FixedVolumes1 | Select-Object Letter,Type,UsedGB,@{l='String';e={$_.Letter + ' ' + $_.UsedGB +' GB'}}
    $FixedVolumeString = $FixedVolumes2.String -join " | "
    Write-output "`n[Detecting $datasource DataSources]"
    Write-output "Fixed Volumes             | $FixedVolumeString"

    <#
    Write-output "Items that can be protected ??"
    get-Volume | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.Driveletter)} | Format-Table

    ## Fixed Volumes + Operationsl + Drive letter

    Write-output "Items not able to be protected ??"
    get-Volume | Where-Object {($_.DriveType -ne "Fixed") -or ($_.OperationalStatus -ne "OK") -or ($_.DriveLetter -eq $null)} | Format-Table

    ## Unprotected Removable, Non Operational, or No Drive Letter
    #>

} ## Get Fixed Volumes

Function Get-USBVolumes {
    # ----- Get Disk Partion Letter for USB / Non USB Bustype ----
        Get-Disk | Select-Object Number | Update-Disk
        $Disk = Get-Disk | Where-Object -FilterScript {$_.Bustype -eq "USB"} | Select-Object Number
    # (Exclude Null Partition Drive Letters)
         try {
            $USBvol = Get-Partition -DiskNumber $Disk.Number | Where-Object {$_.DriveLetter -ne "`0"} | Select-Object @{name="DriveLetter"; expression={$_.DriveLetter+":\"}} | Sort-Object DriveLetter
         }catch{}
    # ----- End Get Disk Partion Letter for USB / Non USB Bustypes ----
        if ($USBvol) {
            $USBvolString = $USBvol.driveletter.replace("\","\\") -join " | " 
            Write-output "USB Volumes               | $($USBvolString.replace("\\","\")) "
        }
    }

Function Get-Filters {
    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { 
        Write-Output "Backup Manager Not Running" 
    }
    else { 
        try { # Command(s) to try
            $ErrorActionPreference = 'Stop'
            & $clienttool -machine-readable control.filter.list | out-file C:\programdata\mxb\filters.csv

            $Filters = import-csv -path C:\programdata\mxb\filters.csv -Header value

            if ($Filters) {
                $FilterString = $Filters.value.replace("\","\\") -join " | "
                Write-output "`n[$Datasource Filters]"
                Write-output "Runtime Exclusion Filters | $($FilterString.replace("\\","\"))"
            }
        }
        catch{ # What to do with terminating errors
        }
    }

} ## Get Filters from ClientTool.exe

Function Get-BackupErrors ($datasource,[int]$Limit=5){
    $retrycounter = 0
    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
            if ($Null -eq ($FunctionalProcess)) {
                Write-Warning "Backup Manager Not Running"
                $global:failed = 1 
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
        
                            $lastsession = ($sessions | Where-object {($_.type -eq "Backup") -and ($_.State -ne "Skipped")})[0] ## Inprogress does clear last error
        
                            & $clienttool -machine-readable control.session.error.list -datasource $DataSource -limit $limit -time $lastsession.start > "C:\ProgramData\MXB\Backup Manager\$DataSource.Errors.tsv"
                            
                            $Script:sessionerrors = Import-Csv -Delimiter "`t" -Path "C:\ProgramData\MXB\Backup Manager\$DataSource.Errors.tsv"
        
                            If ($null -ne $Script:sessionerrors) {
                                $Script:ErrorDateTimeOutTxt = $Script:sessionerrors[-1].datetime
                                $Script:ErrorContentOutTxt = $Script:sessionerrors[-1].content 
                                $Script:ErrorPathOutTxt = $Script:sessionerrors[-1].path
        
                                Write-Warning "[$datasource] Errors Found"
                                $global:failed = 1
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
                    }
                }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 
                if ($DebugDetail) {write-output $retrycounter $BackupStatus}
            }

} ## Get Last Error per Active Data Source from ClientTool.exe

Function Get-StatusReportBase {

    $script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    if (Test-Path $script:StatusReportxml ) {    

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

        if ($Script:OsVersionOutTxt -like "*Server*") { 
            [int]$Script:SuccessHoursInVal = $ServerSuccessHoursInVal
        }

        if ($Script:OsVersionOutTxt -notlike "*Server*") {
            [int]$Script:SuccessHoursInVal = $WrkstnSuccessHoursInVal
        }

        Write-Output "`n[Device]"
        Write-Output "Device Name               | $Script:CoveDeviceNameOutTxt"
        Write-Output "Machine Name              | $Script:MachineNameOutTxt"
        Write-Output "Customer Name             | $Script:CustomerNameOutTxt"
        Write-Output "TimeStamp(UTC)            | $Script:TimeStampUTCOutTxt"
        Write-Output "TimeZone                  | $Script:TimeZoneOutVal"
        Write-Output "OS Version                | $Script:OsVersionOutTxt"
        Write-Output "Profile Name              | $Script:ProfileNameOutTxt"  
        Write-Output "Profile Id                | $Script:ProfileIdOutVal"
        Write-Output "Profile Ver               | $Script:ProfileVersionOutVal"

        }  ## Test if StatusReport.xml file exists
    else{Write-Warning "[$script:StatusReportxml] not found"; $global:failed = 1}
} ## Get Device Data from StatusReport.xml

Function Get-StatusReport {

    $script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    if (Test-Path $script:StatusReportxml ) {    

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
        
        $PluginColorBar                 = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//Plugin$XMLDataSource-ColorBar")."#text"
        $Script:RetentionUnitsOutTxt    = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//RetentionUnits")."#text"

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
                Write-Output "28-Day Status             | $Script:PluginColorBarOutTxt"
                Write-Output "Retention                 | $Script:PluginRetentionOutVal $Script:RetentionUnitsOutTxt "   
                
                Write-Output "Last Success (UTC)        | $Script:LastSuccessTimeUTCOutTxt $Script:LastSuccessStatusOutTxt"
                Write-Output "Last Complete (UTC)       | $Script:LastCompleteTimeUTCOutTxt $Script:LastCompleteStatusOutTxt"      
                Write-Output "Last Session (UTC)        | $Script:LastSessionTimeUTCOutTxt $Script:LastSessionStatusOutTxt"
                Write-Output "Session Duration (HRS)    | $Script:PluginSessionDurationHrsOutVal"
                Write-Output "Success Check (HRS)       | $SuccessHoursInVal"

                if ($LastSuccessTimeUTCOutTxt -gt 0) { 
                    $CurrentTime = Get-Date
                    $TimeSinceLast = New-TimeSpan -Start $LastSuccessTimeUTCOutTxt -End $CurrentTime.ToUniversalTime()
                    [decimal]$HoursSinceLastOutVal = $TimeSinceLast.Totalhours | rnd '' 2
                    if ($HoursSinceLastOutVal -le $SuccessHoursInVal) {
                        Write-Output "Last Success              | $HoursSinceLastOutVal(HRS) Ago"
                    }elseif ($HoursSinceLastOutVal -gt $SuccessHoursInVal) {
                        Write-Warning "Last Success     | $HoursSinceLastOutVal(HRS) Ago"
                        $global:failed = 1
                    }
                }
            }
            Write-Output "Last Error Count          | $Script:LastErrorCountOutVal"

            if ($Script:LastErrorCountOutVal -ge 1) {Get-BackupErrors $DataSource 5}

        }else{Write-Warning "No Prior $Datasource Backup has completed"; $global:failed = 1} ## Prevent null values

        }  ## Test if StatusReport.xml file exists
    else{Write-Warning "[$script:StatusReportxml] not found"; $global:failed = 1 }
} ## Get Data Source Data from StatusReport.xml

Function Get-BackupSelections {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    if ($Null -eq ($FunctionalProcess)) { Write-Warning "Backup Manager Not Running"; $global:failed = 1 }else{ try { $ErrorActionPreference = 'Stop'; $BackupSelections = & $clienttool -machine-readable control.selection.list -delimiter "'" | ConvertFrom-String -Delimiter "'" | Select-Object -skip 1 -property @{N='DataSource'; E={$_.P1}},@{N='Type'; E={$_.P2.replace("Inclusive","Selected (+)").replace("Exclusive","Excluded (-)")}},@{N='Path'; E={$_.P4}} | where-object {$_.DataSource -eq $datasource}}catch{ Write-Warning "ERROR     | $_" }}

    Write-output "`n[$datasource Backup Selections]"

    if ($BackupSelections) {
        $BackupSelections | Select-object Type,Path | Format-Table
    }else { Write-Warning "No data selections exist for $DataSource."; $global:failed = 1}
} ## Get Backup Selections from ClientTool.exe

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
                Write-Warning "ERROR     | $_"
                $global:failed = 1 
            }
        }
    }until (($Backupstatus -notcontains $StatusValue) -or ($retrycounter -ge 5)) 

    $Script:StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 

    $Script:LocalSpeedVaultEnabled                     = ([Xml] (get-content $script:StatusReportxml)).SelectSingleNode("//LocalSpeedVaultEnabled")."#text"
    $Script:BackupServerSynchronizationStatus          = ([Xml] (get-content $script:StatusReportxml)).SelectSingleNode("//BackupServerSynchronizationStatus")."#text"
    $Script:LocalSpeedVaultSynchronizationStatus       = ([Xml] (get-content $script:StatusReportxml)).SelectSingleNode("//LocalSpeedVaultSynchronizationStatus")."#text"
    $Script:SelectedSize                               = ([Xml] (get-content $script:StatusReportxml)).SelectSingleNode("//PluginTotal-LastCompletedSessionSelectedSize")."#text"
    $Script:UsedStorage                                = ([Xml] (get-content $script:StatusReportxml)).SelectSingleNode("//UsedStorage")."#text"
    if (($Script:LocalSpeedVaultEnabled -eq 1) -and ($Script:LSVSettings)) {
        $LSVPath                                = ($Script:LSVSettings | Where-Object { $_ -like "LocalSpeedVaultLocation *"}).replace("LocalSpeedVaultLocation ","")
        $LSVUser                                = ($Script:LSVSettings | Where-Object { $_ -like "LocalSpeedVaultUser *"}).replace("LocalSpeedVaultUser     ","")
    }

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
   
    Write-Output "`n[LocalSpeedVault Status]"
    Write-Output "LSV Enabled               | $Script:LSVEnabledOutTxt"
    Write-Output "LSV Sync Status           | $Script:LSVSyncStatusOutTxt"
    Write-Output "Cloud Sync Status         | $Script:CloudSyncStatusOutTxt"
    Write-Output "Selected Size (GB)        | $Script:TotalSelectedGBOutTxt"        
    Write-Output "Used Storage (GB)         | $Script:TotalUsedGBOutTxt"
    Write-Output "Last LSV Path             | $LSVPath"
    Write-Output "Last LSV User             | $LSVUser"

    # Fail if LSV Enabled = True & ( LSV Sync = Failed or Cloud Sync = Failed )
    if (($LocalSpeedVaultEnabled -eq 1) -and (($BackupServerSynchronizationStatus -eq "Failed") -or ($LocalSpeedVaultSynchronizationStatus -eq "Failed"))) {Write-Warning "LSV Failed"; $global:failed = 1}
            
    elseif (($Script:LocalSpeedVaultEnabled -eq 1) -and ($BackupServerSynchronizationStatus -ne "synchronized")) { 
        if ( ($BackupServerSynchronizationStatus.replace("%","")/1 -lt $SynchThreshold )){Write-Warning "Cloud sync is below $SynchThreshold%"; $global:failed = 1}
    } ## Warn if Sync % is below threshold
        
    elseif (($LocalSpeedVaultEnabled -eq 1) -and ($LocalSpeedVaultSynchronizationStatus -ne "synchronized")) { 
        if ( ($LocalSpeedVaultSynchronizationStatus.replace("%",""))/1 -lt $SynchThreshold ){Write-Warning "LSV sync is below $SynchThreshold%"; $global:failed = 1}
    }  ## Warn if Sync % is below threshold

    # Warn if LSV Enabled = False & LSV Path != ""
    
} ## Get LSV Status

#endregion Functions

Get-BackupState
Get-StatusReportBase
Get-Status $SynchThreshold
Get-Datasources

foreach ($datasource in $datasources) {
    Switch ($script:datasource) { ## "FileSystem","SystemState","VMware","VssHyperV","VssMsSql","Exchange","MySql","NetworkShares","VssSharePoint"
        'FileSystem' {
            Get-FixedVolumes
            Get-USBVolumes
            Get-BackupSelections
            Get-Filters
            }
        'NetworkShares' {
            Get-BackupSelections
            Get-Filters
            }
        'SystemState' {
            Get-BackupSelections
            }     
        'VMware' {
            Get-BackupSelections
            }           
        'MySql' {
            $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Oracle*") -and ($_.name -like "MySQL*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
            Write-output "`n[Detecting $datasource DataSources]"
            if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
            else{Write-Warning "No $Datasource installation detected"; $global:failed = 1}
            Get-BackupSelections
            }
        'Oracle' {
            $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Oracle*") -and ($_.name -like "*Oracle*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
            Write-output "`n[Detecting $datasource DataSources]"
            if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
            else{Write-Warning "No $Datasource installation detected"; $global:failed = 1}
            Get-BackupSelections
            }
        'VssMsSql' {
            $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "Microsoft SQL Server*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
            Write-output "`n[Detecting $datasource DataSources]"
            if ($DatasourcePresent ) {$DatasourcePresent  | Format-Table}
            else{Write-Warning "No $Datasource installation detected"; $global:failed = 1}
            Get-BackupSelections
            }
        'Exchange' {
            $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "*Exchange*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
            Write-output "`n[Detecting $datasource DataSources]"
            if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
            else{Write-Warning "No $Datasource installation detected"; $global:failed = 1}
            Get-BackupSelections
            }
        'VssSharePoint' {
            $DatasourcePresent = Get-WmiObject -Class Win32_Product | where-object {($_.vendor -like "*Microsoft*") -and ($_.name -like "*Sharepoint*")} | Select-Object Name,Vendor,version -Unique | Sort-Object Name
            Write-output "`n[Detecting $datasource DataSources]"
            if ($DatasourcePresent ) {$DatasourcePresent | Format-Table}
            else{Write-Warning "No $Datasource installation detected"; $global:failed = 1}
            Get-BackupSelections
            }
        }
      
        Get-StatusReport
}

#If $global:failed is 1, cause N-sight RMM scriptcheck to fail in dashboard
if ($global:failed -eq 1) {
	Exit 1001
} else {
	Exit 0
}