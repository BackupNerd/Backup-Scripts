<# ----- About: ----
    # N-able Cove Data Protection | Set Prior Backup Manager Settings
    # Revision v2023.05.11 
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
    # Tested with N-able Cove Data Protection 22.11
    # Compatible with Windows OS devices only
    # Not Compatible with Mac or Linux devices
    # Not Compatible with Documents devices
    # Not Compatible with devices using Backup Profiles
    # Script does not set passwords for LocalSpeedVault, Network Shares, etc.
    # Script does not set Archive schedules
    # Script does not set Filters
    # Script does not set Pre/Post Commands
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check for Running Backup Service Controller / Backup Manager (BackupFP.exe)
    # Uses Clienttool.exe command line utiliy to apply local device settings, selections & schedules from previosuly created csv files
    #
    # Use in conjunction with get-prior-settings.v##.ps1 script to store local settings to a csv file
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/command-line.htm

# -----------------------------------------------------------#>  ## Behavior

#Requires -Version 5.1 -RunAsAdministrator

$Script:BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
if ($BackupService.status -eq "Running") {
    Write-Output "Checking service status"
    start-sleep -Seconds 10
    $Script:FunctionalProcess = get-process "BackupFP" -ea SilentlyContinue
    if ($FunctionalProcess) {
        Write-Output "Checking process status"
        start-sleep -Seconds 10
        $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
        $SettingsOut = "c:\programdata\mxb\Backup Manager\Settings.csv"
        $SelectionsOut = "c:\programdata\mxb\Backup Manager\Selections.csv"
        $SchedulesOut= "c:\programdata\mxb\Backup Manager\Schedules.csv"

        if (test-path -Path $SettingsOut){
            Write-Output "Prior settings found at $SettingsOut"
            $Script:Settings = Import-Csv -Path $SettingsOut -Delimiter "`t"
            foreach ($Entry in $Script:Settings) {
                if ($entry.Name -like "*Password") {Continue}  ## Skip entry
                if ($entry.NAme -eq "Device") {Continue} ## Skip entry
                & $clienttool control.setting.modify -name $Entry.Name -value $entry.value
            }
        }else{Write-warning "Prior settings not found at $SettingsOut"}

        if (test-path -Path $SelectionsOut){
            Write-Output "Prior selections found at $SelectionsOut"
            $Script:Selections = Import-Csv -Path $SelectionsOut -Delimiter "`t"
            foreach ($Entry in $Selections) {
                if ($entry.Type -eq "Inclusive") {$entry.Type = "-include"}elseif ($entry.Type -eq "Exclusive") {$entry.Type = "-exclude"}
                if ($entry.Path -eq "") {$entry.Path = "/" }
                & $clienttool control.selection.modify -datasource $Entry.DataSource $entry.Type $entry.Path
            }
        }else{Write-warning "Prior selections not found at $SelectionsOut"}

        if (test-path -Path $SchedulesOut){
            Write-Output "Prior selections found at $SchedulesOut"
            $Script:Schedules = Import-Csv -Path $SchedulesOut -Delimiter "`t"
            foreach ($Entry in $Schedules) {
                if ($entry.Actv -eq "yes") {$entry.Actv = 1}elseif ($entry.Actv -eq "no") {$entry.Actv = 0}

                & $clienttool control.schedule.add -name $Entry.Name -active $Entry.Actv -time $entry.Time -days $Entry.Days -datasources $entry.Dsrc
            }
        }else{Write-warning "Prior schedules not found at $SchedulesOut"}

    }
}

Function Remove-ScheduleRange {
    foreach ($number in (1..500)) {
    & $clienttool control.schedule.remove -id $number
    }
} ## Debugging function used to remove a range of schedules

Function Remove-Schedules {
    $script:Schedules = & $clienttool -machine-readable control.schedule.list | ConvertFrom-String -Delimiter "`t" -property Id,Actv,Name,Time,Days,Dsrc,PreSId,PostSId | Select-Object -skip 1 -property Id
    foreach ($schedule in $schedules) {
        & $clienttool control.schedule.remove -id $schedule.id
    }
} ## Debugging function used to remove configured schedules

Function Remove-Selections {
    & $clienttool control.selection.clear
} ## Debugging function used to remove selections

