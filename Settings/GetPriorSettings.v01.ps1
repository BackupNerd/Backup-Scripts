<# ----- About: ----
    # N-able Cove Data Protection | Get Prior Backup Manager Settings
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
    # Script does not capture passwords for LocalSpeedVault, Network Shares, etc.
    # Script does not capture Archive schedules
    # Script does not capture Filters
    # Script does not capture Pre/Post Commands
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check for Running Backup Service Controller / Backup Manager (BackupFP.exe)
    # Uses Clienttool.exe command line utiliy to capture and store local device settings, selections & schedules to csv files
    #
    # Use in conjunction with set-prior-settings.v##.ps1 script to apply old settings to a newly deployed device.
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

        if (test-path -Path $SettingsOut){Write-warning "Prior settings found at $SettingsOut"}else{
            $script:Settings = & $clienttool -machine-readable control.setting.list | ConvertFrom-String -Delimiter "`t" -property Name,Value | Select-Object -skip 1 -property Name,Value
            Write-Output "Settings saved to $SettingsOut"
            $script:Settings | Export-Csv -Delimiter "`t" -Path $SettingsOut -NoTypeInformation
        }
        
        if (test-path -Path $SelectionsOut){Write-warning "Prior selections found at $SelectionsOut"}else{
            $script:Selections = & $clienttool -machine-readable control.selection.list | ConvertFrom-String -Delimiter "`t" -property DataSource,Type,Priority,Path | Select-Object -skip 1 -property DataSource,Type,Priority,Path
            Write-Output "Selections saved to $SelectionsOut"
            $script:Selections | Export-Csv -Delimiter "`t" -Path $SelectionsOut -NoTypeInformation
        }

        if (test-path -Path $SchedulesOut){Write-warning "Prior schedules found at $SchedulesOut"}else{
            $script:Schedules = & $clienttool -machine-readable control.schedule.list | ConvertFrom-String -Delimiter "`t" -property Id,Actv,Name,Time,Days,Dsrc,PreSId,PostSId | Select-Object -skip 1 -property Id,Actv,Name,Time,Days,Dsrc,PreSId,PostSId
            Write-Output "Schedules saved to $SchedulesOut"
            $script:Schedules | Export-Csv -Delimiter "`t" -Path $SchedulesOut -NoTypeInformation
        }

    }else{ Write-Warning "Cove Data Protection | Backup Functional Process Not Running"; Break}
    
}else{ Write-Warning "Cove Data Protection | Backup Service Controller Not Running"; Break}

Function Remove-Settings {
    remove-item -path $SettingsOut
    remove-item -path $SelectionsOut
    remove-item -path $SchedulesOut
}  ## Debugging function used to remove stored settings
