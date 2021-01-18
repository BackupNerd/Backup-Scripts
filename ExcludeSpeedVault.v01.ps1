<# ----- About: ----
    # Exclude SpeedVault and Seed path from backup
    # Revision v01 - 2021-01-18
    # Author: Eric Harless, Head Backup Nerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with the Standalone edition of SolarWinds Backup
    # For use with the RMM integrated edition of SolarWinds Backup
    # For use with the N-central edition of SolarWinds Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Set local exclusion filter for LSV and seed drives. "None added" is a vaild response if the fiter has already been applied
    # Note, this method is not supported with profile based filters
    # Set Registry exclusion filter for LSV and seed drives. "HKLM\SYSTEM\ControlSet001\Control\BackupRestore\FilesNotToBackup"
    # Note, this method is supported with profile based filters
    #
    # Excludes local and network attached path containung LSV and seed data
    #
    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/reg-add
    # https://docs.microsoft.com/en-us/windows/win32/backup/registry-keys-for-backup-and-restore#filesnottobackup 
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/command-line.htm?
    #
# -----------------------------------------------------------#>  ## Behavior

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Exclude SpeedVault and Seed path from backup`n"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $urljson = "https://api.backup.management/jsonapi"

#endregion ----- Environment, Variables, Names and Paths ----

# ----- Set Backup Filter/ Reg to Exclude LSV/ Seed with ClientTool ---- 

    Write-Host "  Setting backup filters and the 'FilesNotToBackup' registry key to skip LSV/ seed drive paths`n"

        & 'C:\Program Files\Backup Manager\clienttool.exe' control.filter.modify -add "*\storage\cabs\gen*\*"
    
        & REG ADD "HKLM\SYSTEM\ControlSet001\Control\BackupRestore\FilesNotToBackup" /v "ExcludeLSV" /t REG_MULTI_SZ /d "*\storage\cabs\gen*\*" /f
