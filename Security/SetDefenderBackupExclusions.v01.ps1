<# ----- About: ----
    # Set Defender Exclusions for SW Backup
    # Revision v01 - 2021-02-11
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
    # For use with all editions of SolarWinds Backup
    # Tested on Windows 10 & Server 2019
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Sets Windows Defender Exclusions for SolarWinds Backup
    # Includes Backup Manager and Recovery Console applications
    # Reduces disk IO during Backup and Recovery
    #
    # https://docs.microsoft.com/en-us/powershell/module/defender/?view=win10-ps
    # https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-antivirus/configure-process-opened-file-exclusions-microsoft-defender-antivirus
    # https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-antivirus/configure-extension-file-exclusions-microsoft-defender-antivirus
# -----------------------------------------------------------#>  ## Behavior

    $ExcludePath =    @("C:\Program Files\Backup Manager",
                        "C:\Program Files\RecoveryConsole",
                        "C:\ProgramData\Managed Online Backup",
                        "C:\ProgramData\MXB")

    $ExcludeProcess = @("C:\Program Files\Backup Manager\BackupFP.exe",
                        "C:\Program Files\Backup Manager\BackupIP.exe",
                        "C:\Program Files\Backup Manager\ClientTool.exe",
                        "C:\Program Files\Backup Manager\BRMigrationTool.exe",
                        "C:\Program Files\Backup Manager\ProcessController.exe",
                        "C:\Program Files\RecoveryConsole\BackupFP.exe",
                        "C:\Program Files\RecoveryConsole\BackupIP.exe",
                        "C:\Program Files\RecoveryConsole\ClientTool.exe",
                        "C:\Program Files\RecoveryConsole\ProcessController.exe",
                        "C:\Program Files\RecoveryConsole\BRMigrationTool.exe")

    Add-MpPreference -ExclusionPath $ExcludePath -ExclusionProcess $ExcludeProcess 

    $Exclude = get-mppreference | Select-Object Exclusion*

    Write-Output " Defender Process Exclusions`n" $Exclude.ExclusionProcess
    Write-Output "`n Defender Path Exclusions`n" $Exclude.ExclusionPath
    Write-Output "`n Defender Extension Exclusions`n" $Exclude.ExclusionExtension
    Write-Output "`n Defender IP Exclusions`n" $Exclude.ExclusionIpAddress





