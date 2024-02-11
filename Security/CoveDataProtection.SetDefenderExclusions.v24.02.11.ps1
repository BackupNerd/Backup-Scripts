<# ----- About: ----
    # Cove Data Protection | Set Exclusions for Windows Defender
    # Revision v24.02.11 - 2024-02-11
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com
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
    # For use with the standalone edition of N-able | Cove Data Protection
    # Tested on Windows 10 & Server 2019
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Sets Cove Data Protection exclusions in Windows Defender
    # Includes Backup Manager and Recovery Console applications
    # Reduces disk IO, RAM and CPU utilization during Backup and Recovery
    #
    # https://docs.microsoft.com/en-us/powershell/module/defender/?view=win10-ps
    # https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-antivirus/configure-process-opened-file-exclusions-microsoft-defender-antivirus
    # https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-antivirus/configure-extension-file-exclusions-microsoft-defender-antivirus
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/reqs.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/advanced-recovery/recovery-console/installation.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/recovery-locations/azure-recovery-location.htm

# -----------------------------------------------------------#>  ## Behavior

    $ExcludePath =    @("C:\Program Files\Backup Manager",
                                 
                        "C:\Program Files\RecoveryConsole",
                        "C:\Program Files\RecoveryConsole\vddk",
                        
                        "C:\ProgramData\Managed Online Backup",
                        "C:\ProgramData\MXB",

                        "*\StandbyImage",
                        "*\OnDemandRestore"
                        )

    $ExcludeProcess = @("C:\Program Files\Backup Manager\BackupFP.exe",
                        "C:\Program Files\Backup Manager\BackupIP.exe",
                        "C:\Program Files\Backup Manager\BackupUP.exe",
                        "C:\Program Files\Backup Manager\ClientTool.exe",
                        "C:\Program Files\Backup Manager\BRMigrationTool.exe",
                        "C:\Program Files\Backup Manager\ProcessController.exe",

                        "C:\Program Files\RecoveryConsole\BackupFP.exe",
                        "C:\Program Files\RecoveryConsole\BackupIP.exe",
                        "C:\Program Files\RecoveryConsole\BackupUP.exe",
                        "C:\Program Files\RecoveryConsole\ClientTool.exe",
                        "C:\Program Files\RecoveryConsole\RecoveryConsole.exe",
                        "C:\Program Files\RecoveryConsole\ProcessController.exe",
                        "C:\Program Files\RecoveryConsole\BRMigrationTool.exe",

                        "C:\Program Files\Recovery Service\*\AuthTool.exe",
                        "C:\Program Files\Recovery Service\*\unified_entry.exe",
                        "C:\Program Files\Recovery Service\*\BM\RecoveryFP.exe",
                        "C:\Program Files\Recovery Service\*\BM\VdrAgent.exe",
                        "C:\Program Files\Recovery Service\*\BM\ProcessController.exe",
                        "C:\Program Files\Recovery Service\*\BM\RecoveryProcessController.exe",
                        "C:\Program Files\Recovery Service\*\BM\ClientTool.exe",
                        "C:\Program Files\Recovery Service\*\VdrTool.exe"
                        )

    Add-MpPreference -ExclusionPath $ExcludePath -ExclusionProcess $ExcludeProcess 

    $Exclude = get-mppreference | Select-Object Exclusion*

    Write-Output " Defender Process Exclusions`n" $Exclude.ExclusionProcess
    Write-Output "`n Defender Path Exclusions`n" $Exclude.ExclusionPath
    Write-Output "`n Defender Extension Exclusions`n" $Exclude.ExclusionExtension
    Write-Output "`n Defender IP Exclusions`n" $Exclude.ExclusionIpAddress





