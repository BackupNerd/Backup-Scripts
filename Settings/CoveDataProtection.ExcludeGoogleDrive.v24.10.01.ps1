<# ----- About: ----
    # Cove Data Protection | Set Google Drive Exclusion
    # Revision v24.09.30
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com    
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
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
    # For use with N-able | Cove Data Protection 
    # Sample scripts may contain non-public API calls which are subject to change without notification
    # Some script elements may be developed, tested or documented using AI
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Run script on a recuring scheduled on the local device at startup or via script, RMM or task scheduler
    # Identify if Google Drive is installed
    # Add the Google Drive Letter / Path to the Windows FilesNotToBackup Reg Key
    # Prevent Access Denied to G: and similar errors when Google Drive is installed
    #
    # https://support.google.com/a/answer/7644837
    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/reg-add
    # https://docs.microsoft.com/en-us/windows/win32/backup/registry-keys-for-backup-and-restore#filesnottobackup
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/backup-filters-exclusions.htm
# -----------------------------------------------------------#>  ## Behavior

#region ----- Environment, Variables, Names and Paths ----
Clear-Host
$ConsoleTitle = "Cove Data Protection - Set Google Drive Exclusion"
$host.UI.RawUI.WindowTitle = $ConsoleTitle
Write-Output "$ConsoleTitle`n`n$($myInvocation.MyCommand.Path)"
if ($myInvocation.MyCommand.Path) { Split-path $MyInvocation.MyCommand.Path | Push-Location } ## Set terminal path to match script location (does not support editor 'Run Selection')

$googleDriveKey = "HKCU:\\Software\\Google\\DriveFS"    # Define the registry path for Google Drive
$filesNotToBackupKey = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\BackupRestore\\FilesNotToBackup"  # Define the registry path for FilesNotToBackup

#end region ----- Environment, Variables, Names and Paths ----

$installedPrograms = Get-ItemProperty -Path "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*" | 
    Where-Object { $_.DisplayName -like "*Google Drive*" }

if ($installedPrograms) {
    Write-Host "`nGoogle Drive is installed."
    if (Test-Path $googleDriveKey) {
        try {
            $perAccountPreferences = Get-ItemProperty -Path $googleDriveKey -Name "PerAccountPreferences" -ErrorAction Stop
            if ($perAccountPreferences) {
                $preferences = $perAccountPreferences.PerAccountPreferences | ConvertFrom-Json
                $mountPointPath = $preferences.per_account_preferences.value.mount_point_path
                Write-Output "`nThe current configured mount point or path for Google Drive is: $mountPointPath"
                
                # Add the GoogleDrive path or drive letter to the FilesNotToBackup registry key
                if (Test-Path $filesNotToBackupKey) {
                    if ($mountPointPath.Length -eq 1) {
                        Set-ItemProperty -Path $filesNotToBackupKey -Name "GoogleDrive" -Value "${mountPointPath}:"
                    } elseif ($mountPointPath.Length -gt 1) {
                        Set-ItemProperty -Path $filesNotToBackupKey -Name "GoogleDrive" -Value $mountPointPath
                    }
                    Write-Output "`nThe Google Drive path [ $mountPointPath ] has been added to the Windows FilesNotToBackup registry key."
                } else {
                    Write-Output "`nFilesNotToBackup registry key does not exist."
                }
            } else {
                Write-Warning "`nGoogle Drive is installed, but the PerAccountPreferences property is not found."
            }
        } catch {
            Write-Error "`nAn error occurred while accessing the registry: $_"
        }
    }
} else {
    Write-Host "`nGoogle Drive is not installed."
       # Remove the GoogleDrive entry from the FilesNotToBackup registry key if it exists
    if (Test-Path $filesNotToBackupKey) {
        $property = Get-ItemProperty -Path $filesNotToBackupKey -Name "GoogleDrive" -ErrorAction SilentlyContinue
        if ($property) {
            # Property exists, proceed to remove it
            Remove-ItemProperty -Path $filesNotToBackupKey -Name "GoogleDrive" -ErrorAction SilentlyContinue
            Write-Output "`nGoogleDrive entry has been removed from the FilesNotToBackup registry key."
        }
    } else {
        Write-Output "`nFilesNotToBackup registry key does not exist."
    }
}


