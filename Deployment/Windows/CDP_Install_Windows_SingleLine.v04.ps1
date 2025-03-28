<# ----- About: ----
    # CDP Install Windows - Automatic Deployment, Upgrade, Uninstall - Single Line
    # Revision v04 - 2025-03-27
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
# -----------------------------------------------------------#>
<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>
<# ----- Compatibility: ----
    # For use with N-able | Cove Data Protection 
    # The script requires PowerShell 5.1 or higher and should be run with administrative privileges.
# -----------------------------------------------------------#>
<# ----- Behavior: ----
    # Downloads and uses automatic deployment to add a new (Passphrase compatible) Backup Manager to the dashboard w/ an assigned Profile / Retention Policy
    # The script will download the latest version of the Backup Manager from the CDN and install it silently.
    # The executable will also check if the Backup Manager is already installed and will skip the installation if it is.
    # Replace UID and PROFILEID variables at the begining of the script
    # Optionally supply a RETENTIONPOLICY variable to set the retention policy for the device
    # Run this Script from the N-able TakeControl Shell, PowerShell, or via automation using your N-able or third party RMM or MDM
    #
    # Name: CUID
    # Type: String Variable 
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ https://backup.management | Customers
    #
    # Name: PROFILEID
    # Type: Integer Variable 
    # Value: ProfileID #
    # Note: Found @ https://backup.management  | Profiles (use 0 for No Profile or 5038 for Documents Profiles)
    #
    # Name: RETENTIONPOLICY
    # Type: String Variable
    # Value: Retention Policy / Classic Product Name
    # Note: Found @ Backup.Management | Retention Policies  (i.e. 'All-In' is a case-sensitive classic retention policy with 28 days retention)
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/enable-auto-dep.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/customer-uid.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/profiles.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/retention-policies/retention-policies-overview.htm
    # 
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/update-backup-manager.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/uninstall-win-silent.htm
# -----------------------------------------------------------#>

## Warning: Do not execute this script as is, it contains 3 seperate single line deployment script examples which require modification and seperate upgrade and uninstall examples.   

# Example 1 - Download and install a full edition of Cove - Custom Profile - Uses the Default Retention Policy
$CUID="01d10b9ee-change-me-4868-9ceb-5e3272124cf0"; $PROFILEID='128555'; $INSTALL="c:\windows\temp\bm#$CUID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL

# Example 2 - Download and install a full edition of Cove - Custom Profile - Uses the classic 'All-In' Retention Policy
$CUID="254ebc40-f638-4f97-b1a2-fce5515a405e"; $PROFILEID='36628'; $RETENTIONPOLICY='All-In'; $INSTALL="c:\windows\temp\bm#$CUID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL -product-name "$RETENTIONPOLICY"

# Example 3 - Download and install a Documents edition of Cove
$CUID="01d10b9ee-change-me-4868-9ceb-5e3272124cf0"; $PROFILEID='5038'; $INSTALL="c:\windows\temp\bm#$CUID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL

# Example 4 - If Cove is currently installed, Download the latest installer and upgrade the Backup Manager
if (Test-Path "C:\Program Files\Backup Manager\config.ini") { $INSTALL="c:\windows\temp\cove#update#binariesonly.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL } else { Write-Host "Backup Manager is not installed. Skipping update." }

# Example 5 - Uninstall the Backup Manager
start-process -FilePath "C:\Program Files\Backup Manager\BackupIP.exe" -ArgumentList "uninstall -interactive -path `"C:\Program Files\Backup Manager`" -sure" -PassThru

##### Testing
# Sample script - Copy out the config.ini file to save the settings before running the uninstall script
$source = "C:\Program Files\Backup Manager\config.ini"
$destination = "C:\Data\config.ini"
New-Item -ItemType Directory -Path (Split-Path $destination) -Force | Out-Null
Copy-Item -Path $source -Destination $destination -Force

# Sample script - Copy back the config.ini file to return the original settings before running the upgrade script
$source = "C:\Data\config.ini"
$destination = "C:\Program Files\Backup Manager\config.ini"
New-Item -ItemType Directory -Path (Split-Path $destination) -Force | Out-Null
Copy-Item -Path $source -Destination $destination -Force