<# ----- About: ----
    # Cove Data Protection - Windows Backup Manager - Automatic Deployment (NinjaOne Sample)
    # Revision v24.10 - 2024-10-15
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
    # For use with the Standalone edition of N-able | Cove Data Protection
    # Tested with release 23.12
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # NOTE: Do not use the script as is.  Modify the script to meet your specific requirements.
    #
    # Downloads and deploys a new Backup Manager as a Passphrase compatible device with an assigned Profile
    # Pulls UID, PROFILEID and PRODUCT from Ninja variables
    # Run portions of this script from within NinjaOne RMM
    #
    # Name: UID
    # Type: String Variable 
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ Backup.Management | Customers
    #
    # Name: PROFILEID
    # Type: Integer Variable 
    # Value: ProfileID #
    # Note: Found @ Backup.Management | Profiles (use 0 for No Profile)
    #
    # Name: PRODUCT
    # Type: String Variable
    # Value: Retention Policies or Classic Product Name
    # Note: Found @ Backup.Management | Retention Policies ('All-In' is a case-sensitive default product with 28 days retention)
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-win-silent.htm
    # Ninja Scripting Video: https://youtu.be/uDvodMDsLzQ?si=lqkEekJ2MHq4YW7Q
    

# -----------------------------------------------------------#>

# Begin Install Script Example 1

$UID = Ninja-Property-Get coveCustomerUID                       ## Tenant specific 36 character UID found @ Backup.Management | Customers
$PROFILEID = Ninja-Property-Get coveBackupDefaultProfileID      ## Use '0' for No Profile, Use '5038' for Documents, or a CustomProfileID# found @ Backup.Management | Profiles 
$PRODUCT = Ninja-Property-Get coveBackupDefaultProduct          ## Use "All-In" for 28-Day Retention, Use "Document", or use a Custom Product Name or Retention Policy Name found @ Backup.Management | Retention Policies

$INSTALL="c:\windows\temp\bm#$UID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL -product-name `"$PRODUCT`"

# End Install Script  Example 1




# Begin Install Script Example 2

$UID = Ninja-Property-Get coveCustomerUID
$PROFILEID = Ninja-Property-Get coveBackupDefaultProfileID
$PRODUCT = "Default365"

$INSTALL="c:\windows\temp\bm#$UID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL -product-name `"$PRODUCT`"

# End Install Script Example 2




# Begin Install Script Example 3

$UID="9696c2af4-678a-4727-9b6b-example"; $PROFILEID='128555'; $PRODUCT='All-In'; $INSTALL="c:\windows\temp\bm#$UID#$PROFILEID#.exe"; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)"); & $INSTALL -product-name `"$PRODUCT`"

# End Install Script Example 3