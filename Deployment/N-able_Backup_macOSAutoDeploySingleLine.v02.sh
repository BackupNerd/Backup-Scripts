<<-----About----
    # N-able Backup for macOS - Automatic Deployment
    # Revision v02 - 2021-10-17
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
-----About----

<<-----Legal----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
-----Legal----

<<-----Compatibility----
    # For use with the Standalone edition of N-able Backup
    # Tested with N-able Backup 21.10
    # Tested with N-able TakeControl System Shell and Terminal on macOS 11.6 & 10.15
-----Compatibility----

<<-----Behavior----
    # Downloads and deploys a new Backup Manager as a Passphrase compatible device with an assigned Profile
    # Replace UID and PROFILEID variables at the begining of the script
    # Run this Script from the TakeControl System Shell, Terminal or Putty
    # Remember to enable Full Disk Access for the Backup Manager
    #
    # Name: BACKUP_UID
    # Type: String Variable 
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ Backup.Management | Customers
    #
    # Name: PROFILEID
    # Type: Integer Variable 
    # Value: ProfileID #
    # Note: Found @ Backup.Management | Profiles
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent-mac.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/full-disk-access.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-mac.htm

-----Behavior----

# Begin Install Script

BACKUP_UID="1d10b9ee-2a24-4868-9ceb-5e3272124cf0"; PROFILEID='128555'; INSTALL="/Applications/bm#$BACKUP_UID#$PROFILEID#.pkg"; curl -o $INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; sudo installer -pkg $INSTALL -target /Applications; rm -f $INSTALL

