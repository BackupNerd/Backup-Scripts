<<-----About----
    # N-able Backup for macOS - Automatic Deployment
    # Revision v03 - 2021-12-12
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
    # 3 Script examples based on authentication level
    # Downloads and deploys a new Backup Manager as a Passphrase compatible device with an assigned Profile
    # Replace BACKUP_UID and PROFILE_ID variables at the begining of the script
    # Run this Script from the TakeControl System Shell, Terminal, SSH or Putty
    # Remember to enable Full Disk Access for the Backup Manager
    #
    # Name: BACKUP_UID
    # Type: String Variable 
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ Backup.Management | Customers
    # Note: https://documentation.n-able.com/backup/troubleshooting/Content/kb/MSP-How-to-turn-on-Automatic-Deployment-and-generate-a-UID-num
    #
    # Name: PROFILE_ID
    # Type: Integer Variable 
    # Value: ProfileID#
    # Note: Found @ Backup.Management | Profiles
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent-mac.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/full-disk-access.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-mac.htm

-----Behavior----

# Begin Install Script (root authentication, i.e. using N-able TakeControl System Shell, etc.)

    BACKUP_UID="52a304d8-a2eb-489a-9168-8845e0a846dc0"; PROFILE_ID='1285550'; INSTALL="bm#$BACKUP_UID#$PROFILE_ID#.pkg"; cd /tmp; curl -o ./$INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; installer -pkg ./$INSTALL -target /Applications; rm -f ./$INSTALL; cd /

# Begin Install Script (administrative user with sudo prompt for password, i.e. macOS terminal, SSH, Putty, etc.)

	BACKUP_UID="52a304d8-a2eb-489a-9168-8845e0a846dc0"; PROFILE_ID='1285550'; INSTALL="bm#$BACKUP_UID#$PROFILE_ID#.pkg"; cd ~/Downloads; curl -o ./$INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg && sudo -s installer -pkg ./$INSTALL -target /Applications; rm ./$INSTALL; cd /

# Begin Install Script (administrative user w/ piped password to sudo, i.e. macOS terminal, SSH, Putty, etc.) **LESS SECURE - Enter sudo password in SUDOPW variable ** 

    BACKUP_UID="52a304d8-a2eb-489a-9168-8845e0a846dc0"; PROFILE_ID='1285550'; SUDOPW='PASSWORD'; INSTALL="bm#$BACKUP_UID#$PROFILE_ID#.pkg"; cd ~/Downloads; curl -o ./$INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg && echo $SUDOPW | sudo -S installer -pkg ./$INSTALL -target /Applications; rm ./$INSTALL && cd /
