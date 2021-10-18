<<-----About----
    # N-able Backup for Linux - Automatic Deployment
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
-----Compatibility----

<<-----Behavior----
    # Downloads and deploys a new Backup Manager as a Passphrase compatible device with an assigned Profile
    # Replace UID and PROFILEID variables at the begining of the script
    # Run this Script from the TakeControl System Shell, Terminal, SSH or Putty
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
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/run-installer.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/linux-params.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/verify.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-lin.htm


-----Behavior----


# Begin Install Script

UID="01d10b9ee-2a24-4868-9ceb-5e3272124cf0"; PROFILEID='128555'; INSTALL="bm#$UID#$PROFILEID#.run" && curl -o $INSTALL https://cdn.cloudbackup.management/maxdownloads/mxb-linux-x86_64.run && chmod +x $INSTALL && sudo -s ./$INSTALL; rm ./bm#*.run -f

