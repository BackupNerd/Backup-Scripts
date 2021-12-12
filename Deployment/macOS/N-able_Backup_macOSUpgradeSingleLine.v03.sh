<<-----About----
    # N-able Backup for macOS - Upgrade
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
    # Downloads and deploys latest Backup Manager over existing installations
    # Run this Script from the TakeControl System Shell, Terminal or Putty
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent-mac.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/full-disk-access.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-mac.htm

-----Behavior----

# Begin Upgrade Script (root authentication, i.e. using N-able TakeControl System Shell, etc.)

    UPGRADE="mxb-macosx-x86_64.pkg"; cd /tmp; curl -o ./$UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; installer -pkg ./$UPGRADE -target /Applications; rm -f ./$UPGRADE; cd / 

# Begin Upgrade Script (administrative user with sudo prompt for password, i.e. macOS terminal, SSH, Putty, etc.)

    UPGRADE="mxb-macosx-x86_64.pkg"; cd ~/Downloads; curl -o ./$UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg && sudo -s installer -pkg ./$UPGRADE -target /Applications; rm -f ./$UPGRADE && cd / 

# Begin Upgrade Script (administrative user w/ piped password to sudo, i.e. macOS terminal, SSH, Putty, etc.) **LESS SECURE - Enter sudo password in SUDOPW variable **

	SUDOPW='PASSWORD'; UPGRADE="mxb-macosx-x86_64.pkg"; cd ~/Downloads; curl -o ./$UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg && echo $SUDOPW | sudo -S installer -pkg ./$UPGRADE -target /Applications; rm -f ./$UPGRADE && cd / 

