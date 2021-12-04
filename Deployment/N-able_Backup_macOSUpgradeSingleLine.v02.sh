<<-----About----
    # N-able Backup for macOS - Upgrade
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
    # Downloads and deploys latest Backup Manager over existing installations
    # Run this Script from the TakeControl System Shell, Terminal or Putty
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent-mac.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/full-disk-access.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-mac.htm

-----Behavior----

# Begin Upgrade Script Current macOS versions from ROOT user (Take Control System Shell) or Sudo elevated user (macOS Terminal, SSH or Putty) and prompt for password

    UPGRADE="/Applications/mxb-macosx-x86_64.pkg"; curl -o $UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; sudo installer -pkg $UPGRADE -target /; rm -f $UPGRADE

# Begin Upgrade Script Current macOS versions from Sudo elevated user (macOS Terminal, SSH or Putty) and pipe password to Sudo in script (less secure)

    UPGRADE="/Applications/mxb-macosx-x86_64.pkg"; curl -o $UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; && echo 'PASSWORD' | sudo -S installer -pkg $UPGRADE -target /; rm -f $UPGRADE



# Begin Upgrade Script Some Legacy macOS versions from ROOT user (Take Control System Shell) or Sudo elevated user (macOS Terminal, SSH or Putty) and prompt for password

    UPGRADE="/Applications/mxb-macosx-x86_64.pkg"; curl -o $UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; sudo installer -pkg $UPGRADE -target /Applications; rm -f $UPGRADE

# Begin Upgrade Script Some Legacy macOS versions from Sudo elevated user (macOS Terminal, SSH or Putty) and pipe password to Sudo in script (less secure)

    UPGRADE="/Applications/mxb-macosx-x86_64.pkg"; curl -o $UPGRADE https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg; && echo 'PASSWORD' | sudo -S installer -pkg $UPGRADE -target /Applications; rm -f $UPGRADE