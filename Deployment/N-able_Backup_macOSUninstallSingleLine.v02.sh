<<-----About----
    # N-able Backup for macOS - Uninstall
    # Revision v02 - 2021-12-03
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
    # Uninstalls N-able Backup on macOS
    # Run this Script from the TakeControl System Shell, Terminal, SSH or Putty
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/uninstall-mac.htm

-----Behavior----

# Begin Uninstall Script from ROOT user (Take Control System Shell) or Sudo elevated user (macOS Terminal, SSH or Putty) and prompt for password

    cd /Applications/Backup\ Manager.app/Contents/Resources/Uninstall.app/Contents/MacOS; sudo bash ./Uninstall.sh; cd /

# Begin Uninstall Script from Sudo elevated user (macOS Terminal, SSH or Putty) and pipe password to Sudo in script (less secure)

    cd /Applications/Backup\ Manager.app/Contents/Resources/Uninstall.app/Contents/MacOS && echo 'PASSWORD' | sudo -S bash ./Uninstall.sh && cd /

