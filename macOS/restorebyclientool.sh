# Start a backup using Backup Manager's command line tool
#"/Applications/Backup Manager.app/Contents/MacOS/ClientTool" control.backup.start -datasource FileSystem

# Get Help for restores using Backup Manager's command line tool
#"/Applications/Backup Manager.app/Contents/MacOS/ClientTool" help -c control.restore.start

# list sessions
#"/Applications/Backup Manager.app/Contents/MacOS/ClientTool" control.session.list 


# Trigger a restore using Backup Manager's command line tool
"/Applications/Backup Manager.app/Contents/MacOS/ClientTool" control.restore.start -datasource FileSystem -selection /Users/ -exclude /Users/Shared -existing-files-restore-policy Overwrite -outdated-files-restore-policy CheckContentOfAllFiles 

# Wait for restore to start (not Idle)
counter=0
while true; do
    status=$("/Applications/Backup Manager.app/Contents/MacOS/ClientTool" control.status.get)
    echo "$status"
    if [[ "$status" != *"Idle"* ]]; then
        echo "Restore started - status is no longer idle"
        break
    fi
    counter=$((counter + 10))
    if [[ $counter -ge 60 ]]; then
        echo "Restore job failed to start"
        exit 1
    fi
    sleep 10
done

# Monitor restore status every 10 seconds until idle
while true; do
    status=$("/Applications/Backup Manager.app/Contents/MacOS/ClientTool" control.status.get)
    echo "$status"
    if [[ "$status" == *"Idle"* ]]; then
        echo "Restore completed - status is idle"
        break
    fi
    sleep 10
done

# Command:
#     control.restore.start
#         Start restore. Restores specific or all nodes of specific datasource which were backed up at specific time. Restore destination path
#         can also be specified (but is optional).
# 
#         Example arguments use for nodes selection:
#             -selection /folder1/file1
#             -selection /folder_1/file_2
#             -selection /folder_2
# 
#         You could view existing sessions using `control.session.list' command (which also prints start time for each session).
# 
# Required command arguments:
#    -datasource <NAME>
#         Datasource to start restore for. Possible values are FileSystem.
# 
# Optional command arguments:
#    -add-suffix
#         Add suffix to restored files.
#    -exclude <PATH>
#         Path to node (file, folder, etc.) to exclude from FileSystem restore session. Can be used multiple times to exclude different nodes.
#    -existing-files-restore-policy <POLICY>
#         Existing files restore policy. Possible values are Overwrite or Skip. Default value is Overwrite
#    -outdated-files-restore-policy <POLICY>
#         Outdated files restore policy. Possible values are CheckContentOfAllFiles or CheckContentOfOutdatedFilesOnly. Default value is
#         CheckContentOfAllFiles
#    -restore-to <DIR>
#         Path to start restore of selected sessions to. Default value is empty (in-place restore).
#    -selection <PATH>
#         Path to node (file, folder, etc.) to include in restore session. Can be used multiple times to select different nodes. If not
#         specified, all session nodes are restored.
#    -session-search-policy <POLICY>
#         Backup session search policy. Possible values are ClosestToRequested or OldestIfRequestedNotFound. Default value is ClosestToRequested
#    -time <DATETIME>
#         Start time of backup session. Value must be provided in format "yyyy-mm-dd hh:mm:ss". Default is to restore the most recent session.
# 
# Global arguments:
#    -machine-readable
#         Produce command output (if any) in machine-readable format.
#    -non-interactive
#         Do not ask questions (if any).
#    -version
#         Print program version and exit.
