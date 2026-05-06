#!/bin/bash

# ----- About: ----------------------------------------------------------------
#   Restore by ClientTool | Restore Most Recent Session to Mounted Volume
#   Revision v02.5 - 2026-05-06
#   Author: Eric Harless, Head Backup Nerd - N-able
#   Twitter @Backup_Nerd  Email:eric.harless@n-able.com
# -----------------------------------------------------------------------------

# ----- Legal: ----------------------------------------------------------------
#   Sample scripts are not supported under any N-able support program or service.
#   The sample scripts are provided AS IS without warranty of any kind.
#   N-able expressly disclaims all implied warranties including, warranties
#   of merchantability or of fitness for a particular purpose.
#   In no event shall N-able or any other party be liable for damages arising
#   out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------------------------

# ----- Compatibility: --------------------------------------------------------
#   For use with the Standalone edition of N-able Cove Data Protection
#   Platform: GNU/Linux
#   Requires Backup Manager installed at: /opt/MXB/bin/
#   ClientTool path: /opt/MXB/bin/ClientTool  (must be run as root or with sudo)
#   Sample scripts may contain non-public API calls which are subject to change without notification
#   Ref: https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/command-line.htm
# -----------------------------------------------------------------------------

# ----- Use Case: -------------------------------------------------------------
#   Restore the most recent backup session to a pre-mounted destination volume
#   for recovery, validation, or data extraction without touching the live
#   filesystem.  Volume creation and mount must be performed by the Linux
#   administrator before running this script.
#
#   Exclude the restore destination from backup selection so that restored data
#   is not re-protected.  When using a shared network path across multiple
#   systems, every system that can reach that path must exclude it from its own
#   backup scope.
#
#   Note: NTFS-based network shares do not preserve Linux file ownership and
#   permissions; plan accordingly for permission-sensitive workloads.
# -----------------------------------------------------------------------------

CLIENTTOOL="/opt/MXB/bin/ClientTool"
SELECTION="/"
EXCLUDE="/mnt/restore"                                        # Optional: set to path to exclude from restore, e.g. "/mnt/restore". Leave empty to skip.
RESTORE_TO="/mnt/restore/$(hostname)/$(date +%Y-%m-%d)"       # Set to destination path for restore location, e.g. "/mnt/restore/$(hostname)/$(date +%Y-%m-%d)". Leave empty for in-place restore.
TIMEOUT=60

# Confirm running as root
if [[ $EUID -ne 0 ]]; then
    echo "This script must be run as root (e.g. sudo $0)"
    exit 1
fi

# Start a backup using Backup Manager's command line tool
#"$CLIENTTOOL" control.backup.start -datasource FileSystem

# Get help for restores using Backup Manager's command line tool
#"$CLIENTTOOL" help -c control.restore.start

# List sessions
#"$CLIENTTOOL" control.session.list

# Create restore destination path if set and not already existing
if [[ -n "$RESTORE_TO" ]]; then
    # RESTORE_TO may include subdirs that don't exist yet (e.g. /mnt/restore/hostname/date).
    # Walk up the path to find the first existing directory that is an actual mount point.
    # This prevents mkdir -p from silently writing to the root volume if the recovery
    # volume is not mounted.
    MOUNT_BASE="$RESTORE_TO"
    while [[ "$MOUNT_BASE" != "/" ]]; do
        if [[ -d "$MOUNT_BASE" ]] && mountpoint -q "$MOUNT_BASE"; then
            break
        fi
        MOUNT_BASE="$(dirname "$MOUNT_BASE")"
    done
    if [[ "$MOUNT_BASE" == "/" ]]; then
        echo ""
        echo "  *** ERROR: RESTORE DESTINATION IS NOT MOUNTED ***"
        echo "  No mount point found in path: $RESTORE_TO"
        echo "  Mount the recovery volume before running this script."
        echo ""
        exit 1
    fi
    mkdir -p "$RESTORE_TO" || { echo "Failed to create restore path: $RESTORE_TO"; exit 1; }
    echo "Restore destination created/verified: $RESTORE_TO"
fi

# Require RESTORE_TO - in-place restore is not allowed
if [[ -z "$RESTORE_TO" ]]; then
    echo ""
    echo "  *** ERROR: RESTORE_TO IS REQUIRED ***"
    echo "  In-place restore is not allowed. RESTORE_TO must be set to a"
    echo "  valid restore destination path."
    echo ""
    exit 1
fi

# Trigger a restore using Backup Manager's command line tool
RESTORE_CMD=("$CLIENTTOOL" control.restore.start \
    -datasource FileSystem \
    -selection "$SELECTION" \
    -existing-files-restore-policy Overwrite \
    -outdated-files-restore-policy CheckContentOfAllFiles \
    ${EXCLUDE:+-exclude "$EXCLUDE"} \
    ${RESTORE_TO:+-restore-to "$RESTORE_TO"})

"${RESTORE_CMD[@]}"

# Wait for restore to start (status must leave Idle)
counter=0
while true; do
    status=$("$CLIENTTOOL" control.status.get)
    echo "$status"
    if [[ "$status" != *"Idle"* ]]; then
        echo "Restore started - status is no longer idle"
        break
    fi
    counter=$((counter + 10))
    if [[ $counter -ge $TIMEOUT ]]; then
        echo "Restore job failed to start within ${TIMEOUT}s"
        exit 1
    fi
    sleep 10
done

# Monitor restore status every 10 seconds until idle
while true; do
    status=$("$CLIENTTOOL" control.status.get)
    echo "$status"
    if [[ "$status" == *"Idle"* ]]; then
        echo "Restore completed - status is idle"
        break
    fi
    sleep 10
done

# Command reference:
#     control.restore.start
#         Start restore. Restores specific or all nodes of specific datasource which were backed up at specific time.
#
#         Example arguments for node selection:
#             -selection /home/username
#             -selection /home/username/Documents
#             -selection /srv/data
#
#         You can view existing sessions using 'control.session.list' command (which also prints start time for each session).
#
# Required command arguments:
#    -datasource <NAME>
#         Datasource to start restore for. Possible values are FileSystem.
#
# Optional command arguments:
#    -add-suffix
#         Add suffix to restored files.
#    -exclude <PATH>
#         Path to node (file, folder, etc.) to exclude from FileSystem restore session. Can be used multiple times.
#    -existing-files-restore-policy <POLICY>
#         Existing files restore policy. Possible values are Overwrite or Skip. Default value is Overwrite.
#    -outdated-files-restore-policy <POLICY>
#         Outdated files restore policy. Possible values are CheckContentOfAllFiles or CheckContentOfOutdatedFilesOnly.
#    -restore-to <DIR>
#         Path to start restore of selected sessions to. Default value is empty (in-place restore).
#    -selection <PATH>
#         Path to node (file, folder, etc.) to include in restore session. Can be used multiple times.
#    -session-search-policy <POLICY>
#         Backup session search policy. Possible values are ClosestToRequested or OldestIfRequestedNotFound.
#    -time <DATETIME>
#         Start time of backup session. Format: "yyyy-mm-dd hh:mm:ss". Default is to restore the most recent session.
#
# Global arguments:
#    -machine-readable     Produce command output in machine-readable format.
#    -non-interactive      Do not ask questions.
#    -version              Print program version and exit.
