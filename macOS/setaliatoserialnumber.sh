#!/bin/bash

newest_partner_uid=""
newest_file=""
newest_time=0
PUID=""

# Check if PUID is provided as command line argument
if [[ $# -eq 1 ]]; then
    PUID="$1"
    echo "Using provided Partner UID: $PUID"
else
    # Find all ClientTool_*.log files and search for PartnerUid values
    while IFS= read -r logfile; do
        if grep -q "PartnerUid=" "$logfile" 2>/dev/null; then
            file_time=$(stat -f %m "$logfile" 2>/dev/null)
            if [[ $file_time -gt $newest_time ]]; then
                newest_time=$file_time
                newest_file="$logfile"
                # Get the last (newest) PartnerUid entry from this file
                newest_partner_uid=$(grep "PartnerUid=" "$logfile" | sed 's/.*PartnerUid=\([^[:space:]]*\).*/\1/' | tail -1)
                PUID="$newest_partner_uid"
            fi
            
            echo "File: $logfile"
            current_puid=$(grep "PartnerUid=" "$logfile" | sed 's/.*PartnerUid=\([^[:space:]]*\).*/\1/' | sort -u)  
            echo "---"
            echo $current_puid
        fi
    done < <(find /Library/Logs/MXB/Backup\ Manager/ClientTool -name "ClientTool_*.log" -type f 2>/dev/null)
fi

# Get the local Mac serial number
SERIAL_NUMBER=$(system_profiler SPHardwareDataType | grep "Serial Number" | awk '{print $4}')

# Convert serial number to base64
ALIAS_BASE64=$(echo -n "$SERIAL_NUMBER" | base64)
echo "Local Serial Number (RawTxt): $SERIAL_NUMBER"
echo "Local Serial Number (Base64): $ALIAS_BASE64"
echo "Partner UID: $PUID"

# Execute the command with proper escaping
"/Applications/Backup Manager.app/Contents/MacOS/ClientTool" takeover -config-path "/Library/Application Support/MXB/Backup Manager/config.ini" -partner-uid "$PUID" -device-alias "$ALIAS_BASE64"
