#!/bin/bash

get_orphaned_directories() {
    local orphaned=()
    
    for user_dir in /Users/*; do
        if [[ -d "$user_dir" ]]; then
            local dirname=$(basename "$user_dir")
            
            if [[ "$dirname" != "Shared" && "$dirname" != "Guest" && "$dirname" != ".localized" ]]; then
                if ! dscl . -read /Users/"$dirname" &>/dev/null; then
                    local owner_uid=$(stat -f "%u" "$user_dir" 2>/dev/null || echo "unknown")
                    local owner_name=$(stat -f "%Su" "$user_dir" 2>/dev/null || echo "unknown")
                    local dir_size=$(du -sh "$user_dir" 2>/dev/null | cut -f1 || echo "N/A")
                    local last_modified=$(stat -f "%Sm" -t "%Y-%m-%d_%H:%M" "$user_dir" 2>/dev/null || echo "unknown")
                    
                    orphaned+=("$dirname|$owner_uid|$owner_name|$user_dir|$dir_size|$last_modified|ORPHANED")
                    echo "DEBUG: Found orphaned directory: $dirname"
                fi
            fi
        fi
    done
    
    printf '%s\n' "${orphaned[@]}"
}

echo "Testing orphaned directory detection:"
get_orphaned_directories
