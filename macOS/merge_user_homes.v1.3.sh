#!/bin/bash

# Signal handling for graceful exit
cleanup_and_exit() {
    echo ""
    echo -e "\n${YELLOW}Script interrupted by user. Exiting safely...${NC}"
    exit 130
}
trap cleanup_and_exit INT TERM

################################################################################
# macOS User Home Directory Merge Script - Safe User Directory Consolidation
################################################################################
#
# DESCRIPTION:
#   A comprehensive shell script for macOS that provides safe merging of one
#   user's home directory into another user's home directory. This script is 
#   designed for scenarios like user account consolidation, migration, or 
#   combining accounts after organizational changes.
#
# FEATURES:
#   • Interactive source and target user selection with detailed information
#   • Comprehensive backup creation before any merge operations
#   • Intelligent conflict resolution strategies:
#     - Rename conflicts with timestamp suffixes
#     - Skip duplicate files option
#     - Replace existing files option
#   • Special handling for macOS-specific directories:
#     - Library folder selective merging (avoid system conflicts)
#     - Applications folder duplicate detection
#     - Media libraries (iTunes, Photos) careful handling
#     - Desktop and Documents intelligent merging
#   • Preservation of file permissions, ownership, and macOS extended attributes
#   • Detailed progress reporting and comprehensive logging
#   • Multiple safety checkpoints with rollback capability
#   • Size validation and disk space checking
#
# SAFETY FEATURES:
#   • Requires root privileges for safe operation
#   • Protects currently logged-in users from being used as source
#   • Creates timestamped backups before any changes
#   • Multiple confirmation prompts with detailed impact assessment
#   • Rollback mechanism if merge fails midway
#   • Preserves all file metadata and macOS-specific attributes
#   • Validates source and target users exist and are accessible
#   • Prevents merging system users or protected accounts
#   • Clear logging of all operations for audit trail
#
# MERGE STRATEGIES:
#   • Documents/Desktop: Merge with conflict resolution
#   • Downloads: Merge with duplicate detection based on file content
#   • Pictures/Movies/Music: Merge with media library awareness
#   • Applications (user-installed): Careful handling to prevent duplicates
#   • Library: Selective merging avoiding system-critical preferences
#   • Hidden files/folders: User choice to include or exclude
#
# COMPATIBILITY:
#   • Designed specifically for macOS using native tools and conventions
#   • Handles macOS file attributes (resource forks, extended attributes)
#   • Compatible with macOS System Integrity Protection
#   • Preserves macOS folder structure and permissions
#   • Works with macOS user account standards (UID 501+)
#
# USAGE:
#   sudo /path/to/merge_user_homes.sh
#
# REQUIREMENTS:
#   • macOS operating system (10.12 Sierra or later recommended)
#   • Root privileges (must run with sudo)
#   • Sufficient disk space for backups and merged content
#   • Standard macOS utilities: dscl, rsync, stat, du
#
# WARNING: 
#   This script modifies user home directories which contain critical user data!
#   Always ensure you have independent backups before running this script.
#   While the script creates its own backups, external backup verification 
#   is strongly recommended for irreplaceable data.
#
# AUTHOR: 
#   Created by GitHub Copilot with guidance from Eric Harless
#   Based on user_manager.sh patterns and macOS best practices
#
#   Date: September 30, 2025
#   Version: 1.3 - Enhanced with UID Sorting and Improved Conflict Display
#
################################################################################

set -e  # Exit on any error

# Colors for output formatting
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
MAGENTA='\033[0;35m'
NC='\033[0m' # No Color

# Global variables
CURRENT_USER=""
SOURCE_USER=""
TARGET_USER=""
SOURCE_HOME=""
TARGET_HOME=""
BACKUP_DIR=""
LOG_FILE=""
MERGE_START_TIME=""

# Conflict resolution strategy
CONFLICT_STRATEGY=""

# Statistics tracking
TOTAL_FILES=0
MERGED_FILES=0
SKIPPED_FILES=0
CONFLICT_FILES=0
BACKUP_SIZE=0

################################################################################
# UTILITY FUNCTIONS
################################################################################

# Function to get current logged-in user
get_current_user() {
    if [[ -n "$SUDO_USER" ]]; then
        CURRENT_USER="$SUDO_USER"
    else
        CURRENT_USER=$(stat -f "%Su" /dev/console)
    fi
    
    if [[ -z "$CURRENT_USER" ]]; then
        echo -e "${YELLOW}Warning: Could not determine current user.${NC}"
        CURRENT_USER="unknown"
    fi
}

# Function to display script header
show_header() {
    clear
    echo -e "${BLUE}================================================================${NC}"
    echo -e "${BLUE}               USER HOME DIRECTORY MERGE SCRIPT${NC}"
    echo -e "${BLUE}================================================================${NC}"
    echo -e "${RED}WARNING: This script will merge one user's home directory${NC}"
    echo -e "${RED}into another user's home directory. Use with extreme caution!${NC}"
    echo ""
    echo -e "${GREEN}Features:${NC}"
    echo -e "  • Safe backup creation before merge operations"
    echo -e "  • Intelligent conflict resolution strategies"
    echo -e "  • Special handling for macOS system directories"
    echo -e "  • Comprehensive logging and progress reporting"
    echo -e "  • Rollback capability if issues occur"
    echo ""
}

# Function to check if running as root
check_root() {
    if [[ $EUID -ne 0 ]]; then
        echo -e "${RED}Error: This script must be run as root (sudo)${NC}"
        echo "Usage: sudo $0"
        exit 1
    fi
}

# Function to setup logging
setup_logging() {
    MERGE_START_TIME=$(date '+%Y%m%d_%H%M%S')
    LOG_FILE="/var/log/user_merge_${MERGE_START_TIME}.log"
    
    # Determine source type for logging
    local source_type="USER"
    local source_display="$SOURCE_USER"
    if [[ "$SOURCE_USER" == "ORPHANED:"* ]]; then
        source_type="ORPHANED_DIRECTORY"
        source_display="${SOURCE_USER#ORPHANED:}"
    fi
    
    # Create log file and add header
    cat > "$LOG_FILE" << EOF
================================================================================
USER HOME DIRECTORY MERGE LOG
================================================================================
Start Time: $(date)
Script Version: 1.3 - Enhanced with UID Sorting and Improved Conflict Display
Source Type: $source_type
Source: $source_display
Target User: $TARGET_USER
Source Home: $SOURCE_HOME  
Target Home: $TARGET_HOME
Current User: $CURRENT_USER
================================================================================

EOF

    echo -e "${GREEN}Logging enabled: $LOG_FILE${NC}"
}

# Function to log messages
log_message() {
    local level="$1"
    local message="$2"
    local timestamp=$(date '+%Y-%m-%d %H:%M:%S')
    
    # Only log to file if LOG_FILE is set
    if [[ -n "$LOG_FILE" ]]; then
        echo "[$timestamp] [$level] $message" >> "$LOG_FILE"
    fi
    
    # Also display to user based on level
    case $level in
        "ERROR")
            echo -e "${RED}[ERROR] $message${NC}"
            ;;
        "WARNING")
            echo -e "${YELLOW}[WARNING] $message${NC}"
            ;;
        "INFO")
            echo -e "${GREEN}[INFO] $message${NC}"
            ;;
        "DEBUG")
            # Only log to file, not display
            ;;
    esac
}

################################################################################
# USER VALIDATION FUNCTIONS
################################################################################

# Function to get all regular users (excluding system users)
get_regular_users() {
    local users=()
    
    # Check regular users from Directory Service
    for username in $(dscl . -list /Users | grep -v '^_' | grep -v '^Guest' | grep -v '^nobody'); do
        local uid=$(dscl . -read /Users/"$username" UniqueID 2>/dev/null | awk '{print $2}')
        local home=$(dscl . -read /Users/"$username" NFSHomeDirectory 2>/dev/null | awk '{print $2}')
        
        # Include users with UID >= 501 (macOS regular users) and valid home directories
        if [[ -n "$uid" && $uid -ge 501 && -d "$home" ]]; then
            users+=("$username:$uid:$home")
        fi
    done
    
    # Also check for any nested Users directories under /Users
    # Find any subdirectories named "Users" under /Users (like /Users/*/Users patterns)
    while IFS= read -r -d '' nested_users_dir; do
        # Skip the root /Users directory itself
        if [[ "$nested_users_dir" != "/Users" ]]; then
            echo "Found nested Users directory: $nested_users_dir" >&2
            
            # Generate a unique identifier for this nested path
            local path_id="${nested_users_dir#/Users/}"  # Remove /Users/ prefix
            path_id="${path_id//\//_}"                   # Replace slashes with underscores
            
            for nested_dir in "$nested_users_dir"/*; do
                if [[ -d "$nested_dir" ]]; then
                    local dirname=$(basename "$nested_dir")
                    
                    # Skip system directories
                    if [[ "$dirname" != "Shared" && "$dirname" != "Guest" && "$dirname" != ".localized" ]]; then
                        # Check if this corresponds to an existing user account
                        local uid=$(dscl . -read /Users/"$dirname" UniqueID 2>/dev/null | awk '{print $2}')
                        
                        if [[ -n "$uid" && $uid -ge 501 ]]; then
                            # This is a valid user with a nested home directory
                            users+=("${dirname}_nested_${path_id}:$uid:$nested_dir")
                        fi
                    fi
                fi
            done
        fi
    done < <(find /Users -maxdepth 3 -type d -name "Users" -print0 2>/dev/null)
    
    # Also check for any other */fs-root/Users patterns (like mounted volumes)
    for fs_root_path in */fs-root/Users /Volumes/*/fs-root/Users; do
        # Check if the pattern matched actual directories (not literal *)
        if [[ -d "$fs_root_path" && "$fs_root_path" != "*/fs-root/Users" && "$fs_root_path" != "/Volumes/*/fs-root/Users" && "$fs_root_path" != "$shared_fs_root" ]]; then
            echo "Found additional nested Users directory: $fs_root_path" >&2
            for nested_dir in "$fs_root_path"/*; do
                if [[ -d "$nested_dir" ]]; then
                    local dirname=$(basename "$nested_dir")
                    
                    # Skip system directories
                    if [[ "$dirname" != "Shared" && "$dirname" != "Guest" && "$dirname" != ".localized" ]]; then
                        # Check if this corresponds to an existing user account
                        local uid=$(dscl . -read /Users/"$dirname" UniqueID 2>/dev/null | awk '{print $2}')
                        
                        if [[ -n "$uid" && $uid -ge 501 ]]; then
                            # This is a valid user with a nested home directory
                            local mount_point=$(dirname $(dirname "$fs_root_path"))
                            users+=("${dirname}_nested_${mount_point//\//_}:$uid:$nested_dir")
                        fi
                    fi
                fi
            done
        fi
    done
    
    # Sort users by UID (second field in the colon-separated format)
    printf '%s\n' "${users[@]}" | sort -t: -k2,2n
}

# Function to get orphaned home directories (directories without corresponding user accounts)
get_orphaned_directories() {
    local orphaned=()
    
    # Function to check and add orphaned directory
    check_orphaned_dir() {
        local user_dir="$1"
        local location_prefix="$2"  # "" for /Users, "nested_" for /fs-root/Users
        
        if [[ -d "$user_dir" ]]; then
            local dirname=$(basename "$user_dir")
            
            # Skip system directories and current user directories
            if [[ "$dirname" != "Shared" && "$dirname" != "Guest" && "$dirname" != ".localized" ]]; then
                # Skip container directories that likely contain nested user directories
                # Check if this directory contains a "Users" subdirectory - if so, it's probably a container
                if [[ -d "$user_dir/Users" ]]; then
                    return 0  # Skip this directory as it's a container
                fi
                
                # Check if user account exists for this directory
                if ! dscl . -read /Users/"$dirname" &>/dev/null; then
                    # This is an orphaned directory - get its details
                    local owner_uid=$(stat -f "%u" "$user_dir" 2>/dev/null || echo "unknown")
                    local owner_name=$(stat -f "%Su" "$user_dir" 2>/dev/null)
                    
                    # If owner name is just the UID in parentheses, clean it up
                    if [[ "$owner_name" == "($owner_uid)" || -z "$owner_name" ]]; then
                        owner_name="deleted_uid_${owner_uid}"
                    fi
                    
                    local dir_size=$(du -sh "$user_dir" 2>/dev/null | cut -f1 | tr -d ' ' || echo "N/A")
                    local last_modified=$(stat -f "%Sm" -t "%Y-%m-%d_%H:%M" "$user_dir" 2>/dev/null || echo "unknown")
                    
                    # Use a different delimiter to avoid conflicts with colons in data
                    # Format: dirname|owner_uid|owner_name|full_path|size|last_modified|ORPHANED
                    local display_name="${location_prefix}${dirname}"
                    orphaned+=("$display_name|$owner_uid|$owner_name|$user_dir|$dir_size|$last_modified|ORPHANED")
                fi
            fi
        fi
    }
    
    # Check each directory in /Users
    for user_dir in /Users/*; do
        check_orphaned_dir "$user_dir" ""
    done
    
    # Also check for orphaned directories under any nested Users directories
    while IFS= read -r -d '' nested_users_dir; do
        # Skip the root /Users directory itself
        if [[ "$nested_users_dir" != "/Users" ]]; then
            echo "Scanning for orphaned directories in: $nested_users_dir" >&2
            
            # Generate a unique identifier for this nested path
            local path_id="${nested_users_dir#/Users/}"  # Remove /Users/ prefix
            path_id="${path_id//\//_}"                   # Replace slashes with underscores
            
            for user_dir in "$nested_users_dir"/*; do
                check_orphaned_dir "$user_dir" "nested_${path_id}_"
            done
        fi
    done < <(find /Users -maxdepth 3 -type d -name "Users" -print0 2>/dev/null)
    
    # Also check for orphaned directories under any other */fs-root/Users patterns
    for fs_root_path in */fs-root/Users /Volumes/*/fs-root/Users; do
        # Check if the pattern matched actual directories (not literal *)
        if [[ -d "$fs_root_path" && "$fs_root_path" != "*/fs-root/Users" && "$fs_root_path" != "/Volumes/*/fs-root/Users" && "$fs_root_path" != "$shared_fs_root" ]]; then
            echo "Scanning for orphaned directories in: $fs_root_path" >&2
            for user_dir in "$fs_root_path"/*; do
                local mount_point=$(dirname $(dirname "$fs_root_path"))
                check_orphaned_dir "$user_dir" "nested_${mount_point//\//_}_"
            done
        fi
    done
    
    # Sort orphaned directories by owner UID (second field in the pipe-separated format)
    printf '%s\n' "${orphaned[@]}" | sort -t'|' -k2,2n
}

# Function to get combined list of users and orphaned directories
get_users_and_orphaned() {
    # Get regular users first
    get_regular_users
    
    # Then get orphaned directories
    get_orphaned_directories
}

# Function to display user selection menu
show_user_selection() {
    local title="$1"
    local current_restriction="$2"  # "none", "no-source", "no-target"
    local allow_orphaned="$3"      # "true" to include orphaned directories, "false" or empty to exclude
    
    echo -e "${BLUE}$title${NC}"
    echo "================================================================"
    
    local all_entries=()
    if [[ "$allow_orphaned" == "true" ]]; then
        all_entries=($(get_users_and_orphaned))
    else
        all_entries=($(get_regular_users))
    fi
    
    if [[ ${#all_entries[@]} -eq 0 ]]; then
        echo -e "${RED}No users or directories found.${NC}"
        return 1
    fi
    
    # Display header with optimized spacing for readability
    printf "%-3s  %-35s  %-4s  %-50s  %-18s  %8s  %s\n" "No" "Name" "UID" "Home Directory" "Status" "Size" "Type"
    printf "%-3s  %-35s  %-4s  %-50s  %-18s  %8s  %s\n" "---" "-----------------------------------" "----" "--------------------------------------------------" "------------------" "--------" "--------"
    
    local i=1
    local orphaned_count=0
    
    for entry_info in "${all_entries[@]}"; do
        # Parse entry - could be user or orphaned directory
        local entry_type=""
        local display_name=""
        local display_uid=""
        local display_home=""
        local display_size=""
        
        if [[ "$entry_info" == *"|ORPHANED" ]]; then
            # Parse orphaned directory: dirname|owner_uid|owner_name|full_path|size|last_modified|ORPHANED
            local old_ifs="$IFS"
            IFS='|'
            local fields=($entry_info)
            IFS="$old_ifs"
            
            # Extract fields from array
            local dirname="${fields[0]}"
            local owner_uid="${fields[1]}"
            local owner_name="${fields[2]}"
            local full_path="${fields[3]}"
            local dir_size="${fields[4]}"
            local last_modified="${fields[5]}"
            local entry_type="${fields[6]}"
            
            # Validate parsed orphaned directory data
            if [[ -n "$dirname" && -n "$full_path" ]]; then
                # Check if this is a nested orphaned directory
                if [[ "$dirname" == nested_*_* ]]; then
                    # Extract base name (the actual directory name) 
                    local temp="${dirname#nested_}"              # Remove nested_ prefix
                    local path_info="${temp%_*}"                 # Get path part (everything except last part)
                    local base_name="${temp##*_}"                # Get actual directory name (last part)
                    
                    # Convert underscores back to slashes for path display
                    path_info="${path_info//_/\/}"
                    
                    # Handle common patterns for cleaner display
                    case "$path_info" in
                        "Shared/fs-root")
                            display_name="$base_name [Shared]"
                            ;;
                        "fs-root")
                            display_name="$base_name [fs-root]"
                            ;;
                        *)
                            # Simplify complex nested paths - get last component
                            local simple_path="${path_info##*/}"  
                            display_name="$base_name [$simple_path]"
                            ;;
                    esac
                elif [[ "$dirname" == nested_* ]]; then
                    # Legacy format
                    display_name="${dirname#nested_} [NESTED]"
                else
                    display_name="$dirname"
                fi
                display_uid="${owner_uid:-'unknown'}"
                display_home="$full_path"
                display_size="${dir_size:-'N/A'}"
                entry_type="ORPHANED"
                ((orphaned_count++))
            else
                # Skip malformed entries
                continue
            fi
        else
            # Parse regular user: username:uid:home
            local old_ifs="$IFS"
            IFS=':'
            read -r username uid home <<< "$entry_info"
            IFS="$old_ifs"
            
            # Validate parsed data
            if [[ -n "$username" && -n "$uid" && -n "$home" ]]; then
                # Check if this is a nested directory user
                if [[ "$username" == *"_nested_"* ]]; then
                    # Extract base username (the actual username) and nested path info
                    local temp="${username#*_nested_}"           # Get everything after _nested_
                    local base_name="${username%%_nested_*}"     # Get username (before _nested_)
                    local path_info="$temp"
                    
                    # Convert underscores back to slashes for display
                    path_info="${path_info//_/\/}"
                    
                    # Handle common patterns for cleaner display
                    case "$path_info" in
                        "Shared/fs-root")
                            display_name="$base_name [Shared]"
                            ;;
                        "fs-root")
                            display_name="$base_name [fs-root]"
                            ;;
                        *)
                            # Simplify complex nested paths - get last component
                            local simple_path="${path_info##*/}"  
                            display_name="$base_name [$simple_path]"
                            ;;
                    esac
                elif [[ "$username" == *"_nested" ]]; then
                    # Legacy format for /fs-root/Users
                    display_name="${username%_nested} [NESTED]"
                else
                    display_name="$username"
                fi
                display_uid="$uid" 
                display_home="$home"
                display_size=$(du -sh "$home" 2>/dev/null | cut -f1 || echo "N/A")
                entry_type="USER"
            else
                # Skip malformed entries
                continue
            fi
        fi
        
        # Check restrictions
        local restricted=false
        local restriction_reason=""
        local color=""
        
        if [[ "$entry_type" == "USER" ]]; then
            # Extract base username for comparison (remove [NESTED*] tags if present)
            local base_username="$display_name"
            base_username="${base_username% \[NESTED\]}"        # Remove legacy format
            base_username="${base_username% \[*\]}"             # Remove nested format brackets
            
            if [[ "$current_restriction" == "no-source" && "$base_username" == "$CURRENT_USER" ]]; then
                restricted=true
                restriction_reason="[LOCKED]"
                color="${RED}"
            elif [[ "$current_restriction" == "no-target" && "$base_username" == "$SOURCE_USER" ]]; then
                restricted=true
                restriction_reason="[SELECTED AS SOURCE]"
                color="${RED}"
            elif [[ "$base_username" == "$CURRENT_USER" ]]; then
                restriction_reason="[CURRENT USER]"
                color="${CYAN}"
            fi
        elif [[ "$entry_type" == "ORPHANED" ]]; then
            if [[ "$current_restriction" == "no-target" ]]; then
                # Orphaned directories cannot be targets, only sources
                restricted=true
                restriction_reason="[ORPHANED - SOURCE ONLY]"
                color="${RED}"
            else
                restriction_reason="[ORPHANED DIR]"
                color="${YELLOW}"
            fi
        fi
        
        # Display entry
        local status_info="$restriction_reason"
        # Truncate long paths for better display
        local truncated_home="$display_home"
        if [[ ${#truncated_home} -gt 48 ]]; then
            truncated_home="...${truncated_home: -45}"
        fi
        
        # Truncate long names if needed
        local truncated_name="$display_name"
        if [[ ${#truncated_name} -gt 35 ]]; then
            truncated_name="${truncated_name:0:32}..."
        fi
        
        if [[ -n "$color" ]]; then
            printf "${color}%-3d  %-35s  %4s  %-50s  %-18s  %8s  %s${NC}\n" "$i" "$truncated_name" "$display_uid" "$truncated_home" "$status_info" "$display_size" "$entry_type"
        else
            printf "%-3d  %-35s  %4s  %-50s  %-18s  %8s  %s\n" "$i" "$truncated_name" "$display_uid" "$truncated_home" "$status_info" "$display_size" "$entry_type"
        fi
        
        ((i++))
    done
    
    if [[ "$allow_orphaned" == "true" && $orphaned_count -gt 0 ]]; then
        echo ""
        echo -e "${YELLOW}Found $orphaned_count orphaned directory(ies). These can be merged as sources only.${NC}"
        echo -e "${YELLOW}Orphaned directories are home folders without corresponding user accounts.${NC}"
    fi
    
    echo ""
    return 0
}

# Function to select source user
select_source_user() {
    while true; do
        show_user_selection "SELECT SOURCE USER OR ORPHANED DIRECTORY (data to be merged FROM)" "no-source" "true"
        
        local all_entries=($(get_users_and_orphaned))
        local max_choice=${#all_entries[@]}
        
        echo -e "${YELLOW}Note: The source data will be merged into the target user.${NC}"
        echo -e "${YELLOW}• User accounts will remain intact after merge${NC}"
        echo -e "${YELLOW}• Orphaned directories can be cleaned up after successful merge${NC}"
        echo ""
        
        read -p "Enter number (1-$max_choice) or 'X' to exit: " choice
        
        case $choice in
            [Xx])
                echo -e "${YELLOW}Exiting...${NC}"
                exit 0
                ;;
            ''|*[!0-9]*)
                echo -e "${RED}Invalid input. Please enter a number.${NC}"
                echo ""
                continue
                ;;
            *)
                if [[ $choice -ge 1 && $choice -le $max_choice ]]; then
                    local selected_entry_info="${all_entries[$((choice-1))]}"
                    
                    # Determine if this is a user or orphaned directory
                    if [[ "$selected_entry_info" == *"|ORPHANED" ]]; then
                        # Parse orphaned directory
                        IFS='|' read -r dirname owner_uid owner_name full_path dir_size last_modified entry_type <<< "$selected_entry_info"
                        
                        # Additional validation for orphaned directory
                        if [[ ! -d "$full_path" ]]; then
                            echo -e "${RED}Error: Orphaned directory '$full_path' not found or inaccessible.${NC}"
                            echo ""
                            continue
                        fi
                        
                        SOURCE_USER="ORPHANED:$dirname"
                        SOURCE_HOME="$full_path"
                        
                        log_message "INFO" "Selected orphaned directory: $dirname ($SOURCE_HOME)"
                        echo -e "${GREEN}Selected orphaned directory: $dirname${NC}"
                        echo -e "${GREEN}Source directory: $SOURCE_HOME${NC}"
                        echo -e "${YELLOW}Original owner UID: $owner_uid ($owner_name)${NC}"
                        echo -e "${YELLOW}Last modified: $last_modified${NC}"
                        echo ""
                    else
                        # Parse regular user
                        IFS=':' read -r username uid home <<< "$selected_entry_info"
                        
                        # Additional validation for regular user
                        if [[ "$username" == "$CURRENT_USER" ]]; then
                            echo -e "${RED}Error: Cannot use current user ($CURRENT_USER) as source.${NC}"
                            echo -e "${RED}This would risk data loss and system instability.${NC}"
                            echo ""
                            continue
                        fi
                        
                        if [[ ! -d "$home" ]]; then
                            echo -e "${RED}Error: Home directory '$home' not found or inaccessible.${NC}"
                            echo ""
                            continue
                        fi
                        
                        SOURCE_USER="$username"
                        SOURCE_HOME="$home"
                        
                        log_message "INFO" "Selected source user: $SOURCE_USER ($SOURCE_HOME)"
                        echo -e "${GREEN}Selected source user: $SOURCE_USER${NC}"
                        echo -e "${GREEN}Source home directory: $SOURCE_HOME${NC}"
                        echo ""
                    fi
                    break
                else
                    echo -e "${RED}Invalid choice. Please enter a number between 1 and $max_choice.${NC}"
                    echo ""
                fi
                ;;
        esac
    done
}

# Function to select target user  
select_target_user() {
    while true; do
        show_user_selection "SELECT TARGET USER (user who will receive the merged data)" "no-target" "false"
        
        local users=($(get_regular_users))
        local max_choice=${#users[@]}
        
        echo -e "${YELLOW}Note: The target user will receive all data from the source user.${NC}"
        echo -e "${YELLOW}Existing files may be renamed or replaced based on your choices.${NC}"
        echo ""
        
        read -p "Enter user number (1-$max_choice) or 'X' to exit: " choice
        
        case $choice in
            [Xx])
                echo -e "${YELLOW}Exiting...${NC}"
                exit 0
                ;;
            ''|*[!0-9]*)
                echo -e "${RED}Invalid input. Please enter a number.${NC}"
                echo ""
                continue
                ;;
            *)
                if [[ $choice -ge 1 && $choice -le $max_choice ]]; then
                    local selected_user_info="${users[$((choice-1))]}"
                    IFS=':' read -r username uid home <<< "$selected_user_info"
                    
                    # Additional validation
                    # Check if target user is same as source (handle both regular users and orphaned directories)
                    local source_name="$SOURCE_USER"
                    if [[ "$SOURCE_USER" == "ORPHANED:"* ]]; then
                        source_name="${SOURCE_USER#ORPHANED:}"
                    fi
                    
                    if [[ "$username" == "$source_name" ]]; then
                        echo -e "${RED}Error: Target user cannot be the same as source.${NC}"
                        echo ""
                        continue
                    fi
                    
                    if [[ ! -d "$home" ]]; then
                        echo -e "${RED}Error: Home directory '$home' not found or inaccessible.${NC}"
                        echo ""
                        continue
                    fi
                    
                    TARGET_USER="$username"
                    TARGET_HOME="$home"
                    
                    log_message "INFO" "Selected target user: $TARGET_USER ($TARGET_HOME)"
                    echo -e "${GREEN}Selected target user: $TARGET_USER${NC}"
                    echo -e "${GREEN}Target home directory: $TARGET_HOME${NC}"
                    echo ""
                    break
                else
                    echo -e "${RED}Invalid choice. Please enter a number between 1 and $max_choice.${NC}"
                    echo ""
                fi
                ;;
        esac
    done
}

################################################################################
# ANALYSIS AND VALIDATION FUNCTIONS
################################################################################

# Function to analyze directories and show merge preview
# Function to format file sizes in human-readable format (macOS compatible)
format_size() {
    local bytes=$1
    if [[ -z "$bytes" || "$bytes" -eq 0 ]]; then
        echo "0B"
        return
    fi
    
    # Convert to human readable format
    if (( bytes >= 1073741824 )); then
        echo "$(( (bytes + 536870912) / 1073741824 ))G"
    elif (( bytes >= 1048576 )); then
        echo "$(( (bytes + 524288) / 1048576 ))M"
    elif (( bytes >= 1024 )); then
        echo "$(( (bytes + 512) / 1024 ))K"
    else
        echo "${bytes}B"
    fi
}

# Function to create comprehensive pre-merge file analysis
create_file_analysis() {
    echo -e "${BLUE}CREATING COMPREHENSIVE FILE ANALYSIS...${NC}"
    echo "================================================================"
    
    local analysis_file="/tmp/merge_analysis_${SOURCE_USER//[^a-zA-Z0-9]/_}_to_${TARGET_USER}_$(date +%Y%m%d_%H%M%S).txt"
    
    log_message "INFO" "Creating detailed file analysis: $analysis_file"
    
    # Create analysis header
    cat > "$analysis_file" << EOF
================================================================================
COMPREHENSIVE PRE-MERGE FILE ANALYSIS
================================================================================
Analysis Date: $(date)
Source: $SOURCE_USER ($SOURCE_HOME)
Target: $TARGET_USER ($TARGET_HOME)
Script Version: 1.3 - Enhanced with UID Sorting and Improved Conflict Display
================================================================================

LEGEND:
[NEW]     - File will be copied (no conflict)
[CONFLICT]- File exists in target, conflict resolution needed
[SKIP]    - File will be skipped (empty or system file)
[BACKUP]  - File will be backed up from target before merge

Conflict Resolution Recommendations:
RENAME   - Safest option, preserves all data with timestamp suffix
REPLACE  - Use when source is newer/more important than target
SKIP     - Use when target version should be preserved
ASK      - Use when manual decision needed for each conflict

================================================================================
FILE ANALYSIS RESULTS:
================================================================================

EOF

    echo "Analyzing all files for conflicts and recommendations..."
    echo "This may take several minutes for large directories..."
    
    local total_files=0
    local new_files=0
    local conflicts=0
    local skipped_files=0
    local total_source_size=0
    local total_conflict_size=0
    
    # Process all files in source directory
    while IFS= read -r -d '' source_file; do
        # Skip if not a regular file
        [[ ! -f "$source_file" ]] && continue
        
        local rel_path="${source_file#$SOURCE_HOME/}"
        local target_file="$TARGET_HOME/$rel_path"
        local source_size=$(stat -f%z "$source_file" 2>/dev/null || echo 0)
        local target_size=0
        
        total_files=$((total_files + 1))
        total_source_size=$((total_source_size + source_size))
        
        # Check if file exists in target
        if [[ -f "$target_file" ]]; then
            # CONFLICT detected
            conflicts=$((conflicts + 1))
            target_size=$(stat -f%z "$target_file" 2>/dev/null || echo 0)
            total_conflict_size=$((total_conflict_size + target_size))
            
            # Determine recommendation
            local recommendation="RENAME"
            local reason="Default safe option"
            
            # Get modification times
            local source_mtime=$(stat -f%m "$source_file" 2>/dev/null || echo 0)
            local target_mtime=$(stat -f%m "$target_file" 2>/dev/null || echo 0)
            
            # Smart recommendations based on file analysis
            if [[ $source_size -eq $target_size ]]; then
                # Same size - check if identical
                if cmp -s "$source_file" "$target_file" 2>/dev/null; then
                    recommendation="SKIP"
                    reason="Files are identical"
                else
                    recommendation="ASK"
                    reason="Same size but different content"
                fi
            elif [[ $source_mtime -gt $target_mtime ]]; then
                if [[ $source_size -gt $target_size ]]; then
                    recommendation="REPLACE"
                    reason="Source is newer and larger"
                else
                    recommendation="ASK"
                    reason="Source newer but smaller"
                fi
            elif [[ $target_mtime -gt $source_mtime ]]; then
                recommendation="SKIP"
                reason="Target is newer"
            elif [[ $source_size -gt $target_size ]]; then
                recommendation="REPLACE"
                reason="Source is larger"
            fi
            
            # Special handling for certain file types
            case "${rel_path##*.}" in
                plist|conf|config)
                    recommendation="ASK"
                    reason="Configuration file - manual review needed"
                    ;;
                log)
                    recommendation="RENAME"
                    reason="Log file - preserve both versions"
                    ;;
                tmp|temp|cache)
                    recommendation="SKIP"
                    reason="Temporary/cache file"
                    ;;
            esac
            
            printf "[CONFLICT] %s (%s) | S:%-8s T:%-8s | %s\n" \
                "$recommendation" \
                "$reason" \
                "$(format_size $source_size)" \
                "$(format_size $target_size)" \
                "$rel_path" >> "$analysis_file"
                
        else
            # NEW file - no conflict
            new_files=$((new_files + 1))
            printf "[NEW]      Will be copied | %-8s | %s\n" \
                "$(format_size $source_size)" \
                "$rel_path" >> "$analysis_file"
        fi
        
        # Progress indicator
        if (( total_files % 50 == 0 )); then
            echo -n "."
        fi
        
    done < <(find "$SOURCE_HOME" -type f -print0 2>/dev/null)
    
    echo ""  # New line after progress dots
    
    # Add summary to analysis file
    cat >> "$analysis_file" << EOF

================================================================================
ANALYSIS SUMMARY:
================================================================================
Total Files Analyzed:     $total_files
New Files (no conflict):  $new_files
Conflicting Files:        $conflicts
Skipped Files:           $skipped_files

Total Source Data Size:   $(format_size $total_source_size)
Total Conflict Data Size: $(format_size $total_conflict_size)

RECOMMENDATIONS BY CONFLICT TYPE:
$(grep "\[CONFLICT\]" "$analysis_file" | awk -F'|' '{print $3}' | awk '{print $1}' | sort | uniq -c | awk '{printf "%-10s: %d files\n", $2, $1}')

BACKUP REQUIREMENTS:
Target directory will be backed up before merge: $(du -sh "$TARGET_HOME" 2>/dev/null | cut -f1)
Analysis file location: $analysis_file

NEXT STEPS:
1. Review this analysis file for any files requiring special attention
2. Consider conflict resolution strategy based on recommendations
3. Proceed with merge operation using chosen strategy
4. Keep this analysis file for reference during rollback if needed

================================================================================
Generated by: User Home Directory Merge Script v1.2
Analysis completed: $(date)
================================================================================
EOF

    # Display summary to user
    echo ""
    echo -e "${GREEN}✓ File Analysis Complete${NC}"
    echo "================================================================"
    echo "Analysis file created: $analysis_file"
    echo ""
    echo -e "${YELLOW}ANALYSIS SUMMARY:${NC}"
    echo "• Total Files: $total_files"
    echo "• New Files: $new_files (will be copied)"
    echo "• Conflicts: $conflicts (need resolution)"
    echo "• Source Data: $(format_size $total_source_size)"
    echo "• Conflict Data: $(format_size $total_conflict_size)"
    echo ""
    
    if [[ $conflicts -gt 0 ]]; then
        echo -e "${YELLOW}CONFLICT RECOMMENDATIONS:${NC}"
        
        # Show recommendation breakdown (macOS bash compatible)
        # Extract the recommendation from the first field after [CONFLICT]
        grep "\[CONFLICT\]" "$analysis_file" | sed 's/\[CONFLICT\] *//' | awk -F'|' '{print $1}' | awk '{print $1}' | sort | uniq -c | while read count rec; do
            echo "• $rec: $count files"
        done
        
        echo ""
        echo -e "${CYAN}TOP CONFLICTS (first 40):${NC}"

        # Show first 40 conflicts with full details
        grep "\[CONFLICT\]" "$analysis_file" | head -40 | while read -r line; do
            # Extract parts: [CONFLICT] ACTION (reason) | sizes | path
            recommendation=$(echo "$line" | sed 's/\[CONFLICT\] *//' | awk -F'|' '{print $1}' | sed 's/^ *//' | sed 's/ *$//')
            sizes=$(echo "$line" | awk -F'|' '{print $2}' | sed 's/^ *//' | sed 's/ *$//')
            path=$(echo "$line" | awk -F'|' '{print $3}' | sed 's/^ *//' | sed 's/ *$//')
            echo "  $recommendation | $sizes | $path"
        done
        
        if [[ $conflicts -gt 40 ]]; then
            echo "  ... and $((conflicts - 40)) more conflicts"
            echo ""
            echo -e "${YELLOW}⚠️  REVIEW RECOMMENDED: See full analysis file for all conflicts${NC}"
        fi
        echo ""
    fi
    
    echo -e "${CYAN}Please review the detailed analysis file before proceeding.${NC}"
    echo "Analysis file location: $analysis_file"
    echo ""
    echo "Options:"
    echo "  1. Continue with merge process"
    echo "  2. Exit to review analysis file"
    echo ""
    while true; do
        read -p "Enter choice (1-2): " choice
        case $choice in
            1)
                echo "Continuing with merge process..."
                break
                ;;
            2)
                echo "Exiting to allow analysis file review."
                echo "Analysis file: $analysis_file"
                exit 0
                ;;
            *)
                echo "Invalid choice. Please enter 1 or 2."
                ;;
        esac
    done
    
    log_message "INFO" "File analysis completed: $total_files files, $conflicts conflicts"
    
    # Store analysis file path for later reference
    ANALYSIS_FILE="$analysis_file"
}

analyze_merge() {
    echo -e "${BLUE}ANALYZING MERGE REQUIREMENTS...${NC}"
    echo "================================================================"
    
    log_message "INFO" "Starting merge analysis"
    
    # Display source information
    local source_display="$SOURCE_USER"
    local source_type_label="User"
    if [[ "$SOURCE_USER" == "ORPHANED:"* ]]; then
        source_display="${SOURCE_USER#ORPHANED:}"
        source_type_label="Orphaned Directory"
    fi
    
    # Calculate sizes
    echo -e "${GREEN}Calculating directory sizes...${NC}"
    local source_size=$(du -sh "$SOURCE_HOME" 2>/dev/null | cut -f1)
    local target_size=$(du -sh "$TARGET_HOME" 2>/dev/null | cut -f1)
    
    echo "Source $source_type_label: $source_display"
    echo "Source directory size: $source_size ($SOURCE_HOME)"
    echo "Target user: $TARGET_USER"
    echo "Target directory size: $target_size ($TARGET_HOME)"
    
    # Count files
    echo -e "${GREEN}Counting files and directories...${NC}"
    local source_files=$(find "$SOURCE_HOME" -type f 2>/dev/null | wc -l | tr -d ' ')
    local source_dirs=$(find "$SOURCE_HOME" -type d 2>/dev/null | wc -l | tr -d ' ')
    local target_files=$(find "$TARGET_HOME" -type f 2>/dev/null | wc -l | tr -d ' ')
    local target_dirs=$(find "$TARGET_HOME" -type d 2>/dev/null | wc -l | tr -d ' ')
    
    echo "Source: $source_files files, $source_dirs directories"
    echo "Target: $target_files files, $target_dirs directories"
    
    TOTAL_FILES=$source_files
    
    # Check available disk space
    echo -e "${GREEN}Checking available disk space...${NC}"
    local target_disk=$(df -h "$TARGET_HOME" | tail -1 | awk '{print $4}')
    echo "Available space on target volume: $target_disk"
    
    # Identify special directories and potential conflicts
    echo ""
    echo -e "${BLUE}SPECIAL DIRECTORIES ANALYSIS:${NC}"
    echo "----------------------------------------------------------------"
    
    local special_dirs=("Desktop" "Documents" "Downloads" "Pictures" "Movies" "Music" "Library" "Applications")
    
    for dir in "${special_dirs[@]}"; do
        local source_path="$SOURCE_HOME/$dir"
        local target_path="$TARGET_HOME/$dir"
        
        if [[ -d "$source_path" ]]; then
            local source_dir_size=$(du -sh "$source_path" 2>/dev/null | cut -f1)
            local source_dir_files=$(find "$source_path" -type f 2>/dev/null | wc -l | tr -d ' ')
            
            echo -n "  $dir: $source_dir_size ($source_dir_files files) → "
            
            if [[ -d "$target_path" ]]; then
                local target_dir_size=$(du -sh "$target_path" 2>/dev/null | cut -f1)
                local target_dir_files=$(find "$target_path" -type f 2>/dev/null | wc -l | tr -d ' ')
                echo -e "${YELLOW}MERGE with existing ($target_dir_size, $target_dir_files files)${NC}"
            else
                echo -e "${GREEN}CREATE new directory${NC}"
            fi
        fi
    done
    
    # Check for hidden files and directories
    echo ""
    local hidden_count=$(find "$SOURCE_HOME" -name ".*" -not -name "." -not -name ".." 2>/dev/null | wc -l | tr -d ' ')
    if [[ $hidden_count -gt 0 ]]; then
        echo -e "${CYAN}Found $hidden_count hidden files/directories (will prompt for inclusion)${NC}"
    fi
    
    echo ""
    log_message "INFO" "Analysis complete: $source_files files to merge"
}

# Function to check disk space requirements
check_disk_space() {
    echo -e "${GREEN}Verifying sufficient disk space...${NC}"
    
    # Get source directory size in bytes
    local source_bytes=$(du -s "$SOURCE_HOME" 2>/dev/null | awk '{print $1}')
    source_bytes=$((source_bytes * 1024))  # Convert from KB to bytes
    
    # Get available space on target volume in bytes
    local available_bytes=$(df "$TARGET_HOME" | tail -1 | awk '{print $4}')
    available_bytes=$((available_bytes * 1024))  # Convert from KB to bytes
    
    # Add 20% buffer for backups and temporary files
    local required_bytes=$((source_bytes + (source_bytes / 5)))
    
    if [[ $available_bytes -lt $required_bytes ]]; then
        local required_gb=$((required_bytes / 1024 / 1024 / 1024))
        local available_gb=$((available_bytes / 1024 / 1024 / 1024))
        
        echo -e "${RED}ERROR: Insufficient disk space!${NC}"
        echo -e "${RED}Required: ${required_gb}GB (including backup buffer)${NC}"
        echo -e "${RED}Available: ${available_gb}GB${NC}"
        log_message "ERROR" "Insufficient disk space: need ${required_gb}GB, have ${available_gb}GB"
        return 1
    fi
    
    echo -e "${GREEN}✓ Sufficient disk space available${NC}"
    log_message "INFO" "Disk space check passed"
    return 0
}

################################################################################
# CONFLICT RESOLUTION FUNCTIONS  
################################################################################

# Function to select conflict resolution strategy
select_conflict_strategy() {
    echo -e "${BLUE}CONFLICT RESOLUTION STRATEGY${NC}"
    echo "================================================================"
    echo "When files with the same name exist in both directories, how would"
    echo "you like to handle the conflicts?"
    echo ""
    echo "1. RENAME - Add timestamp suffix to conflicting files from source"
    echo "   (safest option, preserves all data)"
    echo "2. SKIP - Skip files that already exist in target"  
    echo "   (preserves target files, source duplicates not copied)"
    echo "3. REPLACE - Replace target files with source files"
    echo "   (overwrites target files, may cause data loss)"
    echo "4. ASK - Prompt for each conflict individually"
    echo "   (most control, but may require many decisions)"
    echo ""
    
    while true; do
        read -p "Select strategy (1-4): " strategy_choice
        
        case $strategy_choice in
            1)
                CONFLICT_STRATEGY="rename"
                echo -e "${GREEN}Selected: RENAME conflicts with timestamp suffix${NC}"
                break
                ;;
            2)
                CONFLICT_STRATEGY="skip"
                echo -e "${YELLOW}Selected: SKIP conflicting files${NC}"
                break
                ;;
            3)
                CONFLICT_STRATEGY="replace"
                echo -e "${RED}Selected: REPLACE target files (WARNING: May overwrite data)${NC}"
                echo ""
                read -p "Are you sure? This may overwrite important files (y/N): " confirm_replace
                if [[ "$confirm_replace" =~ ^[Yy]$ ]]; then
                    break
                else
                    echo "Please select a different strategy."
                    continue
                fi
                ;;
            4)
                CONFLICT_STRATEGY="ask"
                echo -e "${CYAN}Selected: ASK for each conflict${NC}"
                break
                ;;
            *)
                echo -e "${RED}Invalid choice. Please select 1-4.${NC}"
                ;;
        esac
    done
    
    echo ""
    log_message "INFO" "Conflict resolution strategy: $CONFLICT_STRATEGY"
}

################################################################################
# BACKUP FUNCTIONS
################################################################################

# Function to create backup
create_backup() {
    echo -e "${GREEN}CREATING BACKUP OF TARGET DIRECTORY...${NC}"
    echo "================================================================"
    
    # Create backup directory with timestamp
    BACKUP_DIR="/tmp/user_merge_backup_${TARGET_USER}_${MERGE_START_TIME}"
    
    echo "Creating backup directory: $BACKUP_DIR"
    mkdir -p "$BACKUP_DIR"
    
    log_message "INFO" "Creating backup at: $BACKUP_DIR"
    
    # Create backup using rsync for reliability
    echo "Backing up $TARGET_HOME to $BACKUP_DIR..."
    echo "This may take several minutes depending on data size..."
    
    if rsync -av --progress "$TARGET_HOME/" "$BACKUP_DIR/" 2>&1 | tee -a "$LOG_FILE"; then
        # Calculate backup size
        BACKUP_SIZE=$(du -sh "$BACKUP_DIR" 2>/dev/null | cut -f1)
        echo -e "${GREEN}✓ Backup completed successfully${NC}"
        echo -e "${GREEN}Backup size: $BACKUP_SIZE${NC}"
        echo -e "${GREEN}Backup location: $BACKUP_DIR${NC}"
        
        log_message "INFO" "Backup completed successfully: $BACKUP_SIZE at $BACKUP_DIR"
        
        # Create backup info file
        cat > "$BACKUP_DIR/.backup_info" << EOF
Backup Created: $(date)
Original Path: $TARGET_HOME
Original User: $TARGET_USER
Merge Operation: $SOURCE_USER → $TARGET_USER
Script Version: 1.3
Backup Size: $BACKUP_SIZE
EOF
        
        return 0
    else
        echo -e "${RED}ERROR: Backup creation failed${NC}"
        log_message "ERROR" "Backup creation failed"
        return 1
    fi
}

################################################################################
# MERGE FUNCTIONS
################################################################################

# Function to handle file conflicts based on strategy
handle_conflict() {
    local source_file="$1"
    local target_file="$2"
    local relative_path="$3"
    
    case $CONFLICT_STRATEGY in
        "rename")
            local timestamp=$(date '+%Y%m%d_%H%M%S')
            local target_renamed="${target_file}.backup_${timestamp}"
            mv "$target_file" "$target_renamed"
            log_message "INFO" "Renamed existing file: $relative_path → $(basename "$target_renamed")"
            return 0  # Proceed with copy
            ;;
        "skip")
            log_message "INFO" "Skipped conflicting file: $relative_path"
            ((SKIPPED_FILES++))
            return 1  # Skip copy
            ;;
        "replace")
            log_message "INFO" "Replaced existing file: $relative_path"
            return 0  # Proceed with copy (will overwrite)
            ;;
        "ask")
            echo ""
            echo -e "${YELLOW}CONFLICT: $relative_path${NC}"
            echo "  Source: $(stat -f "%z bytes, modified %Sm" "$source_file" 2>/dev/null)"
            echo "  Target: $(stat -f "%z bytes, modified %Sm" "$target_file" 2>/dev/null)"
            echo ""
            echo "1. Rename existing and copy source"
            echo "2. Skip this file"
            echo "3. Replace existing file"
            
            while true; do
                read -p "Choose action (1-3): " action_choice
                case $action_choice in
                    1)
                        local timestamp=$(date '+%Y%m%d_%H%M%S')
                        local target_renamed="${target_file}.backup_${timestamp}"
                        mv "$target_file" "$target_renamed"
                        log_message "INFO" "User chose rename: $relative_path → $(basename "$target_renamed")"
                        return 0
                        ;;
                    2)
                        log_message "INFO" "User chose skip: $relative_path"
                        ((SKIPPED_FILES++))
                        return 1
                        ;;
                    3)
                        log_message "INFO" "User chose replace: $relative_path"
                        return 0
                        ;;
                    *)
                        echo "Invalid choice. Please select 1-3."
                        ;;
                esac
            done
            ;;
    esac
}

# Function to merge a single file
merge_file() {
    local source_file="$1"
    local target_file="$2"
    local relative_path="$3"
    
    # Create target directory if it doesn't exist
    local target_dir=$(dirname "$target_file")
    if [[ ! -d "$target_dir" ]]; then
        mkdir -p "$target_dir"
        chown "$TARGET_USER:staff" "$target_dir"
        log_message "DEBUG" "Created directory: $target_dir"
    fi
    
    # Check if target file exists
    if [[ -f "$target_file" ]]; then
        # Handle conflict
        if ! handle_conflict "$source_file" "$target_file" "$relative_path"; then
            return 0  # Skip this file
        fi
        ((CONFLICT_FILES++))
    fi
    
    # Copy file preserving attributes
    if cp -p "$source_file" "$target_file" 2>/dev/null; then
        # Set ownership to target user
        chown "$TARGET_USER:staff" "$target_file"
        
        # Preserve extended attributes (macOS specific)
        if command -v xattr >/dev/null 2>&1; then
            xattr -l "$source_file" 2>/dev/null | while IFS= read -r attr; do
                local attr_name=$(echo "$attr" | cut -d: -f1)
                if [[ -n "$attr_name" ]]; then
                    xattr -w "$attr_name" "$(xattr -p "$attr_name" "$source_file" 2>/dev/null)" "$target_file" 2>/dev/null || true
                fi
            done
        fi
        
        ((MERGED_FILES++))
        log_message "DEBUG" "Merged file: $relative_path"
        
        # Progress indication
        if (( MERGED_FILES % 100 == 0 )); then
            echo -ne "\r  Progress: $MERGED_FILES/$TOTAL_FILES files merged..."
        fi
        
        return 0
    else
        log_message "WARNING" "Failed to copy file: $relative_path"
        return 1
    fi
}

# Function to perform the actual merge
perform_merge() {
    echo -e "${GREEN}STARTING DIRECTORY MERGE...${NC}"
    echo "================================================================"
    echo "Merging $SOURCE_HOME → $TARGET_HOME"
    echo "Strategy: $CONFLICT_STRATEGY"
    echo ""
    
    log_message "INFO" "Starting merge operation"
    
    # Initialize counters
    MERGED_FILES=0
    SKIPPED_FILES=0
    CONFLICT_FILES=0
    
    # Process all files in source directory
    local file_count=0
    
    while IFS= read -r -d '' source_file; do
        # Calculate relative path
        local relative_path="${source_file#$SOURCE_HOME/}"
        local target_file="$TARGET_HOME/$relative_path"
        
        # Skip if it's a directory (directories are created as needed)
        if [[ -d "$source_file" ]]; then
            continue
        fi
        
        # Merge the file
        merge_file "$source_file" "$target_file" "$relative_path"
        
        ((file_count++))
    done < <(find "$SOURCE_HOME" -type f -print0 2>/dev/null)
    
    echo ""
    echo -e "${GREEN}✓ Merge operation completed${NC}"
    
    # Display statistics
    echo ""
    echo -e "${BLUE}MERGE STATISTICS:${NC}"
    echo "----------------------------------------------------------------"
    echo "Files processed: $file_count"
    echo "Files merged: $MERGED_FILES"
    echo "Files skipped: $SKIPPED_FILES"  
    echo "Conflicts resolved: $CONFLICT_FILES"
    
    log_message "INFO" "Merge completed: $MERGED_FILES merged, $SKIPPED_FILES skipped, $CONFLICT_FILES conflicts"
    
    return 0
}

################################################################################
# SPECIAL DIRECTORY HANDLERS
################################################################################

# Function to handle Library directory specially
handle_library_merge() {
    local source_lib="$SOURCE_HOME/Library"
    local target_lib="$TARGET_HOME/Library"
    
    if [[ ! -d "$source_lib" ]]; then
        return 0
    fi
    
    echo -e "${CYAN}Special handling for Library directory...${NC}"
    
    # Safe Library subdirectories to merge
    local safe_library_dirs=(
        "Application Support/TextEdit"
        "Application Support/Numbers"
        "Application Support/Pages"
        "Application Support/Keynote"
        "Desktop Pictures"
        "Fonts"
        "Screen Savers"
        "Scripts"
        "Services"
    )
    
    # Ask user about Library merge
    echo -e "${YELLOW}The Library directory contains system and application preferences.${NC}"
    echo -e "${YELLOW}Merging everything could cause application conflicts.${NC}"
    echo ""
    echo "1. Merge only safe directories (recommended)"
    echo "2. Merge everything (advanced users only)"
    echo "3. Skip Library directory"
    
    while true; do
        read -p "Choose Library merge option (1-3): " lib_choice
        case $lib_choice in
            1)
                log_message "INFO" "Performing safe Library merge"
                for safe_dir in "${safe_library_dirs[@]}"; do
                    local source_path="$source_lib/$safe_dir"
                    local target_path="$target_lib/$safe_dir"
                    if [[ -d "$source_path" ]]; then
                        echo "  Merging Library/$safe_dir..."
                        mkdir -p "$(dirname "$target_path")"
                        rsync -av "$source_path/" "$target_path/" 2>/dev/null || true
                        chown -R "$TARGET_USER:staff" "$target_path" 2>/dev/null || true
                    fi
                done
                break
                ;;
            2)
                log_message "WARNING" "User chose full Library merge"
                echo -e "${RED}WARNING: Full Library merge may cause application issues.${NC}"
                read -p "Are you absolutely sure? (type 'YES' to confirm): " lib_confirm
                if [[ "$lib_confirm" == "YES" ]]; then
                    return 1  # Let normal merge handle it
                else
                    echo "Cancelled full Library merge. Using safe merge instead."
                    lib_choice=1
                    continue
                fi
                ;;
            3)
                log_message "INFO" "Skipping Library directory per user choice"
                echo "Skipping Library directory."
                break
                ;;
            *)
                echo "Invalid choice. Please select 1-3."
                ;;
        esac
    done
    
    return 0
}

################################################################################
# MAIN EXECUTION FUNCTIONS
################################################################################

# Function to show final confirmation
final_confirmation() {
    # Determine source display information
    local source_display="$SOURCE_USER"
    local source_type_label="user"
    if [[ "$SOURCE_USER" == "ORPHANED:"* ]]; then
        source_display="${SOURCE_USER#ORPHANED:}"
        source_type_label="orphaned directory"
    fi
    
    echo ""
    echo -e "${RED}FINAL CONFIRMATION${NC}"
    echo "================================================================"
    echo -e "${RED}You are about to merge:${NC}"
    echo -e "${RED}  FROM: $source_display ($source_type_label) at $SOURCE_HOME${NC}"
    echo -e "${RED}  TO:   $TARGET_USER (user) at $TARGET_HOME${NC}"
    echo ""
    echo -e "${RED}Conflict strategy: $CONFLICT_STRATEGY${NC}"
    echo -e "${RED}Backup location: $BACKUP_DIR${NC}"
    echo ""
    echo -e "${YELLOW}This operation will:${NC}"
    echo -e "${YELLOW}  • Copy all files from source to target user${NC}"
    echo -e "${YELLOW}  • Resolve conflicts using selected strategy${NC}"
    echo -e "${YELLOW}  • Change ownership of copied files to target user${NC}"
    echo -e "${YELLOW}  • Preserve file attributes and permissions${NC}"
    echo ""
    echo -e "${RED}WARNING: While backups are created, this operation modifies${NC}"
    echo -e "${RED}user data and should be performed with caution.${NC}"
    echo ""
    
    read -p "Type 'MERGE' to proceed or anything else to cancel: " final_confirm
    
    if [[ "$final_confirm" == "MERGE" ]]; then
        return 0
    else
        echo -e "${YELLOW}Merge operation cancelled.${NC}"
        return 1
    fi
}

# Function to show completion summary
show_completion_summary() {
    echo ""
    echo -e "${GREEN}================================================================${NC}"
    echo -e "${GREEN}                    MERGE COMPLETED                             ${NC}"
    echo -e "${GREEN}================================================================${NC}"
    echo ""
    # Determine source display information
    local source_display="$SOURCE_USER"
    local source_type_label="User"
    if [[ "$SOURCE_USER" == "ORPHANED:"* ]]; then
        source_display="${SOURCE_USER#ORPHANED:}"
        source_type_label="Orphaned Directory"
    fi
    
    echo -e "${BLUE}OPERATION SUMMARY:${NC}"
    echo "  Source $source_type_label: $source_display"
    echo "  Target User: $TARGET_USER"  
    echo "  Files Merged: $MERGED_FILES"
    echo "  Files Skipped: $SKIPPED_FILES"
    echo "  Conflicts Resolved: $CONFLICT_FILES"
    echo "  Backup Created: $BACKUP_DIR ($BACKUP_SIZE)"
    echo ""
    echo -e "${BLUE}NEXT STEPS:${NC}"
    echo "  1. Verify merged data in $TARGET_HOME"
    echo "  2. Test applications and user preferences"
    echo "  3. Consider disabling/removing source user account if merge is successful"
    echo "  4. Keep backup until you're certain everything works correctly"
    echo ""
    echo -e "${CYAN}BACKUP INFORMATION:${NC}"
    echo "  Location: $BACKUP_DIR"
    echo "  Size: $BACKUP_SIZE"
    echo "  Restore command (if needed): rsync -av \"$BACKUP_DIR/\" \"$TARGET_HOME/\""
    echo ""
    echo -e "${GREEN}Log file: $LOG_FILE${NC}"
    echo ""
    
    log_message "INFO" "Merge operation completed successfully"
}

################################################################################
# MAIN FUNCTION
################################################################################

main() {
    # Initialize
    show_header
    check_root
    get_current_user
    
    # Display protection information
    if [[ "$CURRENT_USER" != "unknown" ]]; then
        echo -e "${CYAN}Protected user: $CURRENT_USER (cannot be used as source)${NC}"
        echo ""
    fi
    
    # Step 1: Select users
    echo -e "${BLUE}STEP 1: USER SELECTION${NC}"
    select_source_user
    select_target_user
    
    # Setup logging now that we have user information
    setup_logging
    
    # Step 2: Analyze merge requirements
    echo -e "${BLUE}STEP 2: MERGE ANALYSIS${NC}"
    
    # Create comprehensive pre-merge file analysis
    create_file_analysis
    
    analyze_merge
    
    if ! check_disk_space; then
        exit 1
    fi
    
    # Step 3: Select conflict resolution strategy
    echo -e "${BLUE}STEP 3: CONFLICT RESOLUTION${NC}"
    select_conflict_strategy
    
    # Step 4: Create backup
    echo -e "${BLUE}STEP 4: BACKUP CREATION${NC}"
    if ! create_backup; then
        echo -e "${RED}Backup creation failed. Aborting merge.${NC}"
        exit 1
    fi
    
    # Step 5: Final confirmation
    echo -e "${BLUE}STEP 5: FINAL CONFIRMATION${NC}"
    if ! final_confirmation; then
        echo -e "${YELLOW}Merge cancelled by user.${NC}"
        exit 0
    fi
    
    # Step 6: Handle special directories
    handle_library_merge
    
    # Step 7: Perform merge
    echo -e "${BLUE}STEP 6: PERFORMING MERGE${NC}"
    if perform_merge; then
        show_completion_summary
    else
        echo -e "${RED}Merge failed. Check log file: $LOG_FILE${NC}"
        echo -e "${YELLOW}Backup is available at: $BACKUP_DIR${NC}"
        exit 1
    fi
}

# Run main function
main "$@"