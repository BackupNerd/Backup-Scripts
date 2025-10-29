#!/bin/bash

#=============================================================================
# FILE OWNERSHIP MANAGER SCRIPT
#=============================================================================
# 
# Script Name: changefileowner.v03.sh
# Author: System Administrator
# Version: 3.0
# Date: September 30, 2025
#
# DESCRIPTION:
# Interactive script for safely changing file and directory ownership in macOS.
# Provides a user-friendly interface to select paths, users, and operation modes
# with comprehensive error handling and preview capabilities.
#
# FEATURES:
# - Interactive path selection from /Users directory with ownership details
# - User selection with full profile information display
# - Preview mode to see what would be changed before making actual changes
# - Safe ownership changes with detailed progress reporting
# - Multiple exit points (X option) throughout the interface
# - Proper handling of system-protected files
# - Root user support with correct home directory detection
#
# USAGE:
# ./changefileowner.v03.sh                    # Run as current user
# sudo ./changefileowner.v03.sh               # Run with elevated privileges
#
# REQUIREMENTS:
# - macOS system with standard utilities (stat, dscl, finger, chown)
# - Bash shell
# - Appropriate permissions for the target directories
#
# SAFETY FEATURES:
# - Preview mode shows changes before execution
# - Confirmation prompts for destructive operations
# - Graceful handling of protected system files
# - Detailed error reporting and progress tracking
#
# EXIT CODES:
# 0 - Success or user-initiated exit
# 1 - Error (invalid path, user, or selection)
#
# Note: Removed 'set -e' to allow better error handling
# Individual functions handle their own errors gracefully

#=============================================================================
# FUNCTIONS
#=============================================================================
#
# FUNCTION SUMMARY:
#
# UI Functions:
#   print_header()           - Display script title banner
#   print_separator()        - Print visual separator line
#
# Validation Functions:
#   validate_path($path)     - Verify that specified path exists on filesystem
#   validate_user($user)     - Check if user exists in system user database
#   get_user_full_name($user)- Retrieve full profile name for given username
#
# Interactive Menu Functions:
#   list_user_paths()        - Display /Users paths with owner info, get selection
#   list_users()             - Show available users with UIDs, get user choice  
#   get_user_choice()        - Present ownership options (current/different user)
#   get_operation_mode()     - Choose between preview and actual change modes
#
# Operation Functions:
#   preview_changes($path,$user) - Show what files would change (dry run)
#   confirm_action($path,$user)  - Get user confirmation for ownership changes
#   change_ownership_safe($path,$user) - Safely execute ownership changes
#
# Main Function:
#   main()                   - Primary script logic and user interaction flow
#
# All functions include comprehensive error handling and user-friendly output.
# Interactive functions support 'X' option to exit at any point.
#=============================================================================

print_header() {
    # Print script header
    echo "=========================================="
    echo "         File Ownership Manager"
    echo "=========================================="
}

print_separator() {
    # Print a separator line
    echo "------------------------------------------"
}

validate_path() {
    # Check if the given path exists
    local path="$1"
    if [[ ! -e "$path" ]]; then
        echo "[ERROR] Path '$path' does not exist"
        exit 1
    fi
}

validate_user() {
    # Check if the given user exists on the system
    local user="$1"
    if ! id "$user" &>/dev/null; then
        echo "[ERROR] User '$user' does not exist"
        exit 1
    fi
}

get_user_full_name() {
    # Get the full profile name for a given user
    local username="$1"
    local full_name=""
    
    # Try to get full name from system records
    if [[ "$username" == "root" ]]; then
        full_name="System Administrator (root)"
    else
        # Try dscl to get RealName for local users
        full_name=$(dscl . -read /Users/"$username" RealName 2>/dev/null | sed 's/RealName: //' | head -n1 | sed 's/^ *//')
        
        # If dscl didn't work or returned empty, try finger command
        if [[ -z "$full_name" || "$full_name" == "RealName:" ]]; then
            full_name=$(finger "$username" 2>/dev/null | head -n1 | sed 's/.*Name: //' | sed 's/Directory:.*//' | sed 's/^ *//' | sed 's/ *$//')
        fi
        
        # If still empty, try getent passwd
        if [[ -z "$full_name" ]]; then
            full_name=$(getent passwd "$username" 2>/dev/null | cut -d: -f5 | cut -d, -f1)
        fi
        
        # If still empty, use username as fallback
        if [[ -z "$full_name" ]]; then
            full_name="$username"
        fi
    fi
    
    echo "$full_name"
}

list_user_paths() {
    # List available paths under /Users and prompt for selection
    echo "Available paths under /Users:"
    print_separator
    
    local path_count=0
    declare -a paths
    declare -a owners
    declare -a uids
    
    # Add /Users directory itself as first option
    local users_stat_info=$(stat -f "%Su %u" "/Users" 2>/dev/null || echo "unknown 0")
    local users_owner_name=$(echo "$users_stat_info" | cut -d' ' -f1)
    local users_owner_uid=$(echo "$users_stat_info" | cut -d' ' -f2)
    local users_full_profile=$(get_user_full_name "$users_owner_name")
    
    ((path_count++))
    paths[$path_count]="/Users"
    owners[$path_count]="$users_owner_name"
    uids[$path_count]="$users_owner_uid"
    printf "  [%-2d]\t%-30s\tOwner: %-15s\tUID: %-5s\tProfile: %s\n" "$path_count" "/Users" "$users_owner_name" "$users_owner_uid" "$users_full_profile"

    # Get subdirectories from /Users
    for user_dir in /Users/*; do
        if [[ -d "$user_dir" ]]; then
            # Get ownership information for each subdirectory
            local stat_info=$(stat -f "%Su %u" "$user_dir" 2>/dev/null || echo "unknown 0")
            local owner_name=$(echo "$stat_info" | cut -d' ' -f1)
            local owner_uid=$(echo "$stat_info" | cut -d' ' -f2)
            local owner_full_profile=$(get_user_full_name "$owner_name")
            
            ((path_count++))
            paths[$path_count]="$user_dir"
            owners[$path_count]="$owner_name"
            uids[$path_count]="$owner_uid"
            
            printf "  [%-2d]\t%-30s\tOwner: %-15s\tUID: %-5s\tProfile: %s\n" "$path_count" "$user_dir" "$owner_name" "$owner_uid" "$owner_full_profile"
        fi
    done
    
    if [[ $path_count -eq 0 ]]; then
        # No directories found
        echo "  [ERROR] No directories found in /Users"
        return 1
    fi
    
    # Option to enter a custom path
    echo "  [0 ]  Enter custom path"
    echo "  [X ]  Exit script"
    print_separator
    
    read -p "Select path number (0-$path_count, X to exit): " path_choice
    echo
    
    if [[ $path_choice == "X" || $path_choice == "x" ]]; then
        # User chooses to exit
        echo "[INFO] Exiting script..."
        exit 0
    elif [[ $path_choice -eq 0 ]]; then
        # User chooses to enter a custom path
        read -p "Enter custom path: " target_path
        validate_path "$target_path"
    elif [[ $path_choice -ge 1 && $path_choice -le $path_count ]]; then
        # User selects from the listed paths
        target_path="${paths[$path_choice]}"
    else
        # Invalid selection
        echo "[ERROR] Invalid selection"
        exit 1
    fi
}

list_users() {
    # List available users based on /Users directory and prompt for selection
    echo "Available users in /Users:"
    print_separator
    
    local user_count=0
    declare -a users
    declare -a uids
    
    # Get users from /Users directory
    for user_dir in /Users/*; do
        if [[ -d "$user_dir" ]]; then
            username=$(basename "$user_dir")
            if id "$username" &>/dev/null; then
                # Only include valid system users
                user_uid=$(id -u "$username")
                ((user_count++))
                users[$user_count]="$username"
                uids[$user_count]="$user_uid"
                echo "  [$user_count] $username (UID: $user_uid)"
            fi
        fi
    done
    
    if [[ $user_count -eq 0 ]]; then
        # No valid users found
        echo "  [ERROR] No valid users found in /Users"
        return 1
    fi
    
    # Option to enter a custom username
    echo "  [0] Enter custom username"
    echo "  [X] Exit script"
    print_separator
    
    read -p "Select user number (0-$user_count, X to exit): " user_choice
    echo
    
    if [[ $user_choice == "X" || $user_choice == "x" ]]; then
        # User chooses to exit
        echo "[INFO] Exiting script..."
        exit 0
    elif [[ $user_choice -eq 0 ]]; then
        # User chooses to enter a custom username
        read -p "Enter username: " target_user
        validate_user "$target_user"
    elif [[ $user_choice -ge 1 && $user_choice -le $user_count ]]; then
        # User selects from the listed users
        target_user="${users[$user_choice]}"
    else
        # Invalid selection
        echo "[ERROR] Invalid selection"
        exit 1
    fi
}

get_user_choice() {
    # Prompt user to choose ownership option
    echo "Choose ownership option:"
    echo "  [1] Take ownership as current user ($current_user)"
    echo "  [2] Specify a different user"
    echo "  [X] Exit script"
    echo
    read -p "Enter choice (1, 2, or X to exit): " choice
    echo
}

get_operation_mode() {
    # Prompt user to choose operation mode
    echo "Choose operation mode:"
    echo "  [1] Preview only (show what would be changed)"
    echo "  [2] Make actual changes"
    echo "  [X] Exit script"
    echo
    read -p "Enter choice (1, 2, or X to exit): " mode_choice
    echo
}

preview_changes() {
    # Show what would be changed without making actual changes
    local target_path="$1"
    local target_user="$2"
    
    echo "[PREVIEW] Files that would be changed:"
    print_separator
    
    local count=0
    while IFS= read -r -d '' file; do
        local current_owner=$(stat -f "%Su" "$file" 2>/dev/null || echo "unknown")
        if [[ "$current_owner" != "$target_user" ]]; then
            ((count++))
            if [[ $count -le 10 ]]; then
                echo "  $file (currently: $current_owner → would become: $target_user)"
            fi
        fi
    done < <(find "$target_path" -print0 2>/dev/null)
    
    if [[ $count -gt 10 ]]; then
        echo "  ... and $((count - 10)) more files"
    fi
    
    print_separator
    echo "[PREVIEW] Total files that would be changed: $count"
    echo "[PREVIEW] No actual changes were made"
}

confirm_action() {
    # Confirm ownership change action with user
    local path="$1"
    local user="$2"
    
    print_separator
    echo "[WARNING] This will change ownership of ALL files in:"
    echo "   Path: $path"
    echo "   New owner: $user"
    echo ""
    echo "[NOTE] Some system-protected files may be skipped (this is normal)"
    print_separator
    
    read -p "Continue? (y/N): " confirm
    echo
}

change_ownership_safe() {
    # Safely change ownership with better error handling
    local target_path="$1"
    local target_user="$2"
    local success_count=0
    local error_count=0
    local total_files=0
    
    echo "[INFO] Analyzing files to change..."
    
    # Count total files (for progress indication)
    if [[ -d "$target_path" ]]; then
        total_files=$(find "$target_path" -type f 2>/dev/null | wc -l | tr -d ' ')
        echo "[INFO] Found $total_files files to process"
    fi
    
    echo "[INFO] Starting ownership change..."
    echo "[INFO] Note: System-protected files will be skipped automatically"
    
    # Use chown with error handling but continue on errors
    if sudo chown -R "$target_user" "$target_path" 2>/dev/null; then
        echo "[SUCCESS] All files processed successfully!"
    else
        # Run with verbose output to show which files succeeded/failed
        echo "[INFO] Some files were protected - running detailed analysis..."
        
        # Create a more detailed report
        while IFS= read -r -d '' file; do
            if sudo chown "$target_user" "$file" 2>/dev/null; then
                ((success_count++))
            else
                ((error_count++))
                if [[ $error_count -le 5 ]]; then
                    echo "[SKIP] Protected: $(basename "$file")"
                fi
            fi
            
            # Show progress every 100 files
            if (( (success_count + error_count) % 100 == 0 )); then
                echo "[PROGRESS] Processed $((success_count + error_count)) files..."
            fi
        done < <(find "$target_path" -print0 2>/dev/null)
        
        print_separator
        echo "[SUMMARY] Ownership change completed:"
        echo "  ✓ Successfully changed: $success_count files"
        echo "  ✗ Protected/skipped: $error_count files"
        echo "  Total processed: $((success_count + error_count)) files"
        
        if [[ $error_count -gt 5 ]]; then
            echo "[INFO] Only first 5 protected files were shown above"
        fi
        
        if [[ $success_count -gt 0 ]]; then
            echo "[SUCCESS] Operation completed with some files changed!"
        else
            echo "[WARNING] No files were changed - all may be system-protected"
        fi
    fi
}

#=============================================================================
# MAIN SCRIPT
#=============================================================================

main() {
    # Clear the screen for better readability
    clear

    # Display header
    print_header

    # Get current user information (only once)
    current_uid=$(id -u)
    current_user=$(id -un)

    # Get the actual home directory for the effective user
    if [[ $current_uid -eq 0 ]]; then
        # Root user - use system root home directory
        actual_home="/var/root"
    else
        # Regular user - try to get actual home directory, fall back to $HOME
        actual_home=$(dscl . -read /Users/$current_user NFSHomeDirectory 2>/dev/null | cut -d: -f2 | tr -d ' ' || echo "$HOME")
    fi

    # Get current user's full profile
    userprofile=$(system_profiler SPSoftwareDataType | grep "User Name" | sed 's/.*User Name: //' 2>/dev/null || echo "Unknown")
    echo "Current user's full profile: $userprofile"
    echo "[INFO] Current user: $current_user (UID: $current_uid)"
    echo "Your home directory is: $actual_home"    
    echo
 
    # List available paths and get selection
    list_user_paths
    
    # Get user choice for ownership
    get_user_choice
    
    case $choice in
        1)
            # Take ownership as current user
            target_user="$current_user"
            ;;
        2)
            # Specify a different user
            list_users
            ;;
        X|x)
            # User chooses to exit
            echo "[INFO] Exiting script..."
            exit 0
            ;;
        *)
            # Invalid choice
            echo "[ERROR] Invalid choice. Please select 1, 2, or X."
            return 1
            ;;
    esac
    
    # Get operation mode (preview or actual change)
    get_operation_mode
    
    case $mode_choice in
        1)
            # Preview mode
            echo "[INFO] Running in PREVIEW mode - no changes will be made"
            preview_changes "$target_path" "$target_user"
            ;;
        2)
            # Actual change mode
            confirm_action "$target_path" "$target_user"
            
            if [[ $confirm =~ ^[Yy]$ ]]; then
                change_ownership_safe "$target_path" "$target_user"
            else
                echo "[CANCELLED] Operation cancelled"
            fi
            ;;
        X|x)
            # User chooses to exit
            echo "[INFO] Exiting script..."
            exit 0
            ;;
        *)
            # Invalid choice
            echo "[ERROR] Invalid choice. Please select 1, 2, or X."
            return 1
            ;;
    esac
}

# Run main function
main "$@"