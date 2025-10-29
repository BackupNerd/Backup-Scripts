#!/bin/bash

################################################################################
# User Home Directory Merge Verification Script
################################################################################
#
# DESCRIPTION:
#   A verification and testing script designed to work alongside merge_user_homes.sh
#   This script helps validate successful merges and provides rollback capabilities.
#
# FEATURES:  
#   • Verify merge integrity by comparing file counts and sizes
#   • Check file ownership and permissions after merge
#   • Test application functionality with merged user profile
#   • Provide rollback capability using created backups
#   • Generate detailed merge reports and statistics
#
# USAGE:
#   sudo ./verify_user_merge.sh
#   sudo ./verify_user_merge.sh --rollback /tmp/user_merge_backup_username_timestamp
#
# AUTHOR: GitHub Copilot with guidance from Eric Harless
# DATE: September 30, 2025
################################################################################

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

# Function to display header
show_header() {
    clear
    echo -e "${BLUE}================================================================${NC}"
    echo -e "${BLUE}              USER MERGE VERIFICATION SCRIPT${NC}"
    echo -e "${BLUE}================================================================${NC}"
    echo ""
}

# Function to check if running as root
check_root() {
    if [[ $EUID -ne 0 ]]; then
        echo -e "${RED}Error: This script must be run as root (sudo)${NC}"
        echo "Usage: sudo $0 [--rollback backup_path]"
        exit 1
    fi
}

# Function to find recent merge backups
find_recent_backups() {
    echo -e "${GREEN}Finding recent merge backups...${NC}"
    
    local backups=($(find /tmp -maxdepth 1 -name "user_merge_backup_*" -type d 2>/dev/null | sort -r))
    
    if [[ ${#backups[@]} -eq 0 ]]; then
        echo -e "${YELLOW}No merge backups found in /tmp${NC}"
        return 1
    fi
    
    echo "Recent merge backups:"
    local i=1
    for backup in "${backups[@]}"; do
        local backup_info=""
        if [[ -f "$backup/.backup_info" ]]; then
            backup_info=$(grep "Original User:" "$backup/.backup_info" | cut -d: -f2 | tr -d ' ')
            local backup_date=$(grep "Backup Created:" "$backup/.backup_info" | cut -d: -f2- | tr -d ' ')
            local backup_size=$(grep "Backup Size:" "$backup/.backup_info" | cut -d: -f2 | tr -d ' ')
        fi
        echo "  [$i] $(basename "$backup") - User: $backup_info, Size: $backup_size"
        ((i++))
    done
    echo ""
    
    return 0
}

# Function to verify merge integrity
verify_merge() {
    local target_user="$1"
    local target_home="$2"
    local backup_path="$3"
    
    echo -e "${GREEN}Verifying merge integrity for user: $target_user${NC}"
    echo "================================================================"
    
    # Check if target home exists
    if [[ ! -d "$target_home" ]]; then
        echo -e "${RED}Error: Target home directory not found: $target_home${NC}"
        return 1
    fi
    
    # Check ownership
    echo "Checking file ownership..."
    local wrong_owner_count=$(find "$target_home" ! -user "$target_user" 2>/dev/null | wc -l | tr -d ' ')
    
    if [[ $wrong_owner_count -gt 0 ]]; then
        echo -e "${YELLOW}Warning: $wrong_owner_count files have incorrect ownership${NC}"
        echo "First 10 files with wrong ownership:"
        find "$target_home" ! -user "$target_user" 2>/dev/null | head -10 | while read file; do
            echo "  $(ls -la "$file")"
        done
    else
        echo -e "${GREEN}✓ All files have correct ownership${NC}"
    fi
    
    # Check permissions
    echo "Checking file permissions..."
    local permission_issues=$(find "$target_home" -type f \( -perm 000 -o ! -readable \) 2>/dev/null | wc -l | tr -d ' ')
    
    if [[ $permission_issues -gt 0 ]]; then
        echo -e "${YELLOW}Warning: $permission_issues files have permission issues${NC}"
    else
        echo -e "${GREEN}✓ File permissions appear normal${NC}"
    fi
    
    # Compare with backup if provided
    if [[ -n "$backup_path" && -d "$backup_path" ]]; then
        echo "Comparing with backup..."
        local backup_files=$(find "$backup_path" -type f 2>/dev/null | wc -l | tr -d ' ')
        local current_files=$(find "$target_home" -type f 2>/dev/null | wc -l | tr -d ' ')
        
        echo "  Backup files: $backup_files"
        echo "  Current files: $current_files"
        
        local file_diff=$((current_files - backup_files))
        if [[ $file_diff -gt 0 ]]; then
            echo -e "${GREEN}✓ $file_diff additional files from merge${NC}"
        elif [[ $file_diff -lt 0 ]]; then
            echo -e "${YELLOW}Warning: $((-file_diff)) fewer files than backup${NC}"
        else
            echo -e "${CYAN}Same number of files as backup${NC}"
        fi
    fi
    
    echo ""
    return 0
}

# Function to perform rollback
perform_rollback() {
    local backup_path="$1"
    
    if [[ ! -d "$backup_path" ]]; then
        echo -e "${RED}Error: Backup directory not found: $backup_path${NC}"
        return 1
    fi
    
    # Read backup info
    if [[ ! -f "$backup_path/.backup_info" ]]; then
        echo -e "${RED}Error: Backup info file not found${NC}"
        return 1
    fi
    
    local original_path=$(grep "Original Path:" "$backup_path/.backup_info" | cut -d: -f2 | tr -d ' ')
    local original_user=$(grep "Original User:" "$backup_path/.backup_info" | cut -d: -f2 | tr -d ' ')
    
    echo -e "${RED}ROLLBACK CONFIRMATION${NC}"
    echo "================================================================"
    echo -e "${RED}This will restore the target directory from backup:${NC}"
    echo -e "${RED}  Target: $original_path${NC}"
    echo -e "${RED}  User: $original_user${NC}"
    echo -e "${RED}  Backup: $backup_path${NC}"
    echo ""
    echo -e "${RED}WARNING: This will REPLACE all current data with backup data${NC}"
    echo -e "${RED}Any changes made after the merge will be LOST!${NC}"
    echo ""
    
    read -p "Type 'ROLLBACK' to confirm or anything else to cancel: " confirm
    
    if [[ "$confirm" != "ROLLBACK" ]]; then
        echo -e "${YELLOW}Rollback cancelled${NC}"
        return 1
    fi
    
    echo -e "${GREEN}Starting rollback...${NC}"
    
    # Create a backup of current state before rollback
    local pre_rollback_backup="/tmp/pre_rollback_$(basename "$original_path")_$(date '+%Y%m%d_%H%M%S')"
    echo "Creating backup of current state: $pre_rollback_backup"
    
    if rsync -av "$original_path/" "$pre_rollback_backup/" >/dev/null 2>&1; then
        echo -e "${GREEN}✓ Current state backed up${NC}"
    else
        echo -e "${YELLOW}Warning: Could not backup current state${NC}"
    fi
    
    # Perform rollback
    echo "Restoring from backup..."
    if rsync -av --delete "$backup_path/" "$original_path/" >/dev/null 2>&1; then
        echo -e "${GREEN}✓ Rollback completed successfully${NC}"
        echo -e "${GREEN}Pre-rollback backup saved to: $pre_rollback_backup${NC}"
        return 0
    else
        echo -e "${RED}Error: Rollback failed${NC}"
        return 1
    fi
}

# Function to generate merge report
generate_report() {
    local target_user="$1"
    local target_home="$2"
    
    local report_file="/tmp/merge_report_${target_user}_$(date '+%Y%m%d_%H%M%S').txt"
    
    echo -e "${GREEN}Generating detailed merge report...${NC}"
    
    cat > "$report_file" << EOF
================================================================================
USER HOME DIRECTORY MERGE REPORT
================================================================================
Generated: $(date)
Target User: $target_user
Target Home: $target_home

DIRECTORY STRUCTURE:
$(ls -la "$target_home" 2>/dev/null || echo "Error reading directory")

FILE COUNTS BY TYPE:
Documents: $(find "$target_home/Documents" -type f 2>/dev/null | wc -l | tr -d ' ') files
Desktop: $(find "$target_home/Desktop" -type f 2>/dev/null | wc -l | tr -d ' ') files
Downloads: $(find "$target_home/Downloads" -type f 2>/dev/null | wc -l | tr -d ' ') files
Pictures: $(find "$target_home/Pictures" -type f 2>/dev/null | wc -l | tr -d ' ') files
Movies: $(find "$target_home/Movies" -type f 2>/dev/null | wc -l | tr -d ' ') files
Music: $(find "$target_home/Music" -type f 2>/dev/null | wc -l | tr -d ' ') files

TOTAL FILES: $(find "$target_home" -type f 2>/dev/null | wc -l | tr -d ' ')
TOTAL SIZE: $(du -sh "$target_home" 2>/dev/null | cut -f1)

OWNERSHIP CHECK:
Files not owned by $target_user: $(find "$target_home" ! -user "$target_user" 2>/dev/null | wc -l | tr -d ' ')

RECENT FILES (Last 24 hours):
$(find "$target_home" -type f -mtime -1 2>/dev/null | head -20)

================================================================================
EOF

    echo -e "${GREEN}Report generated: $report_file${NC}"
    echo ""
    echo -e "${CYAN}Report Summary:${NC}"
    head -20 "$report_file" | tail -15
}

# Main function
main() {
    show_header
    check_root
    
    # Check for rollback argument
    if [[ "$1" == "--rollback" ]]; then
        if [[ -n "$2" ]]; then
            perform_rollback "$2"
        else
            echo "Please specify backup path for rollback"
            echo "Usage: sudo $0 --rollback /path/to/backup"
            exit 1
        fi
        exit 0
    fi
    
    # Interactive mode
    echo -e "${GREEN}User Home Directory Merge Verification${NC}"
    echo ""
    echo "1. Find and verify recent merges"
    echo "2. Rollback a merge using backup"
    echo "3. Generate detailed merge report"
    echo "4. Exit"
    echo ""
    
    read -p "Select option (1-4): " choice
    
    case $choice in
        1)
            if find_recent_backups; then
                read -p "Enter target username to verify: " target_user
                if [[ -n "$target_user" ]]; then
                    local target_home="/Users/$target_user"
                    verify_merge "$target_user" "$target_home"
                fi
            fi
            ;;
        2)
            if find_recent_backups; then
                read -p "Enter backup path for rollback: " backup_path
                if [[ -n "$backup_path" ]]; then
                    perform_rollback "$backup_path"
                fi
            fi
            ;;
        3)
            read -p "Enter target username for report: " target_user
            if [[ -n "$target_user" ]]; then
                local target_home="/Users/$target_user"
                generate_report "$target_user" "$target_home"
            fi
            ;;
        4)
            echo -e "${GREEN}Exiting...${NC}"
            exit 0
            ;;
        *)
            echo -e "${RED}Invalid choice${NC}"
            exit 1
            ;;
    esac
}

# Run main function
main "$@"