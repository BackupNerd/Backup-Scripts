#!/bin/bash

################################################################################
# Enhanced macOS User Management Script
################################################################################
#
# DESCRIPTION:
#   A comprehensive shell script for macOS that provides complete user account 
#   management including creation, enumeration, detailed viewing, and safe 
#   deletion with proper security configuration and home directory handling.
#
# FEATURES:
#   USER CREATION:
#   • Interactive user account creation with validation
#   • Automatic UID assignment and conflict detection
#   • Home directory creation with proper permissions
#   • Administrative privilege configuration (standard, admin, admin+sudo)
#   • Password strength validation and security checks
#   • Account verification and rollback capabilities
#
#   USER MANAGEMENT:
#   • List all regular users (UID ≥ 501) with formatted display
#   • View detailed user information including groups and processes
#   • Safe user deletion with multi-step confirmation process
#   • Automatic process termination before user removal
#   • Orphaned home directory detection and cleanup
#   • Current user protection (cannot delete logged-in user)
#
#   SECURITY FEATURES:
#   • Root privilege enforcement
#   • Multiple confirmation prompts for deletions
#   • Protection against deleting current user
#   • System user exclusion (UID < 501)
#   • Strong password requirements
#   • Comprehensive logging for audit purposes
#
# REQUIREMENTS:
#   • macOS 10.12 Sierra or later
#   • Administrator privileges (run with sudo)
#   • Standard macOS utilities (dscl, id, mkdir, chown, chmod, ps, who)
#
# USAGE:
#   sudo ./enhanced_user_manager.sh
#
# AUTHOR: GitHub Copilot with guidance from Eric Harless
# DATE: October 29, 2025
# VERSION: 2.0 - Enhanced User Creation and Management
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
NEW_USERNAME=""
NEW_FULLNAME=""
NEW_PASSWORD=""
NEW_UID=""
NEW_GID=""
NEW_HOME=""
NEW_SHELL=""
ADMIN_LEVEL=""
LOG_FILE=""

################################################################################
# UTILITY FUNCTIONS
################################################################################

# Function to display main header
show_header() {
    echo "================================================================"
    echo -e "${BLUE}           ENHANCED MACOS USER MANAGEMENT SCRIPT${NC}"
    echo "================================================================"
    echo "This script provides comprehensive macOS user account management"
    echo "including creation, enumeration, viewing, and safe deletion."
    echo ""
    echo -e "${RED}WARNING: This script can create and delete user accounts!${NC}"
    echo -e "${RED}Use with extreme caution and verify all actions!${NC}"
    echo ""
}

# Function to check if running as root
check_root() {
    if [[ $EUID -ne 0 ]]; then
        echo -e "${RED}This script must be run as root (use sudo)${NC}"
        echo "Usage: sudo $0"
        exit 1
    fi
}

# Function to get current user (the real user behind sudo)
get_current_user() {
    if [[ -n "$SUDO_USER" ]]; then
        CURRENT_USER="$SUDO_USER"
    elif [[ -n "$USER" ]]; then
        CURRENT_USER="$USER"
    else
        CURRENT_USER="unknown"
    fi
}

# Function to set up logging
setup_logging() {
    local timestamp=$(date +"%Y%m%d_%H%M%S")
    LOG_FILE="/var/log/user_management_${timestamp}.log"
    
    # Create log file and set permissions
    touch "$LOG_FILE" 2>/dev/null || {
        echo -e "${YELLOW}Warning: Could not create log file at $LOG_FILE${NC}"
        LOG_FILE=""
        return 1
    }
    
    chmod 640 "$LOG_FILE" 2>/dev/null
    log_message "INFO" "Enhanced user management script started by: $CURRENT_USER"
}

# Function to log messages
log_message() {
    local level="$1"
    local message="$2"
    local timestamp=$(date '+%Y-%m-%d %H:%M:%S')
    
    # Only log if LOG_FILE is set and writable
    if [[ -n "$LOG_FILE" && -w "$LOG_FILE" ]]; then
        echo "[$timestamp] [$level] $message" >> "$LOG_FILE"
    fi
}

################################################################################
# MENU FUNCTIONS
################################################################################

# Function to show main menu
show_main_menu() {
    echo ""
    echo -e "${GREEN}=== MAIN MENU ===${NC}"
    echo "1. Create New User Account"
    echo "2. List All Users"
    echo "3. View User Details"
    echo "4. Delete User Account"
    echo "5. Clean Up Orphaned Home Directories"
    echo "6. Exit"
    echo ""
    
    if [[ -n "$CURRENT_USER" && "$CURRENT_USER" != "unknown" ]]; then
        echo -e "${CYAN}Protected User: $CURRENT_USER (cannot be deleted)${NC}"
    fi
    
    if [[ -n "$LOG_FILE" ]]; then
        echo -e "${CYAN}Logging to: $LOG_FILE${NC}"
    fi
    echo ""
}

# Function to get menu choice
get_menu_choice() {
    read -r choice
    echo "$choice"
}

################################################################################
# USER ENUMERATION AND MANAGEMENT FUNCTIONS
################################################################################

# Function to check if user has admin privileges
check_admin_privileges() {
    local username=$1
    local admin_status=""
    
    # Check if user is in admin group
    if dscl . -read /Groups/admin GroupMembership 2>/dev/null | grep -q "\b$username\b"; then
        admin_status="ADMIN"
    fi
    
    # Check if user is in wheel group (sudo privileges)
    if dscl . -read /Groups/wheel GroupMembership 2>/dev/null | grep -q "\b$username\b"; then
        if [[ -n "$admin_status" ]]; then
            admin_status="ADMIN+SUDO"
        else
            admin_status="SUDO"
        fi
    fi
    
    # Check for other administrative groups
    local other_admin_groups=""
    
    # Check _appserveradm (App Server Admin)
    if dscl . -read /Groups/_appserveradm GroupMembership 2>/dev/null | grep -q "\b$username\b"; then
        other_admin_groups="${other_admin_groups}APPSERVER "
    fi
    
    # Check _lpadmin (Print Admin)
    if dscl . -read /Groups/_lpadmin GroupMembership 2>/dev/null | grep -q "\b$username\b"; then
        other_admin_groups="${other_admin_groups}PRINT "
    fi
    
    # Combine admin status with other privileges
    if [[ -n "$other_admin_groups" ]]; then
        if [[ -n "$admin_status" ]]; then
            admin_status="$admin_status+${other_admin_groups% }"
        else
            admin_status="${other_admin_groups% }"
        fi
    fi
    
    echo "$admin_status"
}

# Function to get all users (excluding system users)
list_all_users() {
    echo -e "${GREEN}Enumerating users...${NC}"
    echo ""
    
    # Get users with UID >= 501 (macOS standard for regular users)
    local users=()
    
    for username in $(dscl . -list /Users | grep -v '^_' | grep -v '^Guest' | grep -v '^nobody'); do
        local uid=$(dscl . -read /Users/"$username" UniqueID 2>/dev/null | awk '{print $2}')
        local gid=$(dscl . -read /Users/"$username" PrimaryGroupID 2>/dev/null | awk '{print $2}')
        local home=$(dscl . -read /Users/"$username" NFSHomeDirectory 2>/dev/null | awk '{print $2}')
        local shell=$(dscl . -read /Users/"$username" UserShell 2>/dev/null | awk '{print $2}')
        local gecos=$(dscl . -read /Users/"$username" RealName 2>/dev/null | cut -d: -f2- | xargs)
        
        # Skip system users (UID < 501 on macOS)
        if [[ -n "$uid" && $uid -ge 501 ]] && [[ $username != "nobody" ]]; then
            local admin_status=$(check_admin_privileges "$username")
            users+=("$username:$uid:$gid:$home:$shell:$gecos:$admin_status")
        fi
    done
    
    if [[ ${#users[@]} -eq 0 ]]; then
        echo -e "${YELLOW}No regular users found.${NC}"
        return 1
    fi
    
    # Display users in a formatted table
    printf "%-4s %-6s %-15s %-20s %-6s %-6s %-25s %-12s %s\n" "No." "Current" "Username" "Full Name" "UID" "GID" "Home Directory" "Shell" "Admin Privileges"
    echo "-------------------------------------------------------------------------------------------------------"
    
    local i=1
    for user_info in "${users[@]}"; do
        IFS=':' read -r username uid gid home shell gecos admin_status <<< "$user_info"
        
        local status=""
        local color=""
        local admin_color=""
        
        # Set colors based on admin privileges
        if [[ -n "$admin_status" ]]; then
            if [[ "$admin_status" == *"ADMIN"* ]]; then
                admin_color="${RED}"
            elif [[ "$admin_status" == *"SUDO"* ]]; then
                admin_color="${YELLOW}"
            else
                admin_color="${BLUE}"
            fi
        fi
        
        if [[ "$username" == "$CURRENT_USER" ]]; then
            status="*"
            color="${CYAN}"
        fi
        
        local admin_display="${admin_status:-"none"}"
        
        # Print the row with appropriate colors
        if [[ -n "$color" ]]; then
            printf "${color}%-4d %-6s %-15s %-20s %-6s %-6s %-25s %-12s ${admin_color}%s${NC}\n" "$i" "$status" "$username" "$gecos" "$uid" "$gid" "$home" "$shell" "$admin_display"
        elif [[ -n "$admin_color" ]]; then
            printf "%-4d %-6s %-15s %-20s %-6s %-6s %-25s %-12s ${admin_color}%s${NC}\n" "$i" "$status" "$username" "$gecos" "$uid" "$gid" "$home" "$shell" "$admin_display"
        else
            printf "%-4d %-6s %-15s %-20s %-6s %-6s %-25s %-12s %s\n" "$i" "$status" "$username" "$gecos" "$uid" "$gid" "$home" "$shell" "$admin_display"
        fi
        ((i++))
    done
    
    echo ""
    echo -e "${BLUE}Admin Privilege Legend:${NC}"
    echo -e "${RED}ADMIN${NC}       - Member of admin group (full administrative access)"
    echo -e "${YELLOW}SUDO${NC}        - Member of wheel group (sudo privileges)"
    echo -e "${RED}ADMIN+SUDO${NC}  - Both admin and sudo privileges"
    echo -e "${BLUE}APPSERVER${NC}   - App Server administrative privileges"
    echo -e "${BLUE}PRINT${NC}       - Print system administrative privileges"
    echo -e "none        - No administrative privileges"
    
    if [[ -n "$CURRENT_USER" && "$CURRENT_USER" != "unknown" ]]; then
        echo ""
        echo -e "${CYAN}Note: The current user ($CURRENT_USER) is protected and cannot be deleted.${NC}"
    fi
    
    echo ""
    log_message "INFO" "Listed all users"
}

# Function to get user details
get_user_details() {
    local username=$1
    
    echo -e "${BLUE}User Details for: $username${NC}"
    echo "================================================================"
    
    if ! dscl . -read /Users/"$username" &>/dev/null; then
        echo -e "${RED}Error: User '$username' not found${NC}"
        return 1
    fi
    
    local uid=$(dscl . -read /Users/"$username" UniqueID 2>/dev/null | awk '{print $2}')
    local gid=$(dscl . -read /Users/"$username" PrimaryGroupID 2>/dev/null | awk '{print $2}')
    local home_dir=$(dscl . -read /Users/"$username" NFSHomeDirectory 2>/dev/null | awk '{print $2}')
    local shell=$(dscl . -read /Users/"$username" UserShell 2>/dev/null | awk '{print $2}')
    local real_name=$(dscl . -read /Users/"$username" RealName 2>/dev/null | cut -d: -f2- | xargs)
    
    echo "User ID: $uid"
    echo "Group ID: $gid"
    echo "Real Name: $real_name"
    echo "Shell: $shell"
    
    # Groups
    if id "$username" &>/dev/null; then
        echo "Groups: $(id -Gn "$username" | tr ' ' ',')"
    fi
    
    # Home directory info
    if [[ -d "$home_dir" ]]; then
        echo "Home Directory: $home_dir"
        echo "Home Directory Size: $(du -sh "$home_dir" 2>/dev/null | cut -f1)"
    else
        echo "Home Directory: $home_dir (not found or inaccessible)"
    fi
    
    # Last login info
    if command -v last &>/dev/null; then
        local last_login=$(last -1 "$username" 2>/dev/null | head -1)
        if [[ -n "$last_login" && "$last_login" != *"utx.log begins"* && "$last_login" != *"wtmp begins"* && "$last_login" != *"begins"* ]]; then
            echo "Last Login: $last_login"
        else
            echo "Last Login: No login records found"
        fi
    fi
    
    # Running processes
    local process_count=$(ps -U "$username" 2>/dev/null | wc -l)
    process_count=$((process_count - 1))
    echo "Running Processes: $process_count"
    if [[ $process_count -gt 0 ]]; then
        echo -e "${YELLOW}Warning: User has running processes!${NC}"
    fi
    
    # Check if user is currently logged in
    if who | grep -q "^$username "; then
        echo -e "${YELLOW}Warning: User is currently logged in!${NC}"
    fi
    
    echo ""
    log_message "INFO" "Displayed user details for: $username"
}

################################################################################
# USER CREATION FUNCTIONS
################################################################################

# Function to validate username format
validate_username() {
    local username="$1"
    
    if [[ ${#username} -lt 3 || ${#username} -gt 20 ]]; then
        echo -e "${RED}Error: Username must be 3-20 characters long${NC}"
        return 1
    fi
    
    if [[ ! "$username" =~ ^[a-z][a-z0-9_-]*$ ]]; then
        echo -e "${RED}Error: Username must start with lowercase letter and contain only lowercase letters, numbers, underscore, and hyphen${NC}"
        return 1
    fi
    
    local reserved_names=("root" "admin" "daemon" "nobody" "www" "mysql" "postgres" "mail" "ftp" "guest")
    for reserved in "${reserved_names[@]}"; do
        if [[ "$username" == "$reserved" ]]; then
            echo -e "${RED}Error: '$username' is a reserved username${NC}"
            return 1
        fi
    done
    
    if dscl . -read /Users/"$username" >/dev/null 2>&1; then
        echo -e "${RED}Error: Username '$username' already exists${NC}"
        return 1
    fi
    
    return 0
}

# Function to validate password strength
validate_password() {
    local password="$1"
    local username="$2"
    
    if [[ ${#password} -lt 8 ]]; then
        echo -e "${RED}Error: Password must be at least 8 characters long${NC}"
        return 1
    fi
    
    if [[ ${#password} -gt 128 ]]; then
        echo -e "${RED}Error: Password must be less than 128 characters${NC}"
        return 1
    fi
    
    if [[ ! "$password" =~ [A-Z] ]]; then
        echo -e "${RED}Error: Password must contain at least one uppercase letter${NC}"
        return 1
    fi
    
    if [[ ! "$password" =~ [a-z] ]]; then
        echo -e "${RED}Error: Password must contain at least one lowercase letter${NC}"
        return 1
    fi
    
    if [[ ! "$password" =~ [0-9] ]]; then
        echo -e "${RED}Error: Password must contain at least one number${NC}"
        return 1
    fi
    
    if [[ "$password" == "$username" ]]; then
        echo -e "${RED}Error: Password cannot be the same as username${NC}"
        return 1
    fi
    
    local weak_passwords=("password" "123456" "admin" "letmein" "welcome" "monkey" "dragon")
    local password_lower=$(echo "$password" | tr '[:upper:]' '[:lower:]')
    for weak in "${weak_passwords[@]}"; do
        if [[ "$password_lower" == "$weak" ]]; then
            echo -e "${RED}Error: '$password' is a commonly used weak password${NC}"
            return 1
        fi
    done
    
    return 0
}

# Function to find next available UID (enhanced to avoid reuse of UIDs with existing file ownership)
find_next_uid() {
    local start_uid=501
    local max_uid=1000
    local check_paths=("/Users" "/var" "/tmp")
    local verbose=${1:-true}  # Allow silent mode for variable assignment
    
    if [[ "$verbose" == "true" ]]; then
        echo -e "${BLUE}Checking for available UID (avoiding file ownership conflicts)...${NC}" >&2
    fi
    
    for ((uid=$start_uid; uid<=max_uid; uid++)); do
        # Check 1: UID not in active Directory Service
        if dscl . -list /Users UniqueID | grep -q " $uid$"; then
            continue  # UID is in use by active user
        fi
        
        # Check 2: No files owned by this UID in critical paths
        local uid_in_use=false
        
        for path in "${check_paths[@]}"; do
            if [[ -d "$path" ]]; then
                # Check for files owned by this UID (limit search depth and time)
                local file_count
                file_count=$(find "$path" -maxdepth 3 -uid "$uid" -print -quit 2>/dev/null | wc -l)
                
                if [[ $file_count -gt 0 ]]; then
                    if [[ "$verbose" == "true" ]]; then
                        echo -e "${YELLOW}  UID $uid: Files found in $path - skipping${NC}" >&2
                    fi
                    uid_in_use=true
                    break
                fi
            fi
        done
        
        # Check 3: No orphaned home directory with this UID
        if [[ "$uid_in_use" == "false" && -d "/Users" ]]; then
            local orphaned_dir
            orphaned_dir=$(find /Users -maxdepth 1 -type d -user "$uid" -print -quit 2>/dev/null)
            
            if [[ -n "$orphaned_dir" ]]; then
                if [[ "$verbose" == "true" ]]; then
                    echo -e "${YELLOW}  UID $uid: Orphaned directory found ($orphaned_dir) - skipping${NC}" >&2
                fi
                uid_in_use=true
            fi
        fi
        
        # Check 4: No running processes owned by this UID
        if [[ "$uid_in_use" == "false" ]] && ps aux | awk '{print $2}' | grep -q "^$uid$" 2>/dev/null; then
            if [[ "$verbose" == "true" ]]; then
                echo -e "${YELLOW}  UID $uid: Active processes found - skipping${NC}" >&2
            fi
            uid_in_use=true
        fi
        
        # If all checks pass, this UID is safe to use
        if [[ "$uid_in_use" == "false" ]]; then
            if [[ "$verbose" == "true" ]]; then
                echo -e "${GREEN}  UID $uid: Safe to use (no conflicts found)${NC}" >&2
            fi
            echo "$uid"
            return 0
        fi
    done
    
    if [[ "$verbose" == "true" ]]; then
        echo -e "${RED}Error: No available UID found in range $start_uid-$max_uid${NC}" >&2
        echo -e "${RED}All UIDs either in use or have file ownership conflicts${NC}" >&2
    fi
    return 1
}

# Function to get username input
get_username_input() {
    while true; do
        echo -e "${CYAN}USERNAME CONFIGURATION${NC}"
        echo "================================================================"
        echo "Username requirements:"
        echo "  • 3-20 characters long"
        echo "  • Start with lowercase letter"
        echo "  • Only lowercase letters, numbers, underscore, hyphen"
        echo "  • Must be unique on this system"
        echo ""
        echo -e -n "Enter new username: "
        read -r NEW_USERNAME
        
        if validate_username "$NEW_USERNAME"; then
            echo -e "${GREEN}✓ Username '$NEW_USERNAME' is valid and available${NC}"
            log_message "INFO" "Username selected: $NEW_USERNAME"
            break
        else
            echo ""
            echo -e -n "${YELLOW}Press Enter to try again...${NC}"
            read -r
            echo ""
        fi
    done
}

# Function to get full name input
get_fullname_input() {
    echo ""
    echo -e "${CYAN}FULL NAME CONFIGURATION${NC}"
    echo "================================================================"
    echo "Enter the user's full name (Real Name field):"
    echo "  • This appears in the login screen and user interface"
    echo "  • Can contain spaces and special characters"
    echo "  • Leave empty to use the username"
    echo ""
    echo -e -n "Enter full name for '$NEW_USERNAME': "
    read -r NEW_FULLNAME
    
    if [[ -z "$NEW_FULLNAME" ]]; then
        NEW_FULLNAME="$NEW_USERNAME"
    fi
    
    echo "Full name set to: $NEW_FULLNAME"
    log_message "INFO" "Full name set: $NEW_FULLNAME"
}

# Function to get password input
get_password_input() {
    while true; do
        echo ""
        echo -e "${CYAN}PASSWORD CONFIGURATION${NC}"
        echo "================================================================"
        echo "Password requirements:"
        echo "  • At least 8 characters long"
        echo "  • At least one uppercase letter"
        echo "  • At least one lowercase letter"
        echo "  • At least one number"
        echo "  • Cannot be same as username"
        echo "  • Cannot be a common weak password"
        echo ""
        echo -e -n "Enter password for '$NEW_USERNAME': "
        read -s NEW_PASSWORD
        echo ""
        echo -e -n "Confirm password: "
        read -s password_confirm
        echo ""
        
        if [[ "$NEW_PASSWORD" != "$password_confirm" ]]; then
            echo -e "${RED}Error: Passwords do not match${NC}"
            echo ""
            continue
        fi
        
        if validate_password "$NEW_PASSWORD" "$NEW_USERNAME"; then
            echo -e "${GREEN}✓ Password meets security requirements${NC}"
            log_message "INFO" "Password validated for user: $NEW_USERNAME"
            break
        else
            echo ""
        fi
    done
}

# Function to get admin level
get_admin_level() {
    while true; do
        echo ""
        echo -e "${CYAN}ADMINISTRATIVE PRIVILEGES${NC}"
        echo "================================================================"
        echo "Select the administrative level for '$NEW_USERNAME':"
        echo ""
        echo "1. Standard User"
        echo "   • No administrative privileges"
        echo "   • Cannot install software or modify system settings"
        echo "   • Safest option for regular users"
        echo ""
        echo "2. Administrator"
        echo "   • Member of admin group"
        echo "   • Can install software and modify system settings"
        echo "   • Can manage other user accounts"
        echo ""
        echo "3. Administrator + Sudo"
        echo "   • Full administrative privileges"
        echo "   • Member of both admin and wheel groups"
        echo "   • Can use sudo command for system administration"
        echo "   • Highest privilege level"
        echo ""
        echo -e -n "Select option (1-3): "
        read -r admin_choice
        
        case "$admin_choice" in
            1)
                ADMIN_LEVEL="standard"
                echo "Selected: Standard User (no admin privileges)"
                log_message "INFO" "Administrative level selected: standard"
                break
                ;;
            2)
                ADMIN_LEVEL="admin"
                echo "Selected: Administrator (admin group)"
                log_message "INFO" "Administrative level selected: admin"
                break
                ;;
            3)
                ADMIN_LEVEL="admin_sudo"
                echo "Selected: Administrator + Sudo (admin + wheel groups)"
                echo -e "${YELLOW}Warning: This grants full administrative privileges${NC}"
                log_message "INFO" "Administrative level selected: admin_sudo"
                break
                ;;
            *)
                echo -e "${RED}Invalid selection. Please choose 1, 2, or 3.${NC}"
                ;;
        esac
    done
}

# Function to get shell input
get_shell_input() {
    while true; do
        echo ""
        echo -e "${CYAN}SHELL CONFIGURATION${NC}"
        echo "================================================================"
        echo "Select the default shell for '$NEW_USERNAME':"
        echo ""
        echo "1. /bin/bash (Bash - Traditional default)"
        echo "2. /bin/zsh (Zsh - macOS Monterey+ default)"
        echo "3. /bin/sh (Bourne Shell - Minimal)"
        echo ""
        echo -e -n "Select shell (1-3) [default: 2]: "
        read -r shell_choice
        
        if [[ -z "$shell_choice" ]]; then
            shell_choice=2
        fi
        
        case "$shell_choice" in
            1)
                NEW_SHELL="/bin/bash"
                echo "Selected: Bash (/bin/bash)"
                log_message "INFO" "Shell selected: /bin/bash"
                break
                ;;
            2)
                NEW_SHELL="/bin/zsh"
                echo "Selected: Zsh (/bin/zsh)"
                log_message "INFO" "Shell selected: /bin/zsh"
                break
                ;;
            3)
                NEW_SHELL="/bin/sh"
                echo "Selected: Bourne Shell (/bin/sh)"
                log_message "INFO" "Shell selected: /bin/sh"
                break
                ;;
            *)
                echo -e "${RED}Invalid selection. Please choose 1, 2, or 3.${NC}"
                ;;
        esac
    done
}

# Function to show creation summary
show_creation_summary() {
    echo ""
    echo -e "${BLUE}ACCOUNT CREATION SUMMARY${NC}"
    echo "================================================================"
    echo -e "${CYAN}Username:${NC} $NEW_USERNAME"
    echo -e "${CYAN}Full Name:${NC} $NEW_FULLNAME"
    echo -e "${CYAN}UID:${NC} $NEW_UID"
    echo -e "${CYAN}GID:${NC} $NEW_GID (staff)"
    echo -e "${CYAN}Home Directory:${NC} $NEW_HOME"
    echo -e "${CYAN}Shell:${NC} $NEW_SHELL"
    
    case $ADMIN_LEVEL in
        "standard")
            echo -e "${CYAN}Administrative Level:${NC} Standard User"
            echo -e "${CYAN}Group Memberships:${NC} staff"
            ;;
        "admin")
            echo -e "${CYAN}Administrative Level:${NC} Administrator"
            echo -e "${CYAN}Group Memberships:${NC} staff, admin"
            ;;
        "admin_sudo")
            echo -e "${CYAN}Administrative Level:${NC} Administrator + Sudo"
            echo -e "${CYAN}Group Memberships:${NC} staff, admin, wheel"
            ;;
    esac
    
    echo ""
    echo "This will create a new macOS user account with the above settings."
    echo ""
    echo -e -n "${YELLOW}Proceed with account creation? (y/N): ${NC}"
    read -r confirm
    
    if [[ "$confirm" =~ ^[Yy]$ ]]; then
        return 0
    else
        return 1
    fi
}

# Function to create user account
create_user_account() {
    echo -e "${GREEN}CREATING USER ACCOUNT...${NC}"
    echo "================================================================"
    
    log_message "INFO" "Starting user account creation for: $NEW_USERNAME"
    
    # Step 1: Create user record
    echo -e "Step 1: Creating user record in Directory Service..."
    if dscl . -create /Users/"$NEW_USERNAME" 2>/dev/null; then
        echo -e "${GREEN}✓ User record created${NC}"
        log_message "INFO" "User record created successfully"
    else
        echo -e "${RED}✗ Failed to create user record${NC}"
        log_message "ERROR" "Failed to create user record"
        return 1
    fi
    
    # Step 2: Set user properties
    echo -e "Step 2: Setting user properties..."
    
    # Set UID
    if dscl . -create /Users/"$NEW_USERNAME" UniqueID "$NEW_UID" 2>/dev/null; then
        echo -e "${GREEN}✓ UID set to $NEW_UID${NC}"
        log_message "INFO" "UID set: $NEW_UID"
    else
        echo -e "${RED}✗ Failed to set UID${NC}"
        log_message "ERROR" "Failed to set UID"
        return 1
    fi
    
    # Set GID (staff group = 20)
    NEW_GID=20
    if dscl . -create /Users/"$NEW_USERNAME" PrimaryGroupID "$NEW_GID" 2>/dev/null; then
        echo -e "${GREEN}✓ Primary group set to staff (GID: $NEW_GID)${NC}"
        log_message "INFO" "Primary group set: $NEW_GID"
    else
        echo -e "${RED}✗ Failed to set primary group${NC}"
        log_message "ERROR" "Failed to set primary group"
        return 1
    fi
    
    # Set Real Name
    if dscl . -create /Users/"$NEW_USERNAME" RealName "$NEW_FULLNAME" 2>/dev/null; then
        echo -e "${GREEN}✓ Real name set to '$NEW_FULLNAME'${NC}"
        log_message "INFO" "Real name set: $NEW_FULLNAME"
    else
        echo -e "${RED}✗ Failed to set real name${NC}"
        log_message "ERROR" "Failed to set real name"
        return 1
    fi
    
    # Set Home Directory
    if dscl . -create /Users/"$NEW_USERNAME" NFSHomeDirectory "$NEW_HOME" 2>/dev/null; then
        echo -e "${GREEN}✓ Home directory set to $NEW_HOME${NC}"
        log_message "INFO" "Home directory set: $NEW_HOME"
    else
        echo -e "${RED}✗ Failed to set home directory${NC}"
        log_message "ERROR" "Failed to set home directory"
        return 1
    fi
    
    # Set Shell
    if dscl . -create /Users/"$NEW_USERNAME" UserShell "$NEW_SHELL" 2>/dev/null; then
        echo -e "${GREEN}✓ Shell set to $NEW_SHELL${NC}"
        log_message "INFO" "Shell set: $NEW_SHELL"
    else
        echo -e "${RED}✗ Failed to set shell${NC}"
        log_message "ERROR" "Failed to set shell"
        return 1
    fi
    
    # Step 3: Set password
    echo -e "Step 3: Setting user password..."
    if dscl . -passwd /Users/"$NEW_USERNAME" "$NEW_PASSWORD" 2>/dev/null; then
        echo -e "${GREEN}✓ Password set successfully${NC}"
        log_message "INFO" "Password set successfully"
    else
        echo -e "${RED}✗ Failed to set password${NC}"
        log_message "ERROR" "Failed to set password"
        return 1
    fi
    
    return 0
}

# Function to create home directory
create_home_directory() {
    echo -e "Step 4: Creating home directory..."
    
    if mkdir -p "$NEW_HOME" 2>/dev/null; then
        echo -e "${GREEN}✓ Home directory created${NC}"
        log_message "INFO" "Home directory created: $NEW_HOME"
    else
        echo -e "${RED}✗ Failed to create home directory${NC}"
        log_message "ERROR" "Failed to create home directory"
        return 1
    fi
    
    if chown -R "$NEW_UID:20" "$NEW_HOME" 2>/dev/null; then
        echo -e "${GREEN}✓ Home directory ownership set${NC}"
        log_message "INFO" "Home directory ownership set"
    else
        echo -e "${RED}✗ Failed to set home directory ownership${NC}"
        log_message "ERROR" "Failed to set home directory ownership"
        return 1
    fi
    
    if chmod 755 "$NEW_HOME" 2>/dev/null; then
        echo -e "${GREEN}✓ Home directory permissions set${NC}"
        log_message "INFO" "Home directory permissions set"
    else
        echo -e "${RED}✗ Failed to set home directory permissions${NC}"
        log_message "ERROR" "Failed to set home directory permissions"
        return 1
    fi
    
    # Create basic directory structure
    local basic_dirs=("Desktop" "Documents" "Downloads" "Pictures" "Music" "Movies" "Public")
    for dir in "${basic_dirs[@]}"; do
        if mkdir -p "$NEW_HOME/$dir" 2>/dev/null && chown "$NEW_USERNAME:staff" "$NEW_HOME/$dir" 2>/dev/null; then
            echo -e "${GREEN}✓ Created $dir folder${NC}"
        else
            echo -e "${YELLOW}⚠ Could not create $dir folder${NC}"
        fi
    done
    
    return 0
}

# Function to configure group memberships
configure_group_memberships() {
    echo -e "Step 5: Configuring group memberships..."
    
    case $ADMIN_LEVEL in
        "standard")
            echo -e "${CYAN}Standard user - no additional groups${NC}"
            log_message "INFO" "Standard user configuration - no additional groups"
            ;;
        "admin")
            echo -e "Adding to admin group..."
            if dscl . -append /Groups/admin GroupMembership "$NEW_USERNAME" 2>/dev/null; then
                echo -e "${GREEN}✓ Added to admin group${NC}"
                log_message "INFO" "Added to admin group"
            else
                echo -e "${RED}✗ Failed to add to admin group${NC}"
                log_message "ERROR" "Failed to add to admin group"
                return 1
            fi
            ;;
        "admin_sudo")
            echo -e "Adding to admin group..."
            if dscl . -append /Groups/admin GroupMembership "$NEW_USERNAME" 2>/dev/null; then
                echo -e "${GREEN}✓ Added to admin group${NC}"
                log_message "INFO" "Added to admin group"
            else
                echo -e "${RED}✗ Failed to add to admin group${NC}"
                log_message "ERROR" "Failed to add to admin group"
                return 1
            fi
            
            echo -e "Adding to wheel group (sudo privileges)..."
            if dscl . -append /Groups/wheel GroupMembership "$NEW_USERNAME" 2>/dev/null; then
                echo -e "${GREEN}✓ Added to wheel group${NC}"
                log_message "INFO" "Added to wheel group"
            else
                echo -e "${RED}✗ Failed to add to wheel group${NC}"
                log_message "ERROR" "Failed to add to wheel group"
                return 1
            fi
            ;;
    esac
    
    return 0
}

# Function to verify account creation
verify_account_creation() {
    echo -e "Step 6: Verifying account creation..."
    
    if dscl . -read /Users/"$NEW_USERNAME" >/dev/null 2>&1; then
        echo -e "${GREEN}✓ User exists in Directory Service${NC}"
        log_message "INFO" "User verified in Directory Service"
    else
        echo -e "${RED}✗ User not found in Directory Service${NC}"
        log_message "ERROR" "User verification failed"
        return 1
    fi
    
    if [[ -d "$NEW_HOME" ]]; then
        local home_owner=$(ls -ld "$NEW_HOME" 2>/dev/null | awk '{print $3}')
        if [[ "$home_owner" == "$NEW_USERNAME" ]]; then
            echo -e "${GREEN}✓ Home directory exists with correct ownership${NC}"
            log_message "INFO" "Home directory verified"
        else
            echo -e "${YELLOW}⚠ Home directory exists but ownership may be incorrect${NC}"
            log_message "WARNING" "Home directory ownership issue detected"
        fi
    else
        echo -e "${RED}✗ Home directory not found${NC}"
        log_message "ERROR" "Home directory verification failed"
        return 1
    fi
    
    if id "$NEW_USERNAME" >/dev/null 2>&1; then
        echo -e "${GREEN}✓ User ID command works${NC}"
        log_message "INFO" "User ID verification successful"
    else
        echo -e "${RED}✗ User ID command failed${NC}"
        log_message "ERROR" "User ID verification failed"
        return 1
    fi
    
    return 0
}

# Function to show completion summary
show_completion_summary() {
    echo ""
    echo "================================================================"
    echo -e "${GREEN}           ACCOUNT CREATION COMPLETED                     ${NC}"
    echo "================================================================"
    echo ""
    echo "NEW USER ACCOUNT SUMMARY:"
    echo "  Username: $NEW_USERNAME"
    echo "  Full Name: $NEW_FULLNAME"
    echo "  UID: $NEW_UID"
    echo "  Home Directory: $NEW_HOME"
    echo "  Shell: $NEW_SHELL"
    
    case $ADMIN_LEVEL in
        "standard")
            echo "  Administrative Level: Standard User"
            echo "  Group Memberships: staff"
            ;;
        "admin")
            echo "  Administrative Level: Administrator"
            echo "  Group Memberships: staff, admin"
            ;;
        "admin_sudo")
            echo "  Administrative Level: Administrator + Sudo"
            echo "  Group Memberships: staff, admin, wheel"
            ;;
    esac
    
    echo ""
    echo "NEXT STEPS:"
    echo "  1. The user can now log in with username: $NEW_USERNAME"
    echo "  2. Test the account by switching users or logging in"
    echo "  3. Configure additional settings as needed"
    echo ""
    
    if [[ "$ADMIN_LEVEL" != "standard" ]]; then
        echo -e "${YELLOW}SECURITY NOTICE:${NC}"
        echo "  This user has administrative privileges."
        echo "  Ensure they understand proper security practices."
        echo ""
    fi
    
    log_message "INFO" "User account creation completed successfully"
}

# Function to rollback user creation
rollback_user_creation() {
    echo -e "Rolling back user creation..."
    log_message "WARNING" "Starting rollback process"
    
    if dscl . -delete /Users/"$NEW_USERNAME" 2>/dev/null; then
        echo -e "${GREEN}User record removed${NC}"
        log_message "INFO" "User record removed during rollback"
    fi
    
    if [[ -d "$NEW_HOME" ]]; then
        echo -e -n "${YELLOW}Remove home directory '$NEW_HOME'? (y/N): ${NC}"
        read -r remove_home
        if [[ "$remove_home" =~ ^[Yy]$ ]]; then
            if rm -rf "$NEW_HOME" 2>/dev/null; then
                echo -e "${GREEN}Home directory removed${NC}"
                log_message "INFO" "Home directory removed during rollback"
            fi
        fi
    fi
    
    log_message "INFO" "Rollback completed"
}

################################################################################
# USER DELETION FUNCTIONS
################################################################################

# Function to check if user can be safely deleted
check_user_deletion_safety() {
    local username=$1
    local issues=0
    
    echo -e "${BLUE}Performing pre-deletion safety checks...${NC}"
    
    if ! dscl . -read /Users/"$username" &>/dev/null; then
        echo -e "${RED}✗ User does not exist in directory service${NC}"
        return 1
    fi
    
    if who | grep -q "^$username "; then
        echo -e "${RED}✗ User is currently logged in${NC}"
        ((issues++))
    else
        echo -e "${GREEN}✓ User is not currently logged in${NC}"
    fi
    
    local process_count=$(ps -U "$username" 2>/dev/null | wc -l)
    process_count=$((process_count - 1))
    if [[ $process_count -gt 0 ]]; then
        echo -e "${YELLOW}⚠ User has $process_count running processes (will be terminated)${NC}"
    else
        echo -e "${GREEN}✓ No running processes for user${NC}"
    fi
    
    local admin_privs=$(check_admin_privileges "$username")
    if [[ -n "$admin_privs" && "$admin_privs" != "none" ]]; then
        if [[ "$admin_privs" == *"ADMIN"* ]]; then
            echo -e "${RED}⚠ User has ADMIN privileges${NC}"
            ((issues++))
        fi
        if [[ "$admin_privs" == *"SUDO"* ]]; then
            echo -e "${YELLOW}⚠ User has SUDO privileges${NC}"
        fi
    else
        echo -e "${GREEN}✓ User has no administrative privileges${NC}"
    fi
    
    echo ""
    
    if [[ $issues -gt 0 ]]; then
        echo -e "${RED}Found $issues critical issues that prevent safe deletion.${NC}"
        return 1
    fi
    
    return 0
}

# Function to confirm deletion
confirm_deletion() {
    local username=$1
    
    if [[ "$username" == "$CURRENT_USER" ]]; then
        echo -e "${RED}ERROR: Cannot delete the currently logged-in user ($CURRENT_USER)!${NC}"
        echo -e "${RED}This would lock you out of the system and cause system instability.${NC}"
        return 1
    fi
    
    if ! check_user_deletion_safety "$username"; then
        echo -e "${RED}Cannot proceed with deletion due to safety issues.${NC}"
        return 1
    fi
    
    if who | grep -q "^$username "; then
        echo -e "${RED}ERROR: User '$username' is currently logged in!${NC}"
        echo -e "${RED}Cannot delete a user with active login sessions.${NC}"
        return 1
    fi
    
    echo -e "${RED}DANGER: You are about to DELETE user account '$username'${NC}"
    echo -e "${RED}This action will:${NC}"
    echo -e "${RED}  - Remove the user account${NC}"
    echo -e "${RED}  - Optionally remove the home directory${NC}"
    echo -e "${RED}  - Remove user from all groups${NC}"
    echo -e "${RED}  - This action CANNOT be undone!${NC}"
    echo ""
    
    read -p "Are you absolutely sure you want to delete user '$username'? (type 'DELETE' to confirm): " confirm
    if [[ "$confirm" != "DELETE" ]]; then
        echo -e "${YELLOW}Deletion cancelled.${NC}"
        return 1
    fi
    
    read -p "Do you want to remove the home directory as well? (y/N): " remove_home
    
    return 0
}

# Function to delete user
delete_user() {
    local username=$1
    
    echo -e "${YELLOW}Deleting user: $username${NC}"
    
    # Kill user processes first
    local process_count=$(ps -U "$username" 2>/dev/null | wc -l)
    process_count=$((process_count - 1))
    if [[ $process_count -gt 0 ]]; then
        echo "Terminating user processes..."
        pkill -TERM -u "$username" 2>/dev/null || true
        sleep 2
        pkill -KILL -u "$username" 2>/dev/null || true
    fi
    
    # Get home directory
    local home_dir=$(dscl . -read /Users/"$username" NFSHomeDirectory 2>/dev/null | awk '{print $2}')
    if [[ -z "$home_dir" ]]; then
        home_dir="/Users/$username"
    fi
    
    echo "Detected home directory: $home_dir"
    
    # Remove from groups
    echo "Removing user from groups..."
    if dscl . -read /Groups/admin GroupMembership 2>/dev/null | grep -q "$username"; then
        dscl . -delete /Groups/admin GroupMembership "$username" 2>/dev/null || true
    fi
    
    # Remove user account
    echo "Removing user account from directory service..."
    if dscl . -delete /Users/"$username" 2>/dev/null; then
        echo -e "${GREEN}User account '$username' removed from directory service.${NC}"
        log_message "INFO" "User deleted: $username"
    else
        echo -e "${RED}Error: Failed to delete user '$username' from directory service${NC}"
        return 1
    fi
    
    # Handle home directory removal if requested
    if [[ "$remove_home" =~ ^[Yy]$ ]]; then
        if [[ -d "$home_dir" ]]; then
            echo "Removing home directory: $home_dir"
            chown -R root:wheel "$home_dir" 2>/dev/null || true
            if rm -rf "$home_dir" 2>/dev/null; then
                echo -e "${GREEN}Home directory removed successfully.${NC}"
                log_message "INFO" "Home directory removed: $home_dir"
            else
                echo -e "${YELLOW}Some protected files remain due to System Integrity Protection.${NC}"
            fi
        else
            echo -e "${GREEN}Home directory was already removed.${NC}"
        fi
    else
        if [[ -d "$home_dir" ]]; then
            echo -e "${YELLOW}Home directory preserved at: $home_dir${NC}"
        fi
    fi
    
    echo -e "${GREEN}User '$username' has been successfully deleted.${NC}"
    log_message "INFO" "User deletion completed: $username"
    return 0
}

################################################################################
# ORPHANED HOME DIRECTORY CLEANUP
################################################################################

# Function to clean up orphaned home directories
cleanup_orphaned_homes() {
    echo -e "${GREEN}Checking for orphaned home directories...${NC}"
    
    for dir in /Users/*/; do
        if [[ -d "$dir" ]]; then
            local dirname=$(basename "$dir")
            if [[ "$dirname" != "Shared" && "$dirname" != "Guest" && "$dirname" != ".localized" ]]; then
                if ! dscl . -read /Users/"$dirname" &>/dev/null; then
                    local owner_uid=$(stat -f "%u" "$dir" 2>/dev/null)
                    echo -e "${YELLOW}Found orphaned directory: $dir (UID: $owner_uid)${NC}"
                    
                    read -p "Remove orphaned directory '$dir'? (y/N): " remove_orphan
                    if [[ "$remove_orphan" =~ ^[Yy]$ ]]; then
                        echo "Removing orphaned directory: $dir"
                        chown -R root:wheel "$dir" 2>/dev/null
                        if rm -rf "$dir" 2>/dev/null; then
                            echo -e "${GREEN}Orphaned directory removed successfully.${NC}"
                            log_message "INFO" "Orphaned directory removed: $dir"
                        else
                            echo -e "${RED}Failed to remove orphaned directory.${NC}"
                        fi
                    fi
                fi
            fi
        fi
    done
    
    echo ""
}

################################################################################
# WORKFLOW FUNCTIONS
################################################################################

# Function to handle user creation workflow
create_user_workflow() {
    echo -e "${BLUE}CREATING NEW USER ACCOUNT${NC}"
    echo "================================================================"
    echo ""
    
    # Step 1: Get user information
    echo -e "${BLUE}STEP 1: USER INFORMATION COLLECTION${NC}"
    get_username_input
    get_fullname_input
    get_password_input
    get_admin_level
    get_shell_input
    
    # Step 2: Prepare account settings
    echo ""
    echo -e "${BLUE}STEP 2: ACCOUNT PREPARATION${NC}"
    echo "================================================================"
    
    echo -e "${GREEN}Finding next available UID...${NC}"
    find_next_uid true  # Run with verbose output
    NEW_UID=$(find_next_uid false)  # Get UID silently for variable assignment
    if [[ -z "$NEW_UID" ]]; then
        echo -e "${RED}Could not find available UID. Exiting.${NC}"
        return 1
    fi
    echo -e "${GREEN}Available UID found: $NEW_UID${NC}"
    
    NEW_HOME="/Users/$NEW_USERNAME"
    echo -e "${GREEN}Home directory will be: $NEW_HOME${NC}"
    
    # Step 3: Final confirmation
    echo -e "${BLUE}STEP 3: FINAL CONFIRMATION${NC}"
    if ! show_creation_summary; then
        echo -e "${YELLOW}Account creation cancelled.${NC}"
        return 0
    fi
    
    # Step 4: Create the account
    echo ""
    echo -e "${BLUE}STEP 4: ACCOUNT CREATION${NC}"
    
    if create_user_account && create_home_directory && configure_group_memberships && verify_account_creation; then
        show_completion_summary
        return 0
    else
        echo -e "${RED}Account creation failed. Starting rollback...${NC}"
        rollback_user_creation
        echo -e "${RED}Account creation was unsuccessful.${NC}"
        return 1
    fi
}

# Function to handle user details workflow  
show_user_details_workflow() {
    echo -e -n "${YELLOW}Enter username to show details: ${NC}"
    read -r username
    
    if [[ -n "$username" ]]; then
        get_user_details "$username"
    else
        echo -e "${RED}No username provided.${NC}"
    fi
}

# Function to handle user deletion workflow
delete_user_workflow() {
    echo -e "${CYAN}SELECT USER TO DELETE${NC}"
    echo "================================================================"
    
    # Get users with UID >= 501 (macOS standard for regular users)
    local users=()
    local usernames=()
    
    for username in $(dscl . -list /Users | grep -v '^_' | grep -v '^Guest' | grep -v '^nobody'); do
        local uid=$(dscl . -read /Users/"$username" UniqueID 2>/dev/null | awk '{print $2}')
        local gid=$(dscl . -read /Users/"$username" PrimaryGroupID 2>/dev/null | awk '{print $2}')
        local home=$(dscl . -read /Users/"$username" NFSHomeDirectory 2>/dev/null | awk '{print $2}')
        local gecos=$(dscl . -read /Users/"$username" RealName 2>/dev/null | cut -d: -f2- | xargs)
        
        # Skip system users (UID < 501 on macOS)
        if [[ -n "$uid" && $uid -ge 501 ]] && [[ $username != "nobody" ]]; then
            local admin_status=$(check_admin_privileges "$username")
            users+=("$username:$uid:$gid:$home:$gecos:$admin_status")
            usernames+=("$username")
        fi
    done
    
    if [[ ${#users[@]} -eq 0 ]]; then
        echo -e "${YELLOW}No regular users found to delete.${NC}"
        return 1
    fi
    
    # Display users in a numbered list
    printf "%-4s %-6s %-15s %-20s %-6s %-15s %s\n" "No." "Current" "Username" "Full Name" "UID" "Admin Privileges" "Home Directory"
    echo "-------------------------------------------------------------------------------------"
    
    local i=1
    for user_info in "${users[@]}"; do
        IFS=':' read -r username uid gid home gecos admin_status <<< "$user_info"
        
        local current=""
        local color=""
        local admin_color=""
        
        # Set colors and current user indicator
        if [[ "$username" == "$CURRENT_USER" ]]; then
            current="*"
            color="${CYAN}"
        fi
        
        # Set admin privilege colors
        if [[ -n "$admin_status" ]]; then
            if [[ "$admin_status" == *"ADMIN"* ]]; then
                admin_color="${RED}"
            elif [[ "$admin_status" == *"SUDO"* ]]; then
                admin_color="${YELLOW}"
            else
                admin_color="${BLUE}"
            fi
        fi
        
        local admin_display="${admin_status:-"none"}"
        
        # Print the row with appropriate colors
        if [[ -n "$color" ]]; then
            printf "${color}%-4d %-6s %-15s %-20s %-6s ${admin_color}%-15s${color} %s${NC}\n" "$i" "$current" "$username" "$gecos" "$uid" "$admin_display" "$home"
        elif [[ -n "$admin_color" ]]; then
            printf "%-4d %-6s %-15s %-20s %-6s ${admin_color}%-15s${NC} %s\n" "$i" "$current" "$username" "$gecos" "$uid" "$admin_display" "$home"
        else
            printf "%-4d %-6s %-15s %-20s %-6s %-15s %s\n" "$i" "$current" "$username" "$gecos" "$uid" "$admin_display" "$home"
        fi
        ((i++))
    done
    
    echo ""
    if [[ -n "$CURRENT_USER" && "$CURRENT_USER" != "unknown" ]]; then
        echo -e "${CYAN}Note: The current user ($CURRENT_USER) is protected and cannot be deleted.${NC}"
        echo ""
    fi
    
    # Get user selection
    while true; do
        echo -e -n "${YELLOW}Select user number to delete (1-${#users[@]}) or 'q' to quit: ${NC}"
        read -r selection
        
        if [[ "$selection" == "q" || "$selection" == "Q" ]]; then
            echo -e "${GREEN}Deletion cancelled.${NC}"
            return 0
        fi
        
        if [[ "$selection" =~ ^[0-9]+$ ]] && [[ $selection -ge 1 ]] && [[ $selection -le ${#users[@]} ]]; then
            # Valid selection
            local selected_username="${usernames[$((selection-1))]}"
            
            # Prevent deletion of current user
            if [[ "$selected_username" == "$CURRENT_USER" ]]; then
                echo -e "${RED}Error: Cannot delete the current user ($CURRENT_USER)${NC}"
                continue
            fi
            
            echo ""
            echo -e "${BLUE}Selected user: $selected_username${NC}"
            
            get_user_details "$selected_username"
            if confirm_deletion "$selected_username"; then
                delete_user "$selected_username"
            fi
            break
        else
            echo -e "${RED}Invalid selection. Please enter a number between 1 and ${#users[@]} or 'q' to quit.${NC}"
        fi
    done
}

################################################################################
# MAIN FUNCTION
################################################################################

main() {
    # Initialize
    check_root
    get_current_user
    setup_logging
    
    # Check required commands
    for cmd in dscl id ps who; do
        if ! command -v "$cmd" &>/dev/null; then
            echo -e "${RED}Error: Required command '$cmd' not found${NC}"
            exit 1
        fi
    done
    
    while true; do
        show_header
        show_main_menu
        echo -e -n "${YELLOW}Select option (1-6): ${NC}"
        choice=$(get_menu_choice)
        
        case "$choice" in
            1)
                echo ""
                echo -e "${BLUE}Creating New User Account...${NC}"
                echo ""
                create_user_workflow
                ;;
            2)
                echo ""
                echo -e "${BLUE}Listing All Users...${NC}"
                echo ""
                list_all_users
                ;;
            3)
                echo ""
                echo -e "${BLUE}Showing User Details...${NC}"
                echo ""
                show_user_details_workflow
                ;;
            4)
                echo ""
                echo -e "${BLUE}Deleting User Account...${NC}"
                echo ""
                delete_user_workflow
                ;;
            5)
                echo ""
                echo -e "${BLUE}Cleaning Up Orphaned Directories...${NC}"
                echo ""
                cleanup_orphaned_homes
                ;;
            6)
                echo ""
                echo -e "${GREEN}Exiting...${NC}"
                if [[ -n "$LOG_FILE" ]]; then
                    echo "Log file: $LOG_FILE"
                fi
                exit 0
                ;;
            "")
                echo -e "${YELLOW}No option selected. Please try again.${NC}"
                ;;
            *)
                echo -e "${RED}Invalid choice '$choice'. Please select 1-6.${NC}"
                ;;
        esac
        
        echo ""
        echo -e -n "${CYAN}Press Enter to continue...${NC}"
        read -r
    done
}

# Run main function
main "$@"