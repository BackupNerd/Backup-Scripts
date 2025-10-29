#!/bin/bash

# Script to add sample files to existing test users for better merge testing
# Creates both unique files and common files that will conflict during merges

echo "Adding sample files to existing test users for merge testing..."

# Function to add sample files to a user
add_sample_files() {
    local username="$1"
    local fullname="$2"
    local user_home="/Users/$username"
    
    # Check if user home directory exists
    if [ ! -d "$user_home" ]; then
        echo "  ⚠ Home directory $user_home does not exist"
        return 1
    fi
    
    echo "  Adding sample files to $username ($fullname)..."
    
    # Create Documents and Desktop directories if they don't exist
    sudo mkdir -p "$user_home/Documents" "$user_home/Desktop"
    
    # Create unique files (no conflicts during merge)
    echo "Personal notes for $fullname - Updated $(date)" > "/tmp/${username}_personal.txt"
    echo "# ${username}'s Private Configuration File" > "/tmp/${username}_config.conf"
    echo "Welcome to ${username}'s desktop! Last updated: $(date)" > "/tmp/${username}_welcome.txt"
    
    # Create common files (will conflict during merges)
    echo "Project data from $username - Last modified: $(date)" > "/tmp/project.txt"
    echo "# Global settings file managed by $username" > "/tmp/settings.conf"
    echo "TODO List for $username: Complete merge testing, validate conflicts, check backups" > "/tmp/todo.txt"
    echo "Meeting notes from $username's perspective - $(date)" > "/tmp/meeting_notes.txt"
    
    # Create a detailed README with user-specific content
    cat > "/tmp/README.md" << EOF
# User Directory for $fullname

**Username:** $username  
**Last Updated:** $(date)  
**Status:** Active test user  

## Test Files Overview:
This directory contains test files for merge operation validation:

### Unique Files (no conflicts):
- \`${username}_personal.txt\` - Personal notes specific to this user
- \`${username}_config.conf\` - User-specific configuration file  
- \`${username}_welcome.txt\` - Desktop welcome message

### Common Files (will conflict during merge):
- \`project.txt\` - Project data (different content per user)
- \`settings.conf\` - Settings file (user-specific content)
- \`todo.txt\` - TODO list (different priorities per user)
- \`meeting_notes.txt\` - Meeting notes from user's perspective
- \`README.md\` - This file (user-specific documentation)

## Merge Testing Strategy:
- Unique files should merge without conflicts
- Common files will trigger conflict resolution
- README.md will test documentation merging
- Different users have different content for same filenames

## Expected Merge Behaviors:
1. **RENAME strategy:** Common files get timestamp suffixes
2. **SKIP strategy:** Target files preserved, source skipped  
3. **REPLACE strategy:** Source overwrites target
4. **ASK strategy:** Individual decisions per conflict

Generated for merge testing on $(date)
EOF

    # Move files to user directories with proper ownership
    sudo mv "/tmp/${username}_personal.txt" "$user_home/Documents/"
    sudo mv "/tmp/${username}_config.conf" "$user_home/Documents/"
    sudo mv "/tmp/${username}_welcome.txt" "$user_home/Desktop/"
    sudo mv "/tmp/project.txt" "$user_home/Documents/"
    sudo mv "/tmp/settings.conf" "$user_home/Documents/"
    sudo mv "/tmp/todo.txt" "$user_home/Desktop/"
    sudo mv "/tmp/meeting_notes.txt" "$user_home/Documents/"
    sudo mv "/tmp/README.md" "$user_home/Documents/"
    
    # Set proper ownership
    sudo chown -R "$username:staff" "$user_home/Documents/" "$user_home/Desktop/"
    
    # Count files added
    doc_count=$(ls -1 "$user_home/Documents/" 2>/dev/null | wc -l)
    desk_count=$(ls -1 "$user_home/Desktop/" 2>/dev/null | wc -l)
    
    echo "  ✓ Added files: $doc_count in Documents, $desk_count in Desktop"
    
    return 0
}

# Test users to add files to (existing users from our test scenario)
declare -A test_users=(
    ["alice_std"]="Alice Standard"
    ["bob_admin"]="Bob Administrator"  
    ["eve_admin"]="Eve Administrator"
    ["frank_sudo"]="Frank Sudo"
    ["iris_sudo"]="Iris Sudo"
    ["jack_std"]="Jack Standard"
    ["puser"]="Power User"
)

echo ""
echo "Adding sample files to existing test users:"
echo "==========================================="

for username in "${!test_users[@]}"; do
    fullname="${test_users[$username]}"
    
    # Check if user exists
    if id "$username" &>/dev/null; then
        add_sample_files "$username" "$fullname"
    else
        echo "  ⚠ User $username does not exist, skipping"
    fi
done

# Also add files to orphaned directories for merge testing
echo ""
echo "Adding sample files to orphaned directories:"
echo "==========================================="

orphaned_dirs=(
    "/Users/buser"
    "/Users/diana_std" 
    "/Users/henry_admin"
)

for orphaned_dir in "${orphaned_dirs[@]}"; do
    if [ -d "$orphaned_dir" ]; then
        dirname=$(basename "$orphaned_dir")
        echo "  Adding sample files to orphaned directory: $dirname"
        
        # Create directories if needed
        sudo mkdir -p "$orphaned_dir/Documents" "$orphaned_dir/Desktop"
        
        # Add files with orphaned user context
        echo "Orphaned data from $dirname - Last accessed: $(date)" > "/tmp/project.txt"
        echo "# Legacy settings from $dirname account" > "/tmp/settings.conf"  
        echo "OLD TODO items from $dirname user account" > "/tmp/todo.txt"
        echo "Historical data from $dirname - Requires merge" > "/tmp/${dirname}_legacy.txt"
        
        cat > "/tmp/README.md" << EOF
# Orphaned Directory: $dirname

**Directory:** $orphaned_dir  
**Status:** Orphaned (no user account)  
**Last Modified:** $(date)  

## Contents:
This directory contains data from deleted user account '$dirname'.
Files are available for merging into active user accounts.

### Files for merge:
- \`project.txt\` - Legacy project data
- \`settings.conf\` - Old configuration settings
- \`todo.txt\` - Unfinished tasks  
- \`${dirname}_legacy.txt\` - Historical data specific to $dirname
- \`README.md\` - This documentation

## Merge Notes:
- This data should be merged into an active user account
- Common filenames may conflict with existing files
- Legacy data should be preserved during merge operations

Orphaned directory prepared for merge testing.
EOF

        # Move files to orphaned directory
        sudo mv "/tmp/project.txt" "$orphaned_dir/Documents/"
        sudo mv "/tmp/settings.conf" "$orphaned_dir/Documents/"
        sudo mv "/tmp/todo.txt" "$orphaned_dir/Desktop/"
        sudo mv "/tmp/${dirname}_legacy.txt" "$orphaned_dir/Documents/"
        sudo mv "/tmp/README.md" "$orphaned_dir/Documents/"
        
        # Count files
        doc_count=$(ls -1 "$orphaned_dir/Documents/" 2>/dev/null | wc -l)
        desk_count=$(ls -1 "$orphaned_dir/Desktop/" 2>/dev/null | wc -l)
        
        echo "  ✓ Added files: $doc_count in Documents, $desk_count in Desktop"
    else
        echo "  ⚠ Orphaned directory $orphaned_dir does not exist"
    fi
done

echo ""
echo "Sample file generation completed!"
echo ""
echo "File Testing Strategy:"
echo "====================="
echo "✓ Unique files: Each user has files named with their username (no conflicts)"
echo "✓ Common files: Multiple users have same filenames with different content"
echo "✓ Documentation: README.md files with user-specific information"
echo "✓ Orphaned data: Legacy files in orphaned directories ready for merge"
echo ""
echo "This setup will test all merge conflict scenarios:"
echo "- RENAME: Adds timestamps to conflicting source files"
echo "- SKIP: Preserves target files, ignores source conflicts"  
echo "- REPLACE: Overwrites target with source files"
echo "- ASK: Prompts for individual conflict decisions"