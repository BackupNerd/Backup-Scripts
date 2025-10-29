#!/bin/bash

# Script to create 10 test users with various permission levels
# This will be used to test the user management and merge functionality

echo "Creating 10 test users with various permission levels..."

# Array of user data: username:fullname:password:admin_level:shell
users=(
    "alice_std:Alice Standard:TestPass123:1:1"      # Standard user, bash
    "bob_admin:Bob Administrator:TestPass456:2:2"   # Admin, zsh  
    "charlie_sudo:Charlie Sudo:TestPass789:3:1"     # Admin+Sudo, bash
    "diana_std:Diana Standard:TestPass321:1:2"      # Standard user, zsh
    "eve_admin:Eve Administrator:TestPass654:2:1"   # Admin, bash
    "frank_sudo:Frank Sudo:TestPass987:3:2"         # Admin+Sudo, zsh
    "grace_std:Grace Standard:TestPass147:1:1"      # Standard user, bash
    "henry_admin:Henry Administrator:TestPass258:2:2" # Admin, zsh
    "iris_sudo:Iris Sudo:TestPass369:3:1"           # Admin+Sudo, bash
    "jack_std:Jack Standard:TestPass741:1:2"        # Standard user, zsh
)

echo "User creation plan:"
echo "===================="
echo "Standard Users (no admin): alice_std, diana_std, grace_std, jack_std"
echo "Admin Users: bob_admin, eve_admin, henry_admin"  
echo "Admin+Sudo Users: charlie_sudo, frank_sudo, iris_sudo"
echo ""

for user_data in "${users[@]}"; do
    IFS=':' read -r username fullname password admin_level shell <<< "$user_data"
    
    echo "Creating user: $username ($fullname) - Admin Level: $admin_level, Shell: $shell"
    
    # Create user with enhanced_user_manager.sh by piping inputs
    printf "1\n%s\n%s\n%s\n%s\n%s\n%s\ny\n" \
        "$username" \
        "$fullname" \
        "$password" \
        "$password" \
        "$admin_level" \
        "$shell" | sudo /Users/Eric/Documents/Scripts/enhanced_user_manager.sh >/dev/null 2>&1
    
    if [ $? -eq 0 ]; then
        echo "✓ Successfully created: $username"
        
        # Add sample files for merge testing
        echo "  Adding sample files for merge testing..."
        user_home="/Users/$username"
        
        # Create unique files (no conflicts)
        echo "Personal notes for $fullname" > "$user_home/Documents/${username}_personal.txt"
        echo "# ${username}'s Private Settings" > "$user_home/Documents/${username}_config.conf"
        echo "Welcome to ${username}'s desktop!" > "$user_home/Desktop/${username}_welcome.txt"
        
        # Create common files (will conflict during merges)
        echo "Project data from $username - $(date)" > "$user_home/Documents/project.txt"
        echo "# Settings file created by $username" > "$user_home/Documents/settings.conf"
        echo "TODO: Complete tasks - $username" > "$user_home/Desktop/todo.txt"
        echo "Meeting notes from $username's perspective" > "$user_home/Documents/meeting_notes.txt"
        
        # Create a shared file with different content per user
        cat > "$user_home/Documents/README.md" << EOF
# User Directory for $fullname

**Username:** $username  
**Created:** $(date)  
**Admin Level:** $admin_level  
**Shell:** $([[ $shell == "1" ]] && echo "bash" || echo "zsh")  

## Files in this directory:
- ${username}_personal.txt (unique to this user)
- ${username}_config.conf (unique to this user)  
- project.txt (common filename - will conflict)
- settings.conf (common filename - will conflict)
- meeting_notes.txt (common filename - will conflict)

This directory contains test data for merge operations.
EOF
        
        # Set ownership of all created files
        sudo chown -R "$username:staff" "$user_home/Documents/"*
        sudo chown -R "$username:staff" "$user_home/Desktop/"*
        
        echo "  ✓ Added sample files (3 unique, 4 common conflict files)"
    else
        echo "✗ Failed to create: $username"
    fi
    
    sleep 1  # Brief pause between creations
done

echo ""
echo "Test user creation completed!"
echo "Listing all users to verify..."
printf "2\n6\n" | sudo /Users/Eric/Documents/Scripts/enhanced_user_manager.sh