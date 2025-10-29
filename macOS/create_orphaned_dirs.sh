#!/bin/bash

# Script to delete some user accounts while preserving home directories
# This will create orphaned directories for testing merge functionality

echo "Creating orphaned directories by deleting user accounts..."
echo "======================================================="

# Users to delete (will create orphaned directories)
users_to_delete=("charlie_sudo" "grace_std" "henry_admin")

echo "Users to be deleted (home directories will remain):"
for user in "${users_to_delete[@]}"; do
    echo "  - $user"
done
echo ""

# Add some test files to these users' home directories first
echo "Adding test files to user directories before deletion..."
for user in "${users_to_delete[@]}"; do
    if [[ -d "/Users/$user" ]]; then
        echo "Adding test files for $user..."
        sudo mkdir -p "/Users/$user/Documents/TestFiles"
        sudo mkdir -p "/Users/$user/Desktop"
        
        # Create some test files
        echo "Test document for $user" | sudo tee "/Users/$user/Documents/test_document_$user.txt" > /dev/null
        echo "Desktop file for $user" | sudo tee "/Users/$user/Desktop/desktop_note_$user.txt" > /dev/null
        echo "Important data for $user" | sudo tee "/Users/$user/Documents/TestFiles/important_data.txt" > /dev/null
        
        # Set ownership
        sudo chown -R "$user:staff" "/Users/$user/Documents" "/Users/$user/Desktop"
        echo "  ✓ Created test files for $user"
    fi
done

echo ""
echo "Now deleting user accounts (preserving home directories)..."

for user in "${users_to_delete[@]}"; do
    echo "Deleting user account: $user"
    
    # Delete using dscl directly (faster than interactive script)
    if sudo dscl . -delete /Users/"$user" 2>/dev/null; then
        echo "  ✓ Successfully deleted user account: $user"
        echo "  ℹ Home directory preserved at: /Users/$user"
    else
        echo "  ✗ Failed to delete user account: $user"
    fi
done

echo ""
echo "Orphaned directory creation completed!"
echo ""
echo "Verifying orphaned directories..."
for user in "${users_to_delete[@]}"; do
    if [[ -d "/Users/$user" ]]; then
        if ! dscl . -read /Users/"$user" &>/dev/null; then
            size=$(du -sh "/Users/$user" 2>/dev/null | cut -f1)
            echo "  ✓ ORPHANED: /Users/$user ($size)"
        else
            echo "  ✗ NOT ORPHANED: /Users/$user (user account still exists)"
        fi
    else
        echo "  ✗ MISSING: /Users/$user (directory not found)"
    fi
done