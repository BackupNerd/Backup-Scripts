# User Management System Rollback Guide

## Overview

This guide provides comprehensive instructions for rolling back user management operations performed by the enhanced user management and merge system. The system creates automatic backups during all operations, enabling complete restoration to any previous state.

## Backup System Architecture

### Backup Locations
- **Primary Directory**: `/tmp/`
- **Naming Convention**: `user_merge_backup_[username]_[YYYYMMDD_HHMMSS]`
- **Backup Method**: Complete rsync snapshots with preserved permissions

### Types of Backups Created

#### 1. **Merge Operation Backups**
Created automatically before each merge operation:
```
/tmp/user_merge_backup_[target_user]_[timestamp]/
```

#### 2. **User Deletion Backups** 
Created when users are deleted (if home directories preserved):
```
/tmp/user_merge_backup_[deleted_user]_[timestamp]/
```

#### 3. **Incremental State Backups**
Progressive snapshots showing system evolution during mass operations.

## Critical Backup Behavior Information

### What Gets Backed Up During Merge Operations

**❓ IMPORTANT QUESTION**: Are both source and destination backed up during merge?

**✅ ANSWER**: Only the TARGET (destination) user gets backed up

#### Backup Behavior Breakdown:

##### 🎯 TARGET USER (Destination) - ✅ BACKED UP
- **Full backup created** before any changes
- **Location**: `/tmp/user_merge_backup_[target_user]_[timestamp]/`
- **Method**: Complete rsync snapshot of entire home directory
- **Purpose**: Enables complete rollback if something goes wrong

##### 📥 SOURCE USER/DIRECTORY - ❌ NOT BACKED UP
- **No backup created** for source data
- **Source data is COPIED** (not moved) to target
- **Original source remains untouched** in its location
- **Active user accounts remain intact**

#### Why This Design Makes Sense:
1. **Risk Assessment**: The target user's data might be overwritten or modified, so it needs protection
2. **Source Preservation**: Source data is only read from, never modified, so it's inherently safe
3. **Orphaned Directories**: Can't be modified anyway (no user account to modify them)
4. **Active Users**: Source users remain logged in and functional

#### Important Implications:
- ✅ **You CAN rollback the target user completely** using the backup
- ✅ **Source orphaned directories remain available** after merge
- ✅ **Source user accounts (if active) are unchanged**
- ⚠️ **Only the target user is at risk** (and is protected by backup)

#### Code Evidence:
The backup creation function in `merge_user_homes.sh` shows:
```bash
create_backup() {
    echo "CREATING BACKUP OF TARGET DIRECTORY..."
    BACKUP_DIR="/tmp/user_merge_backup_${TARGET_USER}_${MERGE_START_TIME}"
    rsync -av "$TARGET_HOME/" "$BACKUP_DIR/"
}
```

**Key Observations:**
- Only `$TARGET_HOME` is backed up with rsync
- Backup directory named after `$TARGET_USER` only
- Source user/directory is never backed up
- Source data is only READ from, never modified

## Rollback Scenarios & Procedures

### Scenario 1: Complete System Rollback
**Use Case**: Restore admin user to original state before any merges

#### Steps:
1. **Identify Original Backup**:
   ```bash
   # Find the earliest admin backup (original state)
   ls -la /tmp/user_merge_backup_admin_* | head -1
   ```

2. **Stop Any Running User Processes**:
   ```bash
   # Check for admin user processes
   sudo ps -u admin
   # Kill if necessary
   sudo pkill -u admin
   ```

3. **Perform Complete Rollback**:
   ```bash
   # Backup current state (safety measure)
   sudo rsync -av /Users/admin/ /tmp/current_admin_backup_$(date +%Y%m%d_%H%M%S)/
   
   # Restore from original backup
   sudo rsync -av --delete /tmp/user_merge_backup_admin_20251029_064832/ /Users/admin/
   ```

4. **Verify Restoration**:
   ```bash
   # Check directory size and contents
   sudo du -sh /Users/admin
   ls -la /Users/admin/
   ```

### Scenario 2: Partial Rollback (Specific Merge Point)
**Use Case**: Restore admin user to state after specific user merge

#### Steps:
1. **Identify Target Backup**:
   ```bash
   # List all admin backups chronologically
   ls -lt /tmp/user_merge_backup_admin_*
   
   # Choose specific timestamp (e.g., after alice_std merge)
   # user_merge_backup_admin_20251029_064921 = after alice_std
   ```

2. **Review Backup Contents**:
   ```bash
   # Check what was included at that point
   cat /tmp/user_merge_backup_admin_20251029_064921/.backup_info
   ```

3. **Restore to Specific Point**:
   ```bash
   sudo rsync -av --delete /tmp/user_merge_backup_admin_20251029_064921/ /Users/admin/
   ```

### Scenario 3: Individual User Data Recovery
**Use Case**: Recover specific user's original data

#### Steps:
1. **Locate User Backup**:
   ```bash
   # Find specific user backup
   ls -la /tmp/user_merge_backup_alice_std_*
   ```

2. **Create New User Account** (if needed):
   ```bash
   sudo /Users/Eric/Documents/Scripts/UserManage/enhanced_user_manager.sh
   # Choose option 1 (Create User)
   # Enter username: alice_std
   # Follow prompts for user creation
   ```

3. **Restore User Data**:
   ```bash
   # Restore to recreated user account
   sudo rsync -av /tmp/user_merge_backup_alice_std_20251029_054105/ /Users/alice_std/
   
   # Fix ownership
   sudo chown -R alice_std:staff /Users/alice_std/
   ```

### Scenario 4: Selective File Recovery
**Use Case**: Recover specific files without full restoration

#### Steps:
1. **Browse Backup Contents**:
   ```bash
   # List backup contents
   sudo find /tmp/user_merge_backup_admin_20251029_064832 -type f | head -20
   ```

2. **Copy Specific Files**:
   ```bash
   # Recover specific file
   sudo cp /tmp/user_merge_backup_admin_20251029_064832/Documents/important_file.txt /Users/admin/Documents/
   
   # Recover entire directory
   sudo cp -r /tmp/user_merge_backup_admin_20251029_064832/Documents/ProjectFolder/ /Users/admin/Documents/
   ```

3. **Fix Permissions**:
   ```bash
   sudo chown admin:staff /Users/admin/Documents/important_file.txt
   ```

## Advanced Rollback Procedures

### Mass User Recreation
**Use Case**: Recreate all deleted users and restore their data

#### Automated Script:
```bash
#!/bin/bash
# mass_user_restore.sh

USERS=("alice_std" "bob_admin" "eve_admin" "henry_admin" "diana_std" "charlie_sudo" "grace_std")
BACKUP_DIR="/tmp"

for user in "${USERS[@]}"; do
    echo "Processing user: $user"
    
    # Find user backup
    backup=$(ls -d ${BACKUP_DIR}/user_merge_backup_${user}_* 2>/dev/null | head -1)
    
    if [[ -n "$backup" ]]; then
        echo "Found backup: $backup"
        
        # Extract UID from backup info
        uid=$(grep "Original UID" "$backup/.backup_info" 2>/dev/null | cut -d: -f2 | tr -d ' ')
        
        # Recreate user (you'll need to run enhanced_user_manager.sh)
        echo "Create user $user with appropriate permissions"
        
        # Restore data
        sudo rsync -av "$backup/" "/Users/$user/"
        sudo chown -R $user:staff "/Users/$user/"
        
        echo "Restored: $user"
    else
        echo "No backup found for: $user"
    fi
done
```

### Database/Configuration Rollback
**Use Case**: Restore system user database entries

#### Steps:
1. **Check Current Directory Service State**:
   ```bash
   sudo dscl . -list /Users UniqueID | sort -k2 -n
   ```

2. **Manual User Recreation** (if needed):
   ```bash
   # Example for alice_std
   sudo dscl . -create /Users/alice_std
   sudo dscl . -create /Users/alice_std UserShell /bin/bash
   sudo dscl . -create /Users/alice_std RealName "Alice Standard"
   sudo dscl . -create /Users/alice_std UniqueID 504
   sudo dscl . -create /Users/alice_std PrimaryGroupID 20
   sudo dscl . -create /Users/alice_std NFSHomeDirectory /Users/alice_std
   sudo dscl . -passwd /Users/alice_std [password]
   ```

## Backup Validation & Health Checks

### Verify Backup Integrity
```bash
# Check backup completeness
for backup in /tmp/user_merge_backup_admin_*; do
    echo "Checking: $backup"
    if [[ -f "$backup/.backup_info" ]]; then
        echo "✅ Backup info present"
        echo "Created: $(grep 'Backup created' "$backup/.backup_info")"
    else
        echo "❌ Missing backup info"
    fi
done
```

### Test Restore (Dry Run)
```bash
# Perform dry run to test restoration
sudo rsync -av --dry-run /tmp/user_merge_backup_admin_20251029_064832/ /Users/admin/
```

## Emergency Procedures

### Critical System Recovery
If the system is in an unstable state:

1. **Boot to Single User Mode** (if necessary)
2. **Mount filesystems**:
   ```bash
   /sbin/mount -uw /
   ```

3. **Restore from earliest backup**:
   ```bash
   rsync -av /tmp/user_merge_backup_admin_20251029_064832/ /Users/admin/
   ```

### Backup Cleanup (After Successful Operations)
```bash
# Remove old backups (keep recent ones)
find /tmp -name "user_merge_backup_*" -mtime +7 -exec rm -rf {} \;

# Or selectively remove specific backups
rm -rf /tmp/user_merge_backup_admin_20251029_065151
```

## Rollback Verification Checklist

After any rollback operation:

- [ ] **User Authentication**: Test user login capabilities
- [ ] **File Permissions**: Verify file ownership and permissions
- [ ] **Directory Structure**: Confirm all expected directories present
- [ ] **Application Data**: Test application-specific data integrity
- [ ] **System Services**: Ensure no service disruptions
- [ ] **Log Review**: Check system logs for errors

## Best Practices

### Before Rollback
1. **Document Current State**: Record current system status
2. **Create Additional Backup**: Backup current state before rollback
3. **Stop User Processes**: Ensure no active user sessions
4. **Plan Rollback Window**: Schedule during low-usage periods

### During Rollback
1. **Monitor Progress**: Use verbose flags to track restoration
2. **Verify Each Step**: Check intermediate results
3. **Document Actions**: Log all commands executed

### After Rollback
1. **Comprehensive Testing**: Validate all functionality
2. **User Communication**: Notify affected users of changes
3. **Update Documentation**: Record rollback procedure and results

## Troubleshooting Common Issues

### Permission Errors During Restore
```bash
# If rsync fails due to permissions
sudo rsync -av --no-perms --no-owner --no-group /backup/path/ /target/path/
# Then fix permissions separately
sudo chown -R target_user:staff /target/path/
```

### Incomplete Backups
```bash
# Check backup size vs original
du -sh /tmp/user_merge_backup_admin_20251029_064832
du -sh /Users/admin
```

### Restore Command Reference
```bash
# Basic restore
sudo rsync -av /backup/source/ /target/destination/

# Restore with deletion of extra files
sudo rsync -av --delete /backup/source/ /target/destination/

# Restore specific subdirectory
sudo rsync -av /backup/source/Documents/ /target/destination/Documents/

# Dry run (test without changes)
sudo rsync -av --dry-run /backup/source/ /target/destination/
```

## Contact & Support

For additional assistance with rollback procedures:
- Review system logs: `/var/log/user_*.log`
- Check backup info files: `/tmp/user_merge_backup_*/.backup_info`
- Consult the main README: `/Users/Eric/Documents/Scripts/UserManage/README.md`

---

**Important**: Always test rollback procedures in a non-production environment when possible. The backup system is designed for reliability, but verification of restored data is essential for critical operations.

*Document Version: 1.0*  
*Created: October 29, 2025*  
*Based on comprehensive testing of 21 backup operations*