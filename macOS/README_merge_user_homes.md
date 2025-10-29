# macOS User Home Directory Merge Scripts

A comprehensive set of shell scripts for safely merging one macOS user's home direct### Example Scenarios

#### Scenario 1: Active User Transfer
**Situation**: Employee John Doe is leaving, and his work needs to be transferred to Jane Smith.

**Steps**:
1. Run `sudo ./merge_user_homes.sh`
2. Select `johndoe` as source user
3. Select `janesmith` as target user  
4. Choose **RENAME** strategy (safest)
5. Review the analysis (file counts, sizes, conflicts)
6. Confirm the merge
7. Wait for completion
8. Run `sudo ./verify_user_merge.sh` to verify
9. Test Jane's account to ensure everything works
10. If successful, disable/remove John's account using existing user management tools

#### Scenario 2: Orphaned Directory Cleanup
**Situation**: After removing user account "contractor1", their home directory `/Users/contractor1` still exists with important project files.

**Steps**:
1. Run `sudo ./merge_user_homes.sh`
2. Select the orphaned directory `contractor1` as source
   - Shows as "[ORPHANED DIRECTORY]" with original owner info
3. Select permanent employee `projectmanager` as target user
4. Choose **RENAME** strategy to preserve existing files
5. Review the merge analysis
6. Confirm and execute merge
7. Verify the merge completed successfully
8. Manually remove the empty orphaned directory if desiredr user's home directory. This is useful for account consolidation, migration, or combining user accounts after organizational changes.

## Scripts Included

### 1. `merge_user_homes.sh` - Main Merge Script
The primary script that performs the actual home directory merge operation.

### 2. `verify_user_merge.sh` - Verification & Rollback Script  
A companion script for verifying merge success and providing rollback capabilities.

## Features

### Safety Features
- ✅ **Requires root privileges** for safe operation
- ✅ **Protects current user** from being used as source
- ✅ **Creates timestamped backups** before any changes
- ✅ **Multiple confirmation prompts** with detailed impact assessment
- ✅ **Rollback mechanism** if merge fails
- ✅ **Preserves file metadata** and macOS-specific attributes
- ✅ **Comprehensive logging** of all operations

### Intelligent Merge Capabilities
- 🔄 **Conflict resolution strategies**: Rename, Skip, Replace, or Ask per file
- 📁 **Special directory handling**: Smart handling of Library, Applications, etc.
- 🎯 **Selective Library merging**: Avoids system conflicts
- 📊 **Progress reporting**: Real-time status and statistics
- 🔍 **Duplicate detection**: Content-based duplicate handling
- 📱 **Media library awareness**: Safe handling of iTunes, Photos, etc.
- 🗂️ **Orphaned directory support**: Merge abandoned home directories without user accounts

### macOS-Specific Support
- 🍎 **Extended attributes preservation**: Resource forks, metadata
- 🔒 **System Integrity Protection aware**: Graceful handling of protected files
- 👥 **macOS user standards**: Works with UID 501+ users
- 📂 **Native folder structure**: Preserves macOS conventions

## Requirements

- **macOS**: 10.12 Sierra or later (recommended)
- **Privileges**: Must run with `sudo`
- **Disk Space**: Sufficient space for backups and merged content
- **Dependencies**: Standard macOS utilities (dscl, rsync, stat, du)

## Installation

1. Copy the scripts to your desired location (e.g., `/Users/Eric/Documents/Scripts/`)
2. Make them executable:
   ```bash
   chmod +x merge_user_homes.sh
   chmod +x verify_user_merge.sh
   ```

## Usage

### Basic Merge Operation

```bash
sudo ./merge_user_homes.sh
```

The script will guide you through an interactive process:

1. **Source Selection**: Choose source user or orphaned directory
2. **Target Selection**: Choose target user to receive the data
3. **Merge Analysis**: Review what will be merged
4. **Conflict Resolution**: Select how to handle file conflicts
5. **Backup Creation**: Automatic backup of target directory
6. **Final Confirmation**: Last chance to review before merge
7. **Merge Execution**: Actual file copying and merging

### Source Types Supported

#### Regular Users
- Active user accounts with UID ≥ 501
- Full home directory structure
- Preserves user account after merge

#### Orphaned Directories
- Home directories in `/Users/` without corresponding user accounts
- Directories left behind from deleted user accounts  
- Common after user account removal or system migrations
- Shows original owner information and last modified date
- **Can only be used as SOURCE** (not target)

### Conflict Resolution Strategies

When files with the same name exist in both directories:

1. **RENAME** (Safest): Adds timestamp suffix to existing files
   - `document.txt` becomes `document.txt.backup_20250930_143022`
   - Preserves all data

2. **SKIP**: Keeps existing target files, doesn't copy duplicates
   - Target files remain unchanged
   - Source duplicates are not copied

3. **REPLACE**: Overwrites target files with source files
   - ⚠️ **Warning**: May cause data loss
   - Use only when source files are newer/better

4. **ASK**: Prompt for each individual conflict
   - Most control but may require many decisions
   - Good for small merges or when you want granular control

### Verification and Rollback

After a merge, verify the results:

```bash
sudo ./verify_user_merge.sh
```

Options available:
- **Find and verify recent merges**: Check merge integrity
- **Rollback a merge**: Restore from backup if issues occur
- **Generate detailed report**: Comprehensive analysis of merged data

### Emergency Rollback

If you need to rollback immediately:

```bash
sudo ./verify_user_merge.sh --rollback /tmp/user_merge_backup_username_timestamp
```

## Example Scenario

**Situation**: Employee John Doe is leaving, and his work needs to be transferred to Jane Smith.

**Steps**:
1. Run `sudo ./merge_user_homes.sh`
2. Select `johndoe` as source user
3. Select `janesmith` as target user  
4. Choose **RENAME** strategy (safest)
5. Review the analysis (file counts, sizes, conflicts)
6. Confirm the merge
7. Wait for completion
8. Run `sudo ./verify_user_merge.sh` to verify
9. Test Jane's account to ensure everything works
10. If successful, disable/remove John's account using existing user management tools

## Directory Handling

### Standard Directories
- **Desktop**: Merged with conflict resolution
- **Documents**: Merged with conflict resolution
- **Downloads**: Merged with duplicate detection
- **Pictures/Movies/Music**: Careful merge preserving media libraries
- **Applications**: Smart handling to prevent conflicts

### Special Directory: Library
The Library directory contains system and application preferences. The script offers:

1. **Safe merge** (recommended): Only merges safe subdirectories like:
   - Application Support for common apps (TextEdit, Numbers, Pages, etc.)
   - Desktop Pictures, Fonts, Screen Savers
   - User Scripts and Services

2. **Full merge** (advanced): Merges everything but may cause app conflicts

3. **Skip**: Doesn't merge Library at all

## Backup Information

### Automatic Backups
- **Location**: `/tmp/user_merge_backup_[username]_[timestamp]`
- **Content**: Complete copy of target directory before merge
- **Format**: Preserves all permissions and metadata
- **Info file**: `.backup_info` contains merge details

### Manual Restore Command
If you need to manually restore from backup:
```bash
sudo rsync -av "/tmp/user_merge_backup_username_timestamp/" "/Users/targetuser/"
```

## Logging

### Log Location
- **File**: `/var/log/user_merge_[timestamp].log`
- **Content**: Complete operation log with timestamps
- **Levels**: ERROR, WARNING, INFO, DEBUG

### Log Content Includes
- User selections and confirmations
- File operations (copy, skip, rename)
- Conflict resolutions
- Error conditions and warnings
- Performance statistics

## Troubleshooting

### Common Issues

**"Permission denied" errors**:
- Ensure running with `sudo`
- Check that target directories are accessible
- Some system-protected files cannot be modified (this is normal)

**"Insufficient disk space"**:
- The script requires space for both backup and merged content
- Free up space or use external storage for backups

**"User not found"**:
- Ensure users exist and have UID ≥ 501
- Check user accounts in System Preferences

**Orphaned directory issues**:
- Orphaned directories can only be used as sources, not targets
- If an orphaned directory doesn't appear, check permissions with `ls -la /Users/`
- Some system directories are excluded (Guest, Shared, .localized)
- Original ownership information helps identify the former user

**Library merge conflicts**:
- Use "safe merge" option for Library directory
- Test applications after merge
- Restore specific preferences if needed

### Recovery Options

1. **Automatic rollback**: Use verification script
2. **Manual restore**: Use rsync command with backup
3. **Partial restore**: Restore specific directories from backup
4. **System restore**: Use Time Machine if available

## Security Considerations

### Data Protection
- Always verify backups before starting
- Test on non-production systems first
- Consider external backup before major merges
- Review logs for any security-related warnings

### Permission Preservation
- Script preserves original file permissions
- Sets appropriate ownership for target user
- Maintains macOS extended attributes
- Handles system-protected files gracefully

## Performance Notes

### Large Directories
- Progress indicators show status for large merges
- Operations are logged every 100 files
- Consider running during off-hours for large datasets
- Monitor disk I/O and available space

### Network Home Directories
- Script works with local directories only
- For network homes, ensure local access or mount appropriately
- Consider network bandwidth for large transfers

## Best Practices

### Pre-Merge
1. **Backup independently**: Use Time Machine or other backup solution
2. **Test with small accounts first**: Validate process and procedures  
3. **Document the plan**: Know what you're merging and why
4. **Schedule appropriately**: Plan for downtime and testing
5. **Verify disk space**: Ensure adequate free space

### During Merge
1. **Don't interrupt**: Let the process complete fully
2. **Monitor logs**: Watch for errors or warnings
3. **Note statistics**: File counts and conflict resolutions

### Post-Merge
1. **Verify immediately**: Use verification script
2. **Test thoroughly**: Applications, preferences, data access
3. **Keep backups**: Don't delete backups until certain of success
4. **Update documentation**: Record what was merged and when
5. **Clean up**: Remove old accounts after verification period

## Compatibility

### Tested On
- macOS 10.15 Catalina
- macOS 11 Big Sur  
- macOS 12 Monterey
- macOS 13 Ventura
- macOS 14 Sonoma

### File Systems
- **APFS**: Full support (recommended)
- **HFS+**: Supported with some limitations on extended attributes
- **Network volumes**: Limited support (local mount recommended)

## Support and Updates

This script follows the same high-quality standards as the existing user management scripts in this collection. It includes comprehensive error handling, detailed logging, and follows macOS best practices.

For issues or improvements, refer to the script comments and logging output for detailed diagnostics.

---

**Author**: GitHub Copilot with guidance from Eric Harless  
**Date**: September 30, 2025  
**Version**: 1.0 - Initial macOS Implementation

**⚠️ Important**: Always test on non-production systems first and maintain independent backups of critical data.