# ClientTool PowerShell Module

A comprehensive PowerShell module wrapper for the N-able Backup Manager (Cove Data Protection) `ClientTool.exe` command-line interface.

## Overview

This module provides PowerShell cmdlets that wrap all available ClientTool.exe commands, making it easier to:
- Manage backup selections and schedules
- Monitor backup/restore sessions
- Configure datasources (MySQL, Oracle, Network Shares, etc.)
- Automate backup operations
- Integrate with PowerShell scripts and workflows

## 📚 Documentation Guide

This module includes comprehensive documentation to help you get started quickly:

- **QUICKSTART.md** - 3-step getting started guide with the most common commands and quick tips
- **QUICK-REFERENCE.md** - Complete command reference organized by category with code examples, pipeline tips, and common workflows
- **CLIENTTOOL-COMMANDS.md** - Detailed reference for underlying ClientTool.exe commands with all arguments and examples
- **README.md** (this file) - Full module documentation with installation instructions and detailed usage examples
- **Examples.ps1** - Runnable demonstration script with 10 working examples showing actual output

Choose the documentation that fits your needs:
- New to the module? Start with **QUICKSTART.md**
- Need a specific command? Check **QUICK-REFERENCE.md**
- Want to see it in action? Run **Examples.ps1**
- Need underlying command details? See **CLIENTTOOL-COMMANDS.md**
- Need comprehensive details? Read this **README.md**

## Installation

### Option 1: Quick Install (Recommended)

Run the install script to copy the module to your PowerShell modules directory:

```powershell
# Navigate to the module directory
cd "C:\Script Root\0-Script Master\Modules\ClientTool"

# Install for current user (no admin required)
.\Install-Module.ps1

# Or install for all users (requires admin)
.\Install-Module.ps1 -Scope AllUsers
```

After installation, you can import the module from any PowerShell session:

```powershell
Import-Module ClientTool
```

### Option 2: Direct Import

Import directly from the module directory without installing:

```powershell
# Import from the module directory
Import-Module "C:\Script Root\0-Script Master\Modules\ClientTool\ClientTool.psd1"
```

### Requirements

- PowerShell 5.1 or higher
- N-able Backup Manager installed (default path: `C:\Program Files\Backup Manager\ClientTool.exe`)
- Appropriate permissions to run ClientTool.exe

## Configuration

By default, the module looks for `ClientTool.exe` at:
```
C:\Program Files\Backup Manager\ClientTool.exe
```

If your installation is in a different location, update the path:

```powershell
Set-ClientToolPath -Path "D:\Custom\Path\ClientTool.exe"
```

## Usage Examples

### Backup Operations

```powershell
# Start backup for all datasources
Start-ClientToolBackup

# Start backup for specific datasource
Start-ClientToolBackup -DataSource FileSystem

# Start backup non-interactively
Start-ClientToolBackup -DataSource FileSystem -NonInteractive
```

### Managing Selections

```powershell
# List all backup selections
Get-ClientToolSelection

# List selections for specific datasource
Get-ClientToolSelection -DataSource FileSystem

# Add a folder to backup
Set-ClientToolSelection -DataSource FileSystem -Include "C:\Data"

# Add folder with exclusion and high priority
Set-ClientToolSelection -DataSource FileSystem `
    -Include "C:\Data" `
    -Exclude "C:\Data\Temp" `
    -Priority High

# Add multiple paths
Set-ClientToolSelection -DataSource FileSystem `
    -Include @("C:\Data", "C:\Users") `
    -Exclude @("C:\Data\Cache", "C:\Users\Public\Temp")

# Clear all selections (requires confirmation)
Clear-ClientToolSelection
```

### Monitoring Sessions

```powershell
# List all backup/restore sessions
Get-ClientToolSession

# List sessions for specific datasource
Get-ClientToolSession -DataSource FileSystem

# Get session errors
Get-ClientToolSessionError

# View session nodes
Get-ClientToolSessionNode

# Stop running session (requires confirmation)
Stop-ClientToolSession
```

### Status Information

```powershell
# Get current program status
Get-ClientToolStatus

# Get application status
Get-ClientToolApplicationStatus

# Get system information (RAM/CPU)
Get-ClientToolSystemInfo

# Check for initialization errors
Get-ClientToolInitializationError
```

### Managing Schedules

```powershell
# List configured schedules
Get-ClientToolSchedule

# Note: New-ClientToolSchedule, Set-ClientToolSchedule, and Remove-ClientToolSchedule
# require additional parameters based on ClientTool's requirements
```

### Database Management

```powershell
# MySQL
Get-ClientToolMySqlServer
# New-ClientToolMySqlServer (requires parameters)
# Set-ClientToolMySqlServer (requires parameters)
# Remove-ClientToolMySqlServer (requires parameters)

# Oracle
Get-ClientToolOracleServer
# Similar add/modify/remove functions available
```

### Network Shares

```powershell
# List network shares
Get-ClientToolNetworkShare

# Add, modify, or remove shares (functions available)
```

### Scripts and Archiving

```powershell
# List custom scripts
Get-ClientToolScript

# List archiving rules
Get-ClientToolArchivingRule

# Functions available for add/modify/remove operations
```

### Settings and Filters

```powershell
# List application settings
Get-ClientToolSetting

# List file system filters
Get-ClientToolFilter
```

### Miscellaneous

```powershell
# Open Backup Manager UI in browser
Open-ClientToolUI

# Get authentication token
Get-ClientToolAuthToken

# Get password requirements
Get-ClientToolPasswordRequirements

# Reset dashboard email subscription
Reset-ClientToolDashboardEmail
```

## Advanced Usage

### Using -WhatIf and -Confirm

Many cmdlets support standard PowerShell risk mitigation parameters:

```powershell
# Preview what would happen
Start-ClientToolBackup -DataSource FileSystem -WhatIf

# Require confirmation
Clear-ClientToolSelection -Confirm

# Bypass confirmation
Remove-ClientToolSchedule -Confirm:$false
```

### Verbose Output

Enable verbose output to see the actual ClientTool.exe commands being executed:

```powershell
Start-ClientToolBackup -DataSource FileSystem -Verbose
```

### Pipeline Support

Output from list commands returns PowerShell objects that can be piped:

```powershell
# Get sessions and filter by datasource
Get-ClientToolSession | Where-Object { $_.DSRC -eq 'FileSystem' }

# Count selections
(Get-ClientToolSelection).Count

# Export to CSV
Get-ClientToolSession -DataSource FileSystem | Export-Csv -Path sessions.csv -NoTypeInformation
```

## Function Naming Convention

The module follows PowerShell naming conventions:

- `Get-ClientTool*` - Retrieves information
- `Set-ClientTool*` - Modifies existing configuration
- `New-ClientTool*` - Creates new items
- `Remove-ClientTool*` - Deletes items
- `Start-ClientTool*` - Initiates operations
- `Stop-ClientTool*` - Terminates operations
- `Clear-ClientTool*` - Removes all items of a type

## Error Handling

The module includes built-in error handling:

```powershell
try {
    Start-ClientToolBackup -DataSource FileSystem
}
catch {
    Write-Error "Backup failed: $_"
}
```

## Getting Help

All functions include detailed help documentation:

```powershell
# Get help for any cmdlet
Get-Help Start-ClientToolBackup -Full
Get-Help Set-ClientToolSelection -Examples
Get-Help Get-ClientToolSession -Detailed

# List all available commands
Get-Command -Module ClientTool
```

## Notes

- Some cmdlets (like `New-ClientToolSchedule`, `Set-ClientToolSetting`, etc.) are framework stubs that require additional parameter implementation based on the specific arguments ClientTool.exe expects. Use `ClientTool.exe help -command <command.name>` to determine required arguments.
- The module automatically handles machine-readable output parsing for list commands
- Non-interactive mode can be enabled for automation scenarios
- All modify/remove operations support `-WhatIf` and `-Confirm` parameters

## Version History

### 1.0.0 (November 14, 2025)
- Initial release
- Complete cmdlet coverage for all ClientTool.exe commands
- Support for all datasources (FileSystem, Exchange, MySQL, Oracle, VMware, VSS, etc.)
- Automatic output parsing for tabular data
- Error handling and validation
- PowerShell best practices implementation

## Contributing

To extend or modify this module:
1. Edit `ClientTool.psm1` for function implementation
2. Update `ClientTool.psd1` to reflect changes
3. Test thoroughly with `Import-Module -Force`

## License

Copyright (c) 2025. All rights reserved.
