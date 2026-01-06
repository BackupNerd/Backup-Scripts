# Cove Data Protection → ConnectWise Manage Integration

Automated ticket management for N-able Cove Data Protection backup monitoring using ConnectWise Manage PSA.

## 📋 Overview

This suite of PowerShell scripts provides complete automation for monitoring Cove backup failures and syncing them to ConnectWise Manage as support tickets. The integration handles ticket creation, updates, and automatic closure when issues are resolved.

## 🔧 Scripts

### 1. **Cove2CWM-SyncTickets.v10.ps1** (Main Monitoring Script)

**Purpose:** Core monitoring engine that tracks backup failures and manages tickets

**What it does:**
- Queries Cove API for device backup status across all customers
- Detects failures, errors, and stale backups based on configurable rules
- Creates ConnectWise tickets for new issues
- Updates existing tickets with current status
- Automatically closes tickets when issues are resolved
- Supports both Server/Workstation (AccountType=1) and M365 (AccountType=2) devices

**Key Features:**
- **Multi-severity monitoring rules:**
  - Critical: Systems devices with backup failures (Server priority)
  - High: Workstation devices with backup failures (Workstation priority)
  - Medium: M365 devices with backup failures (M365 priority)
- **Intelligent error detection:** Queries recent session errors per datasource
- **Stale backup detection:** Configurable thresholds (7/14/30 days)
- **Auto-closure:** Tickets close automatically when backups succeed
- **CSV export:** Saves device status for analysis and handoff to other scripts

**Configuration Parameters:**
```powershell
-PartnerName         # Cove partner to monitor (required)
-TicketBoard         # CWM board name (e.g., "Service Desk")
-TicketNewStatus     # Status for new tickets (e.g., "New")
-TicketClosedStatus  # Status for resolved tickets (e.g., "Closed")
-TicketPriorityServer      # Priority for server failures
-TicketPriorityWorkstation # Priority for workstation failures
-TicketPriorityM365        # Priority for M365 failures
```

**Usage:**
```powershell
# Basic monitoring
.\Cove2CWM-SyncTickets.v10.ps1 -PartnerName "YourPartnerName"

# With custom configuration
.\Cove2CWM-SyncTickets.v10.ps1 -PartnerName "YourPartnerName" `
    -TicketBoard "Service Desk" `
    -TicketPriorityServer "Priority 1 - Emergency Response" `
    -TicketPriorityWorkstation "Priority 3 - Normal Response"
```

**Output:**
- ConnectWise tickets created/updated automatically
- CSV export: `.\tickets\YYYYMMDD_CoveMonitoring_v03.csv`

---

### 2. **Cove2CWM-SetTicketsConfig.ps1** (Configuration Helper)

**Purpose:** Interactive GUI tool to configure monitoring script parameters

**What it does:**
- Loads current monitoring script configuration
- Displays interactive GridView menus for selecting:
  - ConnectWise Board
  - New ticket status
  - Closed ticket status
  - Server priority level (with recommendations)
  - Workstation priority level (with recommendations)
  - M365 priority level (with recommendations)
- Updates parameters in `Cove2CWM-SyncTickets.v10.ps1`
- Validates all changes were applied correctly
- Saves configuration for easy re-use

**Key Features:**
- **Visual configuration:** No manual parameter editing required
- **Recommendations:** Shows suggested priorities based on severity
- **Current values:** Displays checkmarks (✓) for currently selected options
- **Validation:** Reads back parameters after saving to confirm accuracy
- **Error detection:** Alerts if validation fails with detailed mismatch report

**Usage:**
```powershell
# Interactive configuration
.\Cove2CWM-SetTicketsConfig.ps1

# Follow the GridView prompts to select:
# 1. Board (e.g., "Service Desk")
# 2. New Status (e.g., "New")
# 3. Closed Status (e.g., "Closed")
# 4. Server Priority (recommended: Priority 1 - Critical)
# 5. Workstation Priority (recommended: Priority 3 - Normal)
# 6. M365 Priority (recommended: Priority 2 - Quick Response)
```

**GridView Columns:**
- **Priority Name:** Full priority level name
- **Recommended:** Shows ★ for recommended priorities
- **Current:** Shows ✓ for currently selected priority
- **Priority ID:** ConnectWise priority identifier

**Validation Output:**
```
✓ TicketBoard: 'Service Desk'
✓ TicketNewStatus: 'New'
✓ TicketClosedStatus: 'Closed'
✓ TicketPriorityServer: 'Priority 1 - Emergency Response'
✓ TicketPriorityWorkstation: 'Priority 3 - Normal Response'
✓ TicketPriorityM365: 'Priority 2 - Quick Response'

SUCCESS: All parameters validated successfully!
```

---

### 3. **Cove2CWM-SyncCustomers.ps1** (Company Sync Tool)

**Purpose:** Synchronize Cove customers to ConnectWise companies

**What it does:**
- Retrieves all Cove end-customer partners
- Compares against existing ConnectWise companies
- Identifies missing companies using intelligent matching:
  - Exact name match
  - Case-insensitive match
  - Identifier match
  - Reference code match
- Creates missing companies in ConnectWise
- Resolves partner hierarchy (skips Sites, targets End-Customers)

**Key Features:**
- **Multi-strategy matching:** Finds companies even with slight name variations
- **Hierarchy resolution:** Automatically walks up Cove partner tree to End-Customer level
- **Intelligent identifiers:** Generates clean, unique company identifiers
- **CSV input support:** Can process device list from monitoring script
- **Batch creation:** Create multiple companies at once
- **WhatIf mode:** Preview changes without creating companies

**Configuration Parameters:**
```powershell
-PartnerName      # Cove partner to sync (queries all sub-partners)
-CSVPath          # Import from monitoring script CSV export
-CreateCompany    # Create specific company only (skip GridView)
-CompanyStatus    # Default status for new companies (default: "Active")
-CompanyType      # Default type for new companies (default: "Customer")
-WhatIf           # Preview mode - don't create companies
-NonInteractive   # Skip user prompts (for automation)
```

**Usage:**
```powershell
# Interactive analysis - shows GridView with comparison
.\Cove2CWM-SyncCustomers.ps1 -PartnerName "YourPartnerName"

# Use CSV from monitoring script
.\Cove2CWM-SyncCustomers.ps1 -CSVPath ".\tickets\YYYYMMDD_CoveMonitoring_v03.csv"

# Create specific company (called from monitoring script)
.\Cove2CWM-SyncCustomers.ps1 -PartnerName "YourPartnerName" `
    -CreateCompany "Acme Corporation" `
    -NonInteractive

# Preview mode
.\Cove2CWM-SyncCustomers.ps1 -PartnerName "YourPartnerName" -WhatIf
```

**GridView Columns:**
- **CovePartner:** Partner name from Cove
- **CoveReference:** Cove partner reference code
- **MatchStatus:** Matched / Missing
- **CWMCompany:** Matched company name (if found)
- **CWMIdentifier:** Matched company identifier
- **DeviceCount:** Number of devices with issues (if using CSV input)

**Output:**
- CSV comparison file: `.\Companies\CoveToConnectWise_Comparison_YYYYMMDD_HHMMSS.csv`

**Comparison Results:**
```
=== Comparison Summary ===
Total Cove End-Customers: 45
Matched in CWM: 42 (93.3%)
Missing in CWM: 3 (6.7%)
```

---

## 🔄 Script Relationships

```
┌─────────────────────────────────────────────────────────────┐
│                 Cove2CWM-SetTicketsConfig.ps1               │
│                  (Configuration Helper)                     │
│                                                             │
│  • Loads current monitoring script config                   │
│  • Shows GridView menus for parameter selection             │
│  • Updates Cove2CWM-SyncTickets.v10.ps1 parameters          │
│  • Validates changes were applied correctly                 │
└────────────────────────┬────────────────────────────────────┘
                         │
                         │ Configures
                         ▼
┌─────────────────────────────────────────────────────────────┐
│              Cove2CWM-SyncTickets.v10.ps1                   │
│                 (Main Monitoring Script)                    │
│                                                             │
│  • Queries Cove API for device backup status                │
│  • Detects failures and creates/updates CWM tickets         │
│  • Exports CSV with device status                           │
│  • Automatically closes resolved tickets                    │
└────────────────────────┬────────────────────────────────────┘
                         │
                         │ Exports CSV / Calls for missing companies
                         ▼
┌─────────────────────────────────────────────────────────────┐
│              Cove2CWM-SyncCustomers.ps1                     │
│                  (Company Sync Tool)                        │
│                                                             │
│  • Compares Cove customers to CWM companies                 │
│  • Creates missing companies in ConnectWise                 │
│  • Can process CSV from monitoring script                   │
│  • Can be called directly from monitoring script            │
└─────────────────────────────────────────────────────────────┘
```

### Typical Workflow

1. **Initial Setup:**
   ```powershell
   # Configure monitoring parameters
   .\Cove2CWM-SetTicketsConfig.ps1
   ```

2. **First Run - Company Sync:**
   ```powershell
   # Sync customers to ensure all companies exist
   .\Cove2CWM-SyncCustomers.ps1 -PartnerName "YourPartner"
   ```

3. **Ongoing Monitoring:**
   ```powershell
   # Run daily/weekly via Task Scheduler
   .\Cove2CWM-SyncTickets.v10.ps1 -PartnerName "YourPartner"
   ```

4. **Reconfiguration:**
   ```powershell
   # Change ticket priorities or board
   .\Cove2CWM-SetTicketsConfig.ps1
   ```

---

## 🔐 Authentication

All scripts require two credential files (created on first run):

### Cove API Credentials
**File:** `C:\ProgramData\MXB\{computername}_{username}_API_Credentials.Secure.xml`

**Structure:**
- PartnerName (string)
- Username (string)
- Password (DPAPI encrypted string)

### ConnectWise API Credentials
**File:** `C:\ProgramData\MXB\{computername}_{username}_CWM_Ticketing_Credentials.Secure.xml`

**Structure:**
- Server (string)
- Company (string)
- privateKey (DPAPI encrypted)
- pubKey (DPAPI encrypted)
- clientId (DPAPI encrypted)

**Security:** Both files use Windows DPAPI encryption, meaning credentials are tied to the specific user account and machine. Credentials cannot be decrypted on a different machine or by a different user.

---

## 📊 Cove API Column Codes Reference

Key column codes used by monitoring script:

| Code | Description | Usage |
|------|-------------|-------|
| **AU** | Account ID | Device unique identifier |
| **AN** | Account Name | Device name |
| **AR** | Account Reference | Partner name |
| **PF** | Partner Reference | Partner external reference code |
| **AT** | Account Type | 1=Server/Workstation, 2=M365 |
| **TL** | Last Session Time | Unix timestamp of last backup |
| **T0** | Last Session Status | 5=Success, 2=Failed, 8=CompletedWithErrors |
| **I78** | Datasources | Configured datasource codes (D01=FileSystem, D19=Exchange, etc.) |
| **OT** | OS Type | 1=Workstation, 2=Server |

---

## 🎯 Monitoring Rules

### Critical Severity (Server Priority)
- **Device Type:** Servers (OT=2)
- **Trigger:** Backup failure, error, or stale backup
- **Action:** Create ticket with Server priority level
- **Auto-Close:** When backup succeeds

### High Severity (Workstation Priority)
- **Device Type:** Workstations (OT=1)
- **Trigger:** Backup failure, error, or stale backup
- **Action:** Create ticket with Workstation priority level
- **Auto-Close:** When backup succeeds

### Medium Severity (M365 Priority)
- **Device Type:** M365 (AT=2)
- **Trigger:** Exchange, OneDrive, SharePoint, or Teams backup failure
- **Action:** Create ticket with M365 priority level
- **Auto-Close:** When backup succeeds

### Stale Backup Thresholds
- **7 days:** Warning threshold
- **14 days:** Alert threshold
- **30 days:** Critical threshold

---

## 📦 Prerequisites

### PowerShell Modules
- **ConnectWiseManageAPI** (auto-installed if missing)
  ```powershell
  Install-Module -Name ConnectWiseManageAPI -Scope CurrentUser
  ```

### PowerShell Version
- **PowerShell 5.1** (minimum)
- **PowerShell 7.x** (auto-relaunches if available for better performance)

### API Access
- **Cove API:** Partner-level credentials with device query permissions
- **ConnectWise API:** Member credentials with company/ticket create/update permissions

---

## 🗂️ File Structure

```
CDP2PSA.ConnectWiseManage/
└── CDP2CWM-Ticketing/
    ├── Cove2CWM-SyncTickets.v10.ps1      # Main monitoring script
    ├── Cove2CWM-SetTicketsConfig.ps1     # Configuration helper
    ├── Cove2CWM-SyncCustomers.ps1        # Company sync tool
    ├── README.md                          # This file
    └── tickets/                           # CSV exports
        └── YYYYMMDD_CoveMonitoring_v03.csv
```

---

## 🚀 Quick Start

### 1. First Time Setup

```powershell
# REQUIRED: Run the main monitoring script first (creates credentials and handles everything)
.\Cove2CWM-SyncTickets.v10.ps1

# OPTIONAL: Configure monitoring parameters via GUI (if you want to change defaults)
.\Cove2CWM-SetTicketsConfig.ps1

# OPTIONAL: Pre-sync customers to ConnectWise (monitoring script will auto-create if missing)
.\Cove2CWM-SyncCustomers.ps1
```

**Important:** The monitoring script (`Cove2CWM-SyncTickets.v10.ps1`) should always be run first. It will:
- Prompt for credentials on first run and create encrypted credential files
- Automatically detect partner name from your Reseller-level Cove API credentials (no `-PartnerName` parameter needed)
- Automatically call `Cove2CWM-SyncCustomers.ps1` if it encounters a missing ConnectWise company
- Use default ticket board/status/priority settings (which can be changed later with `SetTicketsConfig.ps1`)

### 2. Schedule Automated Monitoring

**IMPORTANT:** The monitoring script must run as the **same user** who first ran the scripts and created the credential files (DPAPI encryption is user+machine specific).

**Create Task Scheduler job:**

```powershell
# Run as the user who created the credentials (e.g., the user who first ran Cove2CWM-SyncTickets.v10.ps1)
$taskUser = "$env:USERDOMAIN\$env:USERNAME"  # Or specify: "DOMAIN\Username"

$action = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File "<ScriptPath>\Cove2CWM-SyncTickets.v10.ps1"'

# Choose one of these trigger options:

# Option 1: Every 4 hours
$trigger = New-ScheduledTaskTrigger -Once -At "12:00AM" -RepetitionInterval (New-TimeSpan -Hours 4) -RepetitionDuration ([TimeSpan]::MaxValue)

# Option 2: Every 12 hours (e.g., 6:00 AM and 6:00 PM)
# $trigger = New-ScheduledTaskTrigger -Once -At "6:00AM" -RepetitionInterval (New-TimeSpan -Hours 12) -RepetitionDuration ([TimeSpan]::MaxValue)

# Option 3: Daily at specific time
# $trigger = New-ScheduledTaskTrigger -Daily -At "6:00AM"

# Create task settings for reliability
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register task - prompts for password to enable "Run whether user is logged on or not"
$credential = Get-Credential -UserName $taskUser -Message "Enter password for scheduled task"
Register-ScheduledTask -TaskName "Cove Backup Monitoring" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -User $taskUser `
    -Password $credential.GetNetworkCredential().Password `
    -RunLevel Highest
```

**Note:** Providing the `-Password` parameter enables "Run whether user is logged on or not" mode. Without it, the task only runs when the user is logged in.

**Trigger Interval Options:**
- **Every 4 hours:** Recommended for active monitoring (6 checks per day)
- **Every 12 hours:** Balanced approach (2 checks per day)
- **Daily:** Minimal monitoring (1 check per day)

**Critical Notes:**
- **User Context:** Task MUST run as the same user who first ran the scripts and created the credentials
- **DPAPI Security:** Credentials are encrypted per-user and per-machine - cannot be decrypted by different users
- **Password Required:** You'll need to provide the user's password when registering the task to enable "Run whether user is logged on or not"
- **Script Execution:** The main monitoring script (`Cove2CWM-SyncTickets.v10.ps1`) will automatically call helper scripts if needed:
  - Calls `Cove2CWM-SyncCustomers.ps1` if a ConnectWise company is missing
  - Uses `Cove2CWM-SetTicketsConfig.ps1` for interactive configuration (run manually as needed)
- **First Run:** Always run `Cove2CWM-SyncTickets.v10.ps1` interactively first to create credential files before scheduling

---

## 🛠️ Troubleshooting

### Issue: "Cove API credentials not found"
**Solution:** Run `Cove2CWM-SyncTickets.v10.ps1` interactively first to create credentials

### Issue: "ConnectWise company not found for device"
**Solution:** Run `Cove2CWM-SyncCustomers.ps1` to create missing companies

### Issue: Validation errors after configuration change
**Solution:** Check `SetTicketsConfig` output for specific parameter mismatch, restore from backup if needed

### Issue: Credentials don't work on different machine
**Solution:** DPAPI encryption is machine+user specific - create new credentials on target machine

---

## 📝 Version History

- **v10.2** (2026-01-06)
  - Fixed CoveReference column using wrong API field (PF vs PN)
  - Added parameter validation to SetTicketsConfig
  - Fixed regex backreference escaping bug
  - Removed Sort Order column from priority GridViews

- **v10.1** (2025-12-26)
  - Expanded session query to check all non-successful statuses
  - Improved M365 session error detection
  - Added configurable stale backup thresholds

- **v10.0** (2025-12-20)
  - Complete rewrite with M365 support
  - Separate monitoring rules by device type
  - Optimized session queries with server-side filtering

---

## 👤 Author

**Eric Harless** - Head Backup Nerd, N-able

## 📄 License

Sample scripts provided as-is. Not supported under any N-able support program or service.

## 🔗 Related Documentation

- [Cove API Documentation](https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/)
- [ConnectWise Manage API](https://developer.connectwise.com/Products/Manage/REST)


