# OffboardClientToolRestore Script

**Bulk file-level recovery from N-able Cove Data Protection using the local Backup Manager.**

---

## Overview

`OffboardClientToolRestore.v##.ps1` automates unattended restoration of multiple backup sessions from the local Backup Manager without requiring cloud portal access or internet connectivity. Designed for bulk historical restores, forensic recovery, and offboarded machine scenarios.

## When to Use This Script

### Primary Use Cases

- **Bulk historical restores** — Restore multiple sessions in sequence (weekly Fridays, end-of-month, etc.) overnight without manual intervention
- **Forensic/compliance restores** — Recover specific point-in-time versions (Daily/Weekly/Monthly filters) from paths affected by ransomware or accidental deletion
- **Offboarded devices** — Machines being decommissioned or already removed from the backup account need their data extracted before account closure
- **Scripted/automated workflows** — Integrate into recovery runbooks where operators select sessions, then walk away

### Why Not Use the GUI?

- GUI restores one session at a time with no automation
- This script handles queuing, progress monitoring, logging, and automatic skip-on-failure across multiple sessions
- Ideal for overnight/unattended batch operations

---

## Prerequisites

### System Requirements

- **Windows 7 or later** (PowerShell 5.1 minimum)
- **N-able Cove Data Protection Backup Manager** installed locally with ClientTool.exe available
- **Local backup sessions** cached on the machine (SessionReport.xml and backup files present)
- **Target restore volume** with sufficient free space
- **Administrator privileges** to run the script

### Recommended Configuration

- **Restore-Only Backup Manager installation** (reduces resource overhead on active backup systems)
- **Local SpeedVault** as restore source (for faster throughput)
- **config.ini optimization**: Add `RestoreDownloadThreadCount=50` to `[General]` section and restart BackupFP service

---

## Installation

1. Copy `OffboardClientToolRestore.v##.ps1` to your script directory
2. Update the default parameters (lines 112-119) with your environment:
   - `-RestoreSelection`: Source paths to restore
   - `-RestoreBase`: Destination directory for restored files
   - `-SessionType`: Filter (All, Daily, Weekly, Monthly)
   - `-DataSource`: Restore type (FileSystem, NetworkShares, VssMsSql)

3. Save and close the script
4. Open PowerShell as Administrator and navigate to the script directory

---

## Usage

### Basic Restore (All Sessions)

```powershell
.\OffboardClientToolRestore.v##.ps1 -RestoreSelection "C:\Users\Shared\Documents" -RestoreBase "D:\RestoredData\"
```

### Restore Weekly Sessions (Fridays Only)

```powershell
.\OffboardClientToolRestore.v##.ps1 `
  -RestoreSelection "C:\Users\Shared\Documents" `
  -RestoreBase "D:\RestoredData\" `
  -SessionType Weekly `
  -Weekday Friday
```

### Restore Multiple Paths Per Session

```powershell
.\OffboardClientToolRestore.v##.ps1 `
  -RestoreSelection "C:\Users\Shared\Documents","C:\CompanyData" `
  -RestoreBase "D:\RestoredData\" `
  -DataSource FileSystem
```

### Restore End-of-Month Sessions

```powershell
.\OffboardClientToolRestore.v##.ps1 `
  -RestoreSelection "D:\Database" `
  -RestoreBase "D:\RestoredData\" `
  -SessionType Monthly `
  -DataSource VssMsSql
```

### Skip Existing Files (Default: Skip)

```powershell
.\OffboardClientToolRestore.v##.ps1 `
  -RestoreSelection "C:\Data" `
  -RestoreBase "D:\RestoredData\" `
  -ExistingFileRestorePolicy "Skip"
```

### Overwrite Existing Files

```powershell
.\OffboardClientToolRestore.v##.ps1 `
  -RestoreSelection "C:\Data" `
  -RestoreBase "D:\RestoredData\" `
  -ExistingFileRestorePolicy "Overwrite"
```

---

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-RestoreSelection` | string[] | `@("D:\data1","D:\data2")` | One or more source paths to restore (each becomes a separate `-selection` argument) |
| `-RestoreBase` | string | `D:\Restore\` | Base directory where restored files are placed |
| `-CombinedRestore` | switch | `$true` | If true, combine all session restores into one directory; if false, create timestamped subdirectories per session |
| `-SessionType` | string | `Weekly` | Filter sessions: `All`, `Daily`, `Weekly`, or `Monthly` |
| `-Weekday` | string | `Friday` | Day filter for Weekly mode: `Sunday` through `Saturday` |
| `-DataSource` | string | `FileSystem` | Restore type: `FileSystem`, `NetworkShares`, or `VssMsSql` |
| `-ExistingFileRestorePolicy` | string | `Skip` | How to handle existing files: `Skip` or `Overwrite` |
| `-OutdatedFileRestorePolicy` | string | `CheckContentOfOutdatedFilesOnly` | How to handle outdated files: `CheckContentOfOutdatedFilesOnly` or `CheckContentOfAllFiles` |
| `-IncludedStates` | string[] | `@("Completed", "CompletedWithErrors")` | Session states to include: `Completed`, `CompletedWithErrors`, `Aborted`, `Failed` |
| `-AddSessionTimestampSuffix` | switch | `$false` | *Deprecated* — append timestamp suffix to restored files |

---

## How It Works

### Three-Phase Monitoring

1. **Idle Wait (120 minutes)**
   - Waits for the Backup Manager engine to reach Idle state
   - Required because ClientTool rejects `restore.start` if a backup or restore is already running
   - Skips session and moves to next if Idle not reached after 120 minutes

2. **Restore Start Confirmation (60 seconds)**
   - Confirms the restore actually started after issuing `control.restore.start`
   - Two independent signals checked each tick:
     - SessionReport.xml: new Restore record with higher Id than pre-restore snapshot
     - ClientTool status: "Restore" state (catches long-running restores still in progress)
   - Detects sub-second/ultra-fast restores that status polling alone would miss
   - Skips session if neither signal appears within 60 seconds

3. **Restore Completion Wait (Unbounded)**
   - Only entered when Phase 1 confirmed restore is running
   - Polls SessionReport.xml every 5 seconds until new session record appears
   - No timeout — waits as long as needed (may take hours for large datasets)

### Logging

- Each restore operation logged to `<RestoreBase>\<DataSource>_Restore_log.txt`
- Logs include: start time, session timestamp, restore path, selected paths, completion status
- Useful for auditing and troubleshooting failed restores

---

## Performance Optimization

### Local SpeedVault

If a SpeedVault cache exists on the restore machine, configure it as the restore source:

```powershell
# In BackupFP config.ini [General] section
RestoreDataSource=SpeedVault
```

### Parallel Download Threads

Increase download concurrency:

```powershell
# In config.ini [General] section
RestoreDownloadThreadCount=50

# Then restart BackupFP service
Restart-Service BackupFP -Force
```

### Restore-Only Installation

On dedicated restore machines, use a Restore-Only Backup Manager installation to reduce memory/CPU overhead.

---

## Troubleshooting

### Restore Did Not Register (Timeout)

**Symptom:** "Restore did not register within 60 seconds. Skipping session."

**Causes:**
- ClientTool crashed or hung
- SessionReport.xml is locked or inaccessible
- Backup Manager service is stopping/restarting

**Solution:**
- Check Backup Manager service status: `Get-Service BackupFP`
- Verify SessionReport.xml permissions in `C:\ProgramData\MXB\Backup Manager\`
- Review Backup Manager event log for errors
- Restart BackupFP service and re-run script

### Idle Wait Timeout

**Symptom:** "Timed out waiting for Idle state after 120 minutes. Skipping session."

**Causes:**
- A backup is running in the background
- Backup Manager is stuck (hung process)
- Previous restore is still running

**Solution:**
- Check active processes: `ClientTool.exe control.status.get`
- Stop/cancel any running backup: `ClientTool.exe control.backup.stop`
- Restart BackupFP service if needed
- Increase timeout by manually editing the script (line 285, change 120 to desired minutes)

### No Backup Sessions Found

**Symptom:** GridView opens empty; no sessions available to select.

**Causes:**
- No backup sessions exist for specified `-DataSource`
- Sessions are in a non-included state (Aborted, Failed, etc.)
- Sessions have been cleaned/archived

**Solution:**
- Verify backup history: `ClientTool.exe control.session.list -machine-readable`
- Check `-SessionType` filter (All vs. Weekly vs. Monthly)
- Check `-IncludedStates` parameter (add Aborted or Failed if needed)
- Restore from SessionReport.xml backup if sessions were cleaned

### Permission Denied on Restore Directory

**Symptom:** "Access to the path is denied" when creating restore directory.

**Causes:**
- Target volume is read-only
- Insufficient permissions on target directory
- Antivirus is blocking file creation

**Solution:**
- Verify target volume is writeable: `Test-Path -PathType Container "D:\"`
- Run PowerShell as Administrator
- Temporarily disable antivirus scanning on target directory
- Check NTFS permissions on parent directory

---

## Support

For issues with:

- **Script logic**: Verify ClientTool.exe works manually first
- **Backup Manager**: Contact N-able support with Backup Manager logs
- **Session recovery**: Review SessionReport.xml in `C:\ProgramData\MXB\Backup Manager\`

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| v12.2 | 2026-07-07 | Added performance tips; clarified session filtering; improved timeout documentation |
| v12.1 | 2026-06-XX | Initial release; dual-signal monitoring (SessionReport.xml + ClientTool status) |

---

## Author

**Eric Harless** — Head Backup Nerd, N-able

- Twitter: [@Backup_Nerd](https://twitter.com/Backup_Nerd)
- Email: eric.harless@n-able.com
- Repository: https://github.com/backupnerd

---

## License

Sample scripts are not supported under any N-able support program or service. The sample scripts are provided AS IS without warranty of any kind. N-able expressly disclaims all implied warranties including warranties of merchantability or of fitness for a particular purpose.
