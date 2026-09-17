# GetRemoteBackupSelections.v3.ps1

**Cove Data Protection - Remote Device Backup Selections Report**

Enumerates backup selections and schedules for all Cove Backup Manager devices in a partner account by connecting to each device via the Remote Connection Gateway (RCG). Maintains a persistent master CSV across runs and produces an analyst-ready export with anomaly flagging.

---

## Requirements

- PowerShell 7+ (uses `ForEach-Object -Parallel`)
- N-able Cove Data Protection — Standalone edition
- Stored API credentials at `C:\ProgramData\MXB\<COMPUTERNAME>_<USERNAME>_API_Credentials.xml` (DPAPI-encrypted), or environment variables `COVE_USERNAME` / `COVE_PASSWORD`
- Excel (optional) — for XLS export via `Export-Csv` + COM automation

---

## How It Works

1. Authenticates to `backup.management` using the machine/user credential XML or prompted credentials
2. Calls `EnumerateAccountStatistics` to retrieve all Backup Manager devices and their metadata (OS, hardware, profile, product, client version, last success, datasource flags)
3. Loads an existing master CSV if present; initialises a new one on first run
4. Connects to each device in parallel via RCG and calls:
   - `EnumerateBackupSelections` — per-datasource inclusion/exclusion paths
   - `EnumerateBackupSchedule` + `GetHighFrequentBackupSchedule` — schedule details
5. Merges live data into the master; unreachable devices retain their last known selections
6. Saves the updated master CSV, then exports the full master in analyst-friendly format to a dated `Output\<date>` subfolder

---

## Files

| File | Description |
|------|-------------|
| `GetRemoteBackupSelections.v3.ps1` | Main script |
| `RemoteSelections_<Partner>_<ID>_MASTER.csv` | Persistent master — all devices, all runs |
| `Output\<date>\<date>_RemoteSelections_<Partner>.csv` | Dated analyst export (CSV) |
| `Output\<date>\<date>_RemoteSelections_<Partner>.xlsx` | Dated analyst export (XLS) |
| `<ExportPath>\RemoteSelections_<Partner>_<ID>_MASTER.csv.lock` | Temporary lock file while the script is running |

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-PartnerName` | *(from stored credentials)* | Partner name passed to the legacy exact, case-sensitive `GetPartnerInfo` lookup |
| `-AllPartners` | `$true` | Skip GUI partner selection |
| `-AllDevices` | `$true` | Skip GUI device selection |
| `-DeviceCount` | `5000` | Maximum devices returned from API |
| `-Export` | `$true` | Generate CSV / XLS output files |
| `-Launch` | `$true` | Open the XLS/CSV after export |
| `-Delimiter` | `,` | CSV field delimiter |
| `-ExportPath` | Script folder | Root path for master CSV and Output subfolder |
| `-ClearCredentials` | — | Delete stored credentials and re-prompt |
| `-DeviceThrottle` | `20` | Max parallel RCG threads |
| `-RetryFailed` | — | Run an additional end-of-script retry for devices still unreachable after the automatic immediate retry |
| `-RetryOnly` | — | Skip main pass; only retry devices marked unreachable in master |
| `-ActiveWithinDays` | `7` | Only process devices with a heartbeat within N days (0 = all) |
| `-FilterAccountIDs` | — | Process only the specified AccountIDs |
| `-ExcludeColumns` | `VM,SP,ORC,Exch` | Datasource columns to omit from export (see short codes below) |
| `-ExcludeMetaColumns` | `IP,OS,Mfr,Model,CPU,RAM,ProdID,ProfID` | Metadata columns to omit from the analyst export |
| `-PathSeparator` | `Pipe` | Use `Pipe` (` | `) or `Newline` between paths in export columns |
| `-DebugCDP` | — | Enable verbose debug output and schedule dump files |

### ExcludeColumns Short Codes

| Code | Datasource |
|------|-----------|
| `FS` | FileSystem |
| `SS` | System State (all variants) |
| `HypV` | Hyper-V |
| `SQL` | MSSQL |
| `MySQ` | MySQL |
| `Net` | Network Shares |
| `VM` | VMware |
| `SP` | SharePoint |
| `ORC` | Oracle |
| `Exch` | Exchange |

---

## Master CSV

The master CSV is the persistent source of truth. It accumulates device records across runs:

- **Reachable devices** — selections, schedules, and all metadata updated every run
- **Unreachable devices** — metadata refreshed from `EnumerateAccountStatistics`; selections and schedules preserved from the last successful reach
- **Orphaned devices** — devices no longer in the current inventory are flagged `NotInCurrentInventory` but retained in the master
- **Immediate retry** — unreachable devices are automatically retried once after the first pass, regardless of `-RetryFailed`
- **Optional final retry** — `-RetryFailed` retries any devices still unreachable after the immediate retry

On save, the master is checked for CSV corruption. Schedule column corruption is self-healed (cleared for refresh on next run). Metadata column corruption causes the row to be dropped and logged for re-addition on next full run.

A `.lock` file prevents concurrent script instances from corrupting the master CSV.

---

## Analyst Export

The dated export is built from the full master on every run. Columns include:

**Device metadata:** `Anomalies`, `PartnerID`, `PartnerName`, `AccountID`, `DeviceName`, `ComputerName`, `IPAddress`, `OS`, `Physicality`, `Manufacturer`, `Model`, `CPUCores`, `RAMBytes`, `ProductID`, `Product`, `ProfileID`, `Profile`, `ClientVersion`, `DataSources`, `CreationDate`, `LastSuccess`, `TimeStamp`, `Reachable`

**Per datasource (repeated for each active datasource):** `<DS> Sched`, `<DS> HFSched`, `<DS> Last`, `<DS> Inc+`, `<DS> Exc-`

### Selection Path Markers

The `Inc+` and `Exc-` path columns prefix selections with a marker showing their origin:

| Marker | Meaning |
|--------|---------|
| `[P]` | Profile-created selection — controlled by an assigned Cove backup profile |
| `[L]` | Local selection — manually configured on the device |
| `[X]` | Root-style or orphaned selection — typically a leftover profile/root selection without an active profile association |

These markers make it easier to distinguish policy-controlled paths from local configuration and investigate potentially orphaned selections.

---

## Anomaly Flags

Anomalies appear as a semicolon-separated list in the `Anomalies` column.

| Flag | Meaning |
|------|---------|
| `UNREACHABLE` | Device could not be reached via RCG |
| `ORPHANED` | Device is no longer in the current partner inventory |
| `PROFILE` | One or more selections were pushed by an account profile (`[P]`) |
| `EXCLUSIONS` | One or more datasource exclusion paths are configured |
| `SPECIFIC_FS` | FileSystem is not backed up in full — specific paths are selected |
| `ORPHANED_DS` | A datasource is enabled in device settings but has no selections configured |
| `PROFILE_BASED` | On-screen/pivoted row flag for profile-created selections |
| `NO_EXCLUSIONS` | On-screen/pivoted row flag when no exclusion path was returned |
| `RETRY_RECOVERY` | On-screen/pivoted row flag for a device recovered during the optional retry pass |

---

## Usage Examples

```powershell
# Full run — all devices, export to default path
.\GetRemoteBackupSelections.v3.ps1

# Limit to 500 devices, retry unreachable ones
.\GetRemoteBackupSelections.v3.ps1 -DeviceCount 500 -RetryFailed

# Only retry previously unreachable devices (no main pass)
.\GetRemoteBackupSelections.v3.ps1 -RetryOnly

# Exclude Exchange, SharePoint, Oracle, and VMware columns from export
.\GetRemoteBackupSelections.v3.ps1 -ExcludeColumns Exch,SP,ORC,VM

# Process specific devices only
.\GetRemoteBackupSelections.v3.ps1 -FilterAccountIDs 1234567,9876543

# Export to a different path, don't auto-launch
.\GetRemoteBackupSelections.v3.ps1 -ExportPath "D:\Reports" -Launch:$false
```

---

## Credentials

Credentials are stored DPAPI-encrypted via `Export-Clixml` at:

```
C:\ProgramData\MXB\<COMPUTERNAME>_<USERNAME>_API_Credentials.xml
```

The XML may contain the current custom format (`PartnerName`, `Username`, encrypted `Password`) or the legacy `PSCredential` format (`UserName`, encrypted `Password`).

Fallback: `COVE_USERNAME` and `COVE_PASSWORD` environment variables.

Use `-ClearCredentials` to delete stored credentials and re-prompt on next run.

### Partner lookup note

This v3 script still calls the deprecated `GetPartnerInfo(name=...)` method. The lookup is exact and case-sensitive, and duplicate partner names may resolve incorrectly. The modern `GetPartnerTree` + `GetPartnerInfoById` migration has been applied to other scripts but not yet to this v3 script.

---

## Legal

Sample scripts are not supported under any N-able support program or service. Provided AS IS without warranty of any kind. N-able expressly disclaims all implied warranties. Sample scripts may contain non-public API calls which are subject to change without notification.
