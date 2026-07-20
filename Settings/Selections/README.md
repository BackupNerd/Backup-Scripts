# GetLocalBackupSelections.v2.ps1

**Cove Data Protection - Local Device Backup Selections Report**

Enumerates backup selections from the local Cove Backup Manager on the machine where the script is run. This is the local-device companion to `GetRemoteBackupSelections.v3.ps1` and is useful when you want to inspect configured datasource inclusion and exclusion paths directly on a single endpoint without using RCG.

---

## Requirements

- PowerShell 7+
- N-able Cove Backup Manager installed on the local device
- Sufficient permissions to query the local Backup Manager JSON-RPC endpoint

---

## How It Works

1. Connects to the local Backup Manager JSON-RPC interface on the device where the script is executed
2. Authenticates using an InAgent authentication token from the local Backup Manager storage path
3. Calls `EnumerateBackupSelections` for the configured datasources
4. Flattens the returned rows into an analyst-friendly table and prints a per-datasource summary
5. Highlights inclusion paths, exclusion paths, and profile-driven items where applicable

---

## Typical Use Cases

- Validate the exact selections on one device before rolling out profile changes
- Compare local selections with centrally reported remote selections
- Troubleshoot why a datasource is or is not protected
- Review inclusion, exclusion, and profile-applied backup selections on a specific endpoint

---

# GetRemoteBackupSelections.v3.ps1

**Cove Data Protection - Remote Device Backup Selections Report**

Enumerates backup selections and schedules for all Cove Backup Manager devices in a partner account by connecting to each device via the Remote Connection Gateway (RCG). Maintains a persistent master CSV of device selection data across runs.

---

## Requirements

- PowerShell 7+ (uses `ForEach-Object -Parallel`)
- N-able Cove Data Protection — Standalone edition
- Stored API credentials at `C:\ProgramData\MXB\mcpcred.xml` (DPAPI-encrypted), or environment variables `COVE_USERNAME` / `COVE_PASSWORD`
- Excel (optional) — for XLS export via `Export-Csv` + COM automation

---

## How It Works

1. Authenticates to `backup.management` using stored or prompted credentials
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
| `Archive\` | Timestamped backups of script and master CSV |

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-PartnerName` | *(from stored credentials)* | Root partner name to authenticate against |
| `-AllPartners` | `$true` | Skip GUI partner selection |
| `-AllDevices` | `$true` | Skip device selection GUI |
| `-DeviceCount` | `5000` | Maximum devices returned from API |
| `-Export` | `$true` | Generate CSV / XLS output files |
| `-Launch` | `$true` | Open the XLS/CSV after export |
| `-Delimiter` | `,` | CSV field delimiter |
| `-ExportPath` | Script folder | Root path for master CSV and Output subfolder |
| `-ClearCredentials` | — | Delete stored credentials and re-prompt |
| `-DeviceThrottle` | `40` | Max parallel RCG threads |
| `-RetryFailed` | — | Retry unreachable devices once after the first pass |
| `-RetryOnly` | — | Skip main pass; only retry devices marked unreachable in master |
| `-ActiveWithinDays` | `7` | Only process devices with a heartbeat within N days (0 = all) |
| `-FilterAccountIDs` | — | Process only the specified AccountIDs |
| `-ExcludeColumns` | `VM,SP,ORC,Exch` | Datasource columns to omit from export (see short codes below) |
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

On save, the master is checked for CSV corruption. Schedule column corruption is self-healed (cleared for refresh on next run). Metadata column corruption causes the row to be dropped and logged for review.

A `.lock` file prevents concurrent script instances from corrupting the master CSV.

---

## Analyst Export

The dated export is built from the full master on every run. Columns include:

**Device metadata:** `Anomalies`, `PartnerID`, `PartnerName`, `AccountID`, `DeviceName`, `ComputerName`, `IPAddress`, `OS`, `Physicality`, `Manufacturer`, `Model`, `CPUCores`, `RAMBytes`, `ProductID`, `Product`, `ProfileID`, `Profile`, `ClientVersion`, `CreationDate`, `TimeStamp`, `LastSuccess`

**Per datasource (repeated for each active datasource):** `<DS> Sched`, `<DS> HFSched`, `<DS> Last`, `<DS> Inc+`, `<DS> Exc-`

---

## Anomaly Flags

Anomalies appear as a semicolon-separated list in the `Anomalies` column.

| Flag | Meaning |
|------|---------|
| `UNREACHABLE` | Device could not be reached via RCG |
| `ORPHANED` | Device is no longer in the current partner inventory |
| `PROFILE` | One or more selections were pushed by an account profile (`[p]`) |
| `EXCLUSIONS` | One or more datasource exclusion paths are configured |
| `SPECIFIC_FS` | FileSystem is not backed up in full — specific paths are selected |
| `ORPHANED_DS` | A datasource is enabled in device settings but has no selections configured |

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
C:\ProgramData\MXB\mcpcred.xml           # default
C:\ProgramData\MXB\mcpcred-<login>.xml   # named profile
```

Fallback: `COVE_USERNAME` and `COVE_PASSWORD` environment variables.

Use `-ClearCredentials` to delete stored credentials and re-prompt on next run.

---

## Legal

Sample scripts are not supported under any N-able support program or service. Provided AS IS without warranty of any kind. N-able expressly disclaims all implied warranties. Sample scripts may contain non-public API calls which are subject to change without notice.
