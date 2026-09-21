# GetRemoteBackupSelections.v5.ps1

**Cove Data Protection - Remote Device Backup Selections Report**

Collects backup selections and schedules from Cove Backup Manager devices through the Remote Connection Gateway (RCG). The script maintains a persistent master CSV and produces a dated analyst-friendly CSV and optional Excel workbook.

---

## Requirements

- PowerShell 7 or later (`ForEach-Object -Parallel` is required)
- N-able Cove Data Protection Standalone edition
- API credentials in one of these forms:
   - DPAPI-protected XML at `C:\ProgramData\MXB\<COMPUTERNAME>_<USERNAME>_API_Credentials.xml`
   - `COVE_USERNAME` and `COVE_PASSWORD` environment variables
- Microsoft Excel for the default CSV/XLS export; use `-Export:$false` when Excel is unavailable

---

## How It Works

1. Authenticates to `backup.management`.
2. Resolves the target partner through `GetPartnerTree` and `GetPartnerInfoById`.
3. Retrieves Backup Manager devices and metadata with `EnumerateAccountStatistics`.
4. Connects to devices in parallel through RCG.
5. Collects datasource selections with `EnumerateBackupSelections`.
6. Collects regular and high-frequency schedules with `EnumerateBackupSchedule` and `GetHighFrequentBackupSchedule`.
7. Merges successful results into the persistent master CSV.
8. Preserves the last known selections when a device cannot be reached.
9. Automatically retries failed device connections once.
10. Exports the full master inventory in analyst-friendly format.

Datasource calls and schedule calls retry transient failures up to three times. The automatic device retry uses a reduced parallel throttle to avoid placing additional load on RCG relay servers.

---

## Files

| File | Description |
|------|-------------|
| `GetRemoteBackupSelections.v5.ps1` | Main script |
| `RemoteSelections_<Partner>_<ID>_MASTER.csv` | Persistent inventory and selection history |
| `Output\<date>\<timestamp>_RemoteSelections_<Partner>_<ID>.csv` | Dated analyst export |
| `Output\<date>\<timestamp>_RemoteSelections_<Partner>_<ID>.xlsx` | Optional Excel export |
| `Output\<date>\<timestamp>_RCG_Errors_<Partner>.csv` | Devices still failing RCG access |
| `RemoteSelections_<Partner>_<ID>_MASTER.csv.lock` | Lock file used during master updates |

When `-DebugCDP` is specified, the output folder also receives a raw schedule dump in Markdown format.

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-PartnerName` | From credentials or authenticated partner | Partner name or partial name used for partner-tree lookup |
| `-AllPartners` | `$true` | Skip GUI partner selection |
| `-AllDevices` | `$true` | Skip GUI device selection |
| `-DeviceCount` | `5000` | Maximum devices returned from API |
| `-Export` | `$true` | Generate CSV and Excel output |
| `-Launch` | `$true` | Open the XLS/CSV after export |
| `-Delimiter` | `,` | Delimiter used while writing the live CSV; the final analyst export is comma-delimited |
| `-ExportPath` | Script folder | Root path for master CSV and Output subfolder |
| `-ClearCredentials` | — | Delete stored credentials and re-prompt |
| `-DeviceThrottle` | `20` | Maximum parallel RCG device operations |
| `-ActiveWithinDays` | `90` | API heartbeat window; `0` includes all devices |
| `-MaxRcgAgeHours` | `4` | Do not query RCG for devices with an older heartbeat; `0` disables this filter |
| `-SkipRecentHours` | `72` | Reuse recent reachable master data without querying RCG; `0` disables this optimization |
| `-RetryAuthDelayMs` | `500` | Maximum randomized delay before immediate retry authentication |
| `-RetryOnly` | Off | Skip the normal pass and retry devices marked `No` or `Unknown` in the master |
| `-FilterAccountIDs` | None | Process only the specified numeric AccountIDs |
| `-ExcludeColumns` | `VM,SP,ORC,Exch` | Datasource columns to omit from export (see short codes below) |
| `-ExcludeMetaColumns` | `IP,OS,Mfr,Model,CPU,RAM,ProdID,ProfID` | Omit metadata columns from the analyst export |
| `-PathSeparator` | `Pipe` | Use ` | ` or a newline between paths in export columns |
| `-DebugCDP` | Off | Enable diagnostic output and raw schedule dumps |

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

### ExcludeMetaColumns Short Codes

| Code | Export column |
|------|---------------|
| `IP` | `IPAddress` |
| `OS` | `OS` |
| `Phys` | `Physicality` |
| `Mfr` | `Manufacturer` |
| `Model` | `Model` |
| `CPU` | `CPUCores` |
| `RAM` | `RAMBytes` |
| `ProdID` | `ProductID` |
| `Prod` | `Product` |
| `ProfID` | `ProfileID` |
| `Prof` | `Profile` |
| `Excl` | Omit all `Exc-` columns |

---

## Master CSV

The master CSV is the persistent source of truth across runs.

- Reachable devices receive refreshed metadata, selections, and schedules.
- Unreachable devices receive refreshed API metadata while retaining their last successful selections and schedules.
- Devices absent from a complete current inventory are retained and marked `NotInCurrentInventory`.
- Recently validated devices can be skipped while retaining their existing data.
- Devices with stale API heartbeats can be left out of RCG processing.
- A lock file prevents concurrent instances from writing the same master.
- Embedded CSV corruption is checked during save. Schedule-related corruption is cleared for later refresh; unrecoverable metadata rows are dropped so they can be re-added on a later full run.

Datasource calls and schedule calls retry transient failures up to three times. The automatic device retry uses a reduced parallel throttle. Devices that remain unreachable are written to the RCG error log with device identity, partner, OS, profile, storage status, timestamps, RCG host, installation information, and failure reason.

The master contains metadata plus per-datasource fields for schedules, high-frequency schedules, selections, signatures, validation timestamps, and change indicators.

---

## Analyst Export

The dated analyst export is rebuilt from the complete master on every run. It contains device metadata such as:

`Anomalies`, `PartnerID`, `PartnerName`, `AccountID`, `DeviceName`, `ComputerName`, `IPAddress`, `OS`, `Physicality`, `Manufacturer`, `Model`, `CPUCores`, `RAMBytes`, `ProductID`, `Product`, `ProfileID`, `Profile`, `ClientVersion`, `DataSources`, `TimeZone`, `CreationDate`, `LastSuccess`, `TimeStamp`, and `Reachable`.

Datasource columns include:

- `<DS> Sched` — effective schedule; the high-frequency schedule is preferred when one exists
- `<DS> Last` — last validation timestamp
- `<DS> Inc+` — inclusion paths
- `<DS> Exc-` — exclusion paths, unless `Excl` is selected

The master retains regular and high-frequency schedule values separately, even though the analyst export presents the effective schedule in one column.

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

The `Anomalies` column contains semicolon-separated flags.

| Flag | Meaning |
|------|---------|
| `UNREACHABLE` | Device could not be reached via RCG |
| `ORPHANED` | Device is no longer in the current partner inventory |
| `PROFILE` | One or more selections were pushed by an account profile (`[P]`) |
| `EXCLUSIONS` | One or more datasource exclusion paths are configured |
| `SPECIFIC_FS` | FileSystem is not backed up in full — specific paths are selected |
| `ORPHANED_DS` | A datasource is enabled in device settings but has no selections configured |

Additional diagnostics shown during processing include orphaned datasources, missing exclusions, profile-based selection counts, datasource prevalence, configuration standardization, and unexpected selection flags.

---

## Usage Examples

```powershell
# Full run — all devices, export to default path
.\GetRemoteBackupSelections.v5.ps1

# Process no more than 500 devices
.\GetRemoteBackupSelections.v5.ps1 -DeviceCount 500

# Retry devices previously marked unreachable
.\GetRemoteBackupSelections.v5.ps1 -RetryOnly

# Process selected devices
.\GetRemoteBackupSelections.v5.ps1 -FilterAccountIDs 1234567,9876543

# Exclude datasource and metadata columns
.\GetRemoteBackupSelections.v5.ps1 -ExcludeColumns Exch,SP,ORC,VM -ExcludeMetaColumns IP,OS,RAM

# Export elsewhere and do not auto-launch
.\GetRemoteBackupSelections.v5.ps1 -ExportPath "D:\Reports" -Launch:$false

# Enable diagnostics and include devices regardless of heartbeat age
.\GetRemoteBackupSelections.v5.ps1 -DebugCDP -ActiveWithinDays 0 -MaxRcgAgeHours 0
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

## Retry notes

Every normal run performs one automatic immediate retry for devices that fail the first RCG attempt. `-RetryOnly` is available for a later focused retry run.

`-RetryFailed` is not an active v5 parameter; the former optional third retry pass was removed. The automatic retry is always performed.

---

## Legal

Sample scripts are not supported under any N-able support program or service. They are provided AS IS without warranty of any kind. N-able expressly disclaims all implied warranties. Sample scripts may contain non-public API calls that are subject to change without notification.

## Repository

The latest version of this script is maintained in the [Backup-Scripts repository](https://github.com/BackupNerd/Backup-Scripts).
