# CoveDP-2-CWM Workflow Reference
**Script:** `CoveDP-2-CWM.v26.06.18.MetroTech.Base.ps1`
**Generated:** 2026-06-18

---

## Validation: Interactive vs Unattended Modes

**Finding: Each Interactive sub-mode is functionally identical to its unattended counterpart.** After the workflow selection point, the internal state flags are set identically by both paths:

| Interactive choice | Unattended RunMode | `GetCDPUsage` | `LoadCDPUsage` | `ShouldProcessCWMAdditions` |
|---|---|---|---|---|
| `P` (Pull) | `PullOnly` | `$true` | `$false` | `$false` |
| `U` (Upload) | `UploadOnly` | `$false` | `$true` | `$true` |
| `C` (Combined) | `PullAndUpload` | `$true` | `$false` | `$true` |

**The only behavioral difference** between Interactive and unattended:
- **Interactive:** If neither `GetCDPUsage` nor `LoadCDPUsage` is pre-set at launch, the user is prompted to choose P / U / C / X.
- **Unattended:** The RunMode switch sets all flags directly. If an invalid combo is detected (e.g., `ShouldProcessCWMAdditions=$true` but `ConnectCWM=$false`), the script exits immediately with a warning — no prompts.

All three workflow sections below treat each Interactive/Unattended pair as a single mode.

---

## Pull Path vs Upload Path — Data Source Validation

These are fully separate, mutually exclusive code paths. They share only the CWM interaction layer (`LookupCWMCompany` → `UpdateCWMQty`).

| | Pull Path (`GetCDPUsage=$true`) | Upload Path (`LoadCDPUsage=$true`) |
|---|---|---|
| **Data source** | Live Cove API (real-time) | Previously exported `*CDPUsageFile*.csv` |
| **Cove API calls** | Yes — 3 endpoints | None |
| **Period determined by** | `-Period` / `-Last` parameter | Regex extracted from CSV filename |
| **IgnoreDeletedDevices applies** | Yes — strips deleted devices before quantities | No — data already filtered when CSV was generated |
| **TrialState applies** | Yes — gates which partners are included | No — CSV reflects what was exported at pull time |
| **DeviceCount / DeviceFilter apply** | Yes | No |
| **FairUseGB params apply** | Yes — affects OverageGB → CoveGBSelected | No — CoveGBSelected comes from CSV |
| **In-memory data** | `$Script:CDPUsageCollection` (built live) | `$SelectedUsage` (imported from CSV, GridView-filtered) |
| **GridView selection** | No — all production customers processed | Yes (Interactive) — user selects which rows to push |
| **Auto-select in unattended** | N/A | Newest `*CDPUsageFile*.csv` under `$ExportPath` |
| **Output files generated** | All output files (CSV, XLS, CDPUsage, Audit, Exceptions) | Only Audit + Exceptions appended; no new CSV/XLS generated |

**Risk note:** The Upload path trusts the CSV file as the source of truth. If that file was generated with different parameters (`TrialState`, `IgnoreDeletedDevices`, `FairUseGB`, etc.) than the current run, those differences are silently carried through to CWM. There is no validation of the CSV's generation parameters.

---

## Files Created Every Run (Always)

Regardless of mode, the following are created immediately after authentication and partner lookup:

| File | Pattern | Mode | Risk |
|---|---|---|---|
| Export folder | `{ExportPath}\{date}_{partner}_{id}\` | All | Created with `-Force`; safe (timestamped) |
| Audit Log | `*_AuditLog.csv` | All | Header row written at startup; appended per CWM addition |
| Exceptions Log | `*_ExceptionsLog.txt` | All | Created lazily on first exception only |

---

## Symbols Used

| Symbol | Meaning |
|---|---|
| 📥 READ | Data read from external system (Cove API, CWM API, local file) |
| 📤 WRITE | Data written to external system (CWM API) — **modifies live data** |
| 💾 FILE | Data written to local file |
| ⚠️ RISK | Step that can overwrite or lose data if parameters are wrong |
| 🔒 GATE | Parameter that controls whether this step executes |

---

## Mode 1 — Pull Only
**Triggers:** Interactive `P`, or `RunMode=PullOnly`
**Internal state:** `GetCDPUsage=$true`, `LoadCDPUsage=$false`, `ShouldProcessCWMAdditions=$false`
**Purpose:** Pull live Cove usage data, export to files, optionally read CWM additions for comparison. **Never writes to CWM additions.**

### Workflow Steps

```
[START]
   │
   ├─ 📥 READ  Cove API Login (backup.management)
   │           Method: Login → stores visa token (10-min expiry, auto-renewed)
   │           Source: Encrypted cred file at C:\ProgramData\MXB\
   │   🔒 GATE: -ClearCDPCredentials → deletes cred file, re-prompts before authenticating
   │
   ├─ 📥 READ  Cove API: GetPartnerInfo / GetPartnerInfoById
   │           Resolves PartnerId and Level for the target partner
   │   🔒 GATE: -PartnerNameOverride (default="Metrotech") — overrides stored cred partner
   │
   ├─ 💾 FILE  Create timestamped export folder under $ExportPath
   ├─ 💾 FILE  Create AuditLog.csv (header row only at this point)
   │
   ├─ 📥 READ  Cove API: EnumerateAccountStatistics
   │           Fetches all backup device statistics for the partner
   │   🔒 GATE: -DeviceCount (default=5000) — max rows returned; too low may miss devices ⚠️
   │   🔒 GATE: -DeviceFilter (default="AT==1 OR AT==2") — filters device types
   │
   ├─ 📥 READ  Cove API: high-watermark-usage-report (statistics-reporting endpoint)
   │           Downloads TempMVReport.xlsx for the billing period
   │   🔒 GATE: -Period / -Last — determines which month's HWM report is fetched
   │   🔒 GATE: -PartnerNameOverride → PartnerId used in report URL
   │   💾 FILE  Writes TempMVReport.xlsx (temporary; overwritten each run)
   │
   ├─ 📥 READ  Cove API: EnumerateAncestorPartners (per unique customerId in HWM report)
   │           Builds partner hierarchy (Disti → SubDisti → Reseller → SO → EndCustomer → Site)
   │           Cached: re-called only when customerId changes
   │
   ├─ 📥 READ  [Conditional] Cove API: GetPartnerInfoHistory
   │           Called only for deleted/anonymized partners where EnumerateAncestorPartners errors
   │
   ├─ 📥 READ  [Conditional] Cove API: EnumerateAccountHistoryStatistics
   │           Called once per device that has a DeviceDeletionDate to recover Physicality/Product
   │   🔒 GATE: -IgnoreDeletedDevices (default=$true)
   │           → if $true: deleted devices REMOVED from $MVPlus BEFORE quantity calculation ⚠️
   │             (changes which devices count toward Phys/Virt/Workstation/M365 totals)
   │           → if $false: deleted devices INCLUDED in quantities (may inflate CWM additions) ⚠️
   │
   ├─ 💾 FILE  Export *_Statistics.csv (ALL customers, all CustomerStates, no Recycle Bin)
   │
   ├─ 💾 FILE  Export *_Statistics.xlsx — Trial tab (InTrial only)
   │   🔒 GATE: -TrialState has no effect on Trial tab; it always contains InTrial rows
   │
   ├─ 💾 FILE  Export *_Statistics.xlsx — Production tab + Pivot (InProduction, or +InTrial)
   │   🔒 GATE: -TrialState (default=$false)
   │           → $true: InTrial partners included in Production tab and CDPUsage totals ⚠️
   │           → $false: InProduction only
   │   🔒 GATE: -PhysServerFUGB / -VirtServerFUGB / -WorkstationFUGB (default=0)
   │           → affects OverageGB = SelectedSizeGB - FairUseGB
   │           → 0 means full SelectedSizeGB passes through as OverageGB
   │           → affects CoveGBSelected quantity exported to CDPUsage CSV ⚠️
   │
   ├─ 💾 FILE  Export *_CDPUsageFile_yyyy-MM_.csv (per-customer summary, production only)
   │           Columns: SubDisti, Reseller, SO, EndCustomer, PartnerID, PartnerName, LegalName,
   │                    PartnerRef, [CombinedServerQty,] PhysicalServerQty, VirtualServerQty, WorkstationQty,
   │                    DocumentsQty, RecoveryTestingQty, O365UsersQty, CoveGBSelected
   │           ⚠️ CoveGBSelected is cast to [int] — fractional GB is truncated
   │   🔒 GATE: -CombineServers (default=$true)
   │             → $true: CombinedServerQty IS included as the first quantity column (before PhysicalServerQty)
   │             → $false: CombinedServerQty column is omitted entirely
   │
   ├─ [IF -ConnectCWM=$true]  ← default is $true
   │   │
   │   ├─ [ACTION] InstallCWMPSModule — installs/updates ConnectWiseManageAPI PS module
   │   │
   │   ├─ [IF -ClearCWMCredentials] → deletes CWM cred file, triggers re-prompt
   │   │
   │   ├─ [ACTION] AuthenticateCWM
   │   │   📥 READ  CWM cred file ($env:computername CWM_API_Credentials.Secure.metrotech.xml)
   │   │            Decrypts keys in memory → calls Connect-CWM → plaintext cleared immediately
   │   │
   │   ├─ [ACTION] LookupCWMProducts
   │   │   📥 READ  CWM API: Get-CWMProductCatalog (one call per product in CDP2CWMProductMapping)
   │   │            Checks all 8 mapped product IDs exist in CWM catalog
   │   │            ⚠️ Warnings only — missing products do not halt execution
   │   │   💾 FILE  ExceptionsLog appended if any product missing
   │   │
   │   ├─ [IF -DebugCWM=$true]  ← default is $true
   │   │   📥 READ  CWM API: Get-CWMAgreement (all active agreements)
   │   │   💾 FILE  *_CWMAgreements.csv written
   │   │   🔒 GATE: -CWMAgreementBehavior ("Type" or "Name") — affects filter
   │   │   🔒 GATE: -CWMAgreementTypes / -CWMAgreementName / -CWMAgreementSearchType
   │   │
   │   └─ FOR EACH customer in $Script:CDPUsageCollection
   │       │
   │       ├─ [ACTION] LookupCWMCompany
   │       │   📥 READ  CWM API: Get-CWMcompany (by ID, then Name, then LegalName, then Ref)
   │       │            Up to 4 CWM company lookups per customer
   │       │   💾 FILE  ExceptionsLog if no match found
   │       │
   │       ├─ [IF company in $CWMExcludedCompanies] → SKIP, log to ExceptionsLog
   │       │
   │       ├─ 📥 READ  CWM API: Get-CWMAgreement (for matched company)
   │       │   🔒 GATE: -CWMAgreementBehavior / -CWMAgreementTypes / -CWMAgreementName
   │       │   🔒 GATE: -CWMAgreementSearchType (Equals/Contains/StartsWith/EndsWith) ⚠️
   │       │            "Contains" is broadest; could match unintended agreements
   │       │
   │       ├─ [IF 0 or >1 agreements match] → SKIP company, log to ExceptionsLog
   │       │
   │       ├─ [IF exactly 1 agreement matches]
   │       │   🔒 GATE: ($ShouldProcessCWMAdditions -or $NoCWMUpdate)
   │       │            = ($false -or $NoCWMUpdate)
   │       │            → if $NoCWMUpdate=$true (default): calls UpdateCWMQty ✓
   │       │            → if $NoCWMUpdate=$false: SKIPS UpdateCWMQty entirely
   │       │              ⚠️ No CWM reads occur in this case (pull-only with NoCWMUpdate=$false
   │       │                 produces files but no CWM addition snapshot)
   │       │
   │       └─ [ACTION] UpdateCWMQty  [only if NoCWMUpdate=$true]
   │           │
   │           ├─ $ReadOnlyCWMUsageMode = $NoCWMUpdate = $true  → READ-ONLY PATH
   │           │
   │           ├─ 📥 READ  CWM API: Get-CWMAgreementAddition (all active additions for agreement)
   │           │
   │           ├─ Compute proposed qty for each Cove-mapped addition using in-memory CDPUsage
   │           │   🔒 GATE: -CombineServers (default=$true)
   │           │            → $true: Phys/Virt additions proposed as 0; CombinedServer proposed
   │           │            → $false: Phys/Virt proposed; CombinedServer skipped
   │           │
   │           ├─ Display Current vs Proposed table (console)
   │           │   Colors: Green=increase, Red=decrease, Yellow=no change, Gray=not mapped
   │           │
   │           ├─ 💾 FILE  AuditLog.csv appended: current qty AND proposed qty per addition
   │           │           (non-Cove additions get blank proposed value)
   │           │
   │           └─ return  [NO CWM writes — $ReadOnlyCWMUsageMode=$true]
   │
   └─ [END]
      💾 FILE  Console summary: paths of all output files
```

### Mode 1 Risk Summary

| Risk | Condition | Severity |
|---|---|---|
| Devices missed in quantities | `-DeviceCount` too low | Medium |
| Deleted devices inflate counts | `-IgnoreDeletedDevices=$false` | Medium |
| Trial partners billed | `-TrialState=$true` | Medium |
| GB overage miscalculated | `-PhysServerFUGB`/`VirtServerFUGB`/`WorkstationFUGB` wrong | Medium |
| No CWM addition snapshot | `-NoCWMUpdate=$false` (unusual for pull mode) | Low |
| Wrong partner scope | `-PartnerNameOverride` incorrect | High |

---

## Mode 2 — Upload Only
**Triggers:** Interactive `U`, or `RunMode=UploadOnly`
**Internal state:** `GetCDPUsage=$false`, `LoadCDPUsage=$true`, `ShouldProcessCWMAdditions=$true`
**Purpose:** Push previously exported CDPUsage CSV quantities to CWM additions.

**Note on `$ShouldProcessCWMAdditions=$false` in Mode 2:** This cannot occur through normal code paths. `$ShouldProcessCWMAdditions` is always `$true` when entering Mode 2. The `$NoCWMUpdate` flag is the sole gate between read-only snapshot and live write.

**`ConnectCWM` requirement:** If `-ConnectCWM=$false` is passed with Mode 2, the script exits at startup validation before any CWM work begins. Mode 2 requires `$ConnectCWM=$true`.

### Mode 2 has two sub-paths gated by `-NoCWMUpdate`

**Sub-path A — `-NoCWMUpdate=$true` (default): Read-only snapshot**
Reads CWM additions, shows Current vs Proposed table using CSV quantities. No CWM writes.

**Sub-path B — `-NoCWMUpdate=$false`: Live write**
⚠️ Reads CWM additions, then calls `Update-CWMAgreementAddition` to overwrite quantities.

### Workflow Steps

```
[START]
   │
   ├─ [VALIDATION] $ShouldProcessCWMAdditions -and -not $ConnectCWM → EXIT if true
   │
   ├─ 📥 READ  Cove API Login (same as Mode 1 — required for partner lookup)
   │   🔒 GATE: -ClearCDPCredentials / -PartnerNameOverride
   │
   ├─ 📥 READ  Cove API: GetPartnerInfo (establishes PartnerId/PartnerName for file naming)
   │           NOTE: No device or HWM data is pulled. This call only resolves the partner.
   │
   ├─ 💾 FILE  Create export folder, AuditLog.csv (header only)
   │
   ├─ [NO COVE DATA PULL — GetCDPUsage=$false; entire Cove device/HWM block is skipped]
   │   (No EnumerateAccountStatistics, no HWM report, no CDPUsage CSV generated)
   │
   ├─ [IF -ConnectCWM=$true]  ← must be $true for Mode 2
   │   │
   │   ├─ [ACTION] InstallCWMPSModule
   │   ├─ [IF -ClearCWMCredentials] → deletes CWM cred file
   │   ├─ [ACTION] AuthenticateCWM
   │   │   📥 READ  CWM cred file → Connect-CWM
   │   │
   │   ├─ [ACTION] LookupCWMProducts
   │   │   📥 READ  CWM API: Get-CWMProductCatalog (8 product checks)
   │   │   💾 FILE  ExceptionsLog if product missing
   │   │
   │   ├─ [IF -DebugCWM=$true]
   │   │   📥 READ  CWM API: Get-CWMAgreement (all active agreements)
   │   │   💾 FILE  *_CWMAgreements.csv
   │   │
   │   ├─ [LoadCDPUsage block]
   │   │   │
   │   │   ├─ [IF Interactive] GUI file picker → user selects *CDPUsageFile*.csv
   │   │   │   📥 READ  Local file system
   │   │   │
   │   │   ├─ [IF Unattended] Auto-selects newest *CDPUsageFile*.csv under $ExportPath
   │   │   │   📥 READ  Local file system
   │   │   │   ⚠️ If no file found → EXIT
   │   │   │
   │   │   ├─ 📥 READ  Import-Csv from selected file
   │   │   │           Validates 15 expected column headers; exits if any missing
   │   │   │
   │   │   ├─ [IF -CombineServers=$true]
   │   │   │   Adds CombinedServerQty = PhysicalServerQty + VirtualServerQty to each row
   │   │   │
   │   │   ├─ [IF Interactive] Out-GridView — user selects which customers to process
   │   │   │   ⚠️ Unattended mode: ALL rows in CSV are processed (no selection prompt)
   │   │   │   ⚠️ No way to exclude individual customers in unattended; use $CWMExcludedCompanies
   │   │   │
   │   │   └─ FOR EACH selected customer ($CDPusage from CSV row)
   │   │       │
   │   │       ├─ [ACTION] LookupCWMCompany
   │   │       │   📥 READ  CWM API: Get-CWMcompany (by ID → Name → LegalName → Ref)
   │   │       │   💾 FILE  ExceptionsLog if no match
   │   │       │
   │   │       ├─ [IF company in $CWMExcludedCompanies] → SKIP, log to ExceptionsLog
   │   │       │
   │   │       ├─ 📥 READ  CWM API: Get-CWMAgreement (for matched company)
   │   │       │   🔒 GATE: -CWMAgreementBehavior / -CWMAgreementTypes / -CWMAgreementName
   │   │       │   🔒 GATE: -CWMAgreementSearchType ⚠️
   │   │       │
   │   │       ├─ [IF 0 or >1 agreements match] → SKIP, log to ExceptionsLog
   │   │       │
   │   │       └─ [IF exactly 1 agreement matches]
   │   │           🔒 GATE: ($ShouldProcessCWMAdditions -or $NoCWMUpdate) = ($true -or any) = $true
   │   │           → UpdateCWMQty is ALWAYS called in Mode 2
   │   │
   │   └─ [ACTION] UpdateCWMQty
   │       │
   │       ├─ $ReadOnlyCWMUsageMode = $NoCWMUpdate
   │       │
   │       ├─ 📥 READ  CWM API: Get-CWMAgreementAddition (all active additions)
   │       │
   │       ├─ [IF $ReadOnlyCWMUsageMode=$true — Sub-path A]
   │       │   ├─ Compute proposed qty per addition from CSV-sourced CDPusage values
   │       │   ├─ Display Current vs Proposed table (console)
   │       │   ├─ 💾 FILE  AuditLog.csv: current qty + proposed qty per addition
   │       │   └─ return  [NO CWM WRITES]
   │       │
   │       └─ [IF $ReadOnlyCWMUsageMode=$false — Sub-path B]  ⚠️ LIVE WRITE PATH
   │           │
   │           ├─ [IF -DebugCWM=$true — skipped in read-only mode]
   │           │   📥 READ  CWM API: Get-CWMAgreementAddition (stores before-quantities)
   │           │
   │           ├─ FOR EACH active CWM addition:
   │           │   │
   │           │   ├─ Match addition product ID against CDP2CWMProductMapping
   │           │   │   🔒 GATE: -CWMAdditionSearchType (Equals/StartsWith) ⚠️
   │           │   │            "StartsWith" can match multiple additions unexpectedly
   │           │   │
   │           │   ├─ [PhysicalServerQty matched]
   │           │   │   🔒 GATE: -CombineServers
   │           │   │   → $true:  📤 WRITE 0 to CWM  ⚠️ ZEROES the Phys addition
   │           │   │   → $false: 📤 WRITE PhysicalServerQty from CSV
   │           │   │   💾 FILE  AuditLog: before qty, after qty
   │           │   │
   │           │   ├─ [VirtualServerQty matched]
   │           │   │   🔒 GATE: -CombineServers
   │           │   │   → $true:  📤 WRITE 0 to CWM  ⚠️ ZEROES the Virt addition
   │           │   │   → $false: 📤 WRITE VirtualServerQty from CSV
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   ├─ [CombinedServerQty matched]
   │           │   │   🔒 GATE: -CombineServers
   │           │   │   → $true:  📤 WRITE (Phys + Virt) from CSV  ← actual quantity pushed
   │           │   │   → $false: SKIPPED (logged as "[i] Skipped")
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   ├─ [WorkstationQty matched]
   │           │   │   📤 WRITE WorkstationQty from CSV
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   ├─ [DocumentsQty matched]
   │           │   │   📤 WRITE DocumentsQty from CSV
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   ├─ [RecoveryTestingQty matched]
   │           │   │   📤 WRITE RecoveryTestingQty from CSV
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   ├─ [O365UsersQty matched]
   │           │   │   📤 WRITE O365UsersQty from CSV
   │           │   │   💾 FILE  AuditLog
   │           │   │
   │           │   └─ [CoveGBSelected matched]
   │           │       📤 WRITE CoveGBSelected from CSV (integer — fractional GB truncated ⚠️)
   │           │       💾 FILE  AuditLog
   │           │
   │           ├─ [IF -DebugCWM=$true]
   │           │   📥 READ  CWM API: Get-CWMAgreementAddition (after-quantities, 2.5s delay)
   │           │            Displays before/after/change table per addition
   │           │
   │           └─ Exception warnings if Cove qty > 0 but no matching CWM addition found
   │               💾 FILE  ExceptionsLog per unmatched product
   │
   └─ [END]
      💾 FILE  Console summary of all output file paths
```

### Mode 2 Risk Summary

| Risk | Condition | Severity |
|---|---|---|
| ⚠️ Live CWM quantities overwritten | `-NoCWMUpdate=$false` | **CRITICAL** |
| ⚠️ All CSV rows processed without review | `RunMode=UploadOnly` (no GridView) | High |
| ⚠️ CSV from wrong period or parameters pushed | File auto-selection picks wrong CSV | High |
| ⚠️ Phys/Virt additions zeroed | `-CombineServers=$true` (default) with live write | Medium — intended behavior |
| ⚠️ Wrong CombineServers vs CSV generation | CSV generated with different CombineServers | High |
| ⚠️ GB fractional truncation | CoveGBSelected cast to [int] | Low |
| Company excluded silently | In $CWMExcludedCompanies | Low — intended behavior |
| Agreement not found | Wrong CWMAgreementTypes/Name | Medium |
| Multiple agreements matched | Agreement search too broad | Medium |

---

## Mode 3 — Pull + Upload (Combined)
**Triggers:** Interactive `C`, or `RunMode=PullAndUpload`
**Internal state:** `GetCDPUsage=$true`, `LoadCDPUsage=$false`, `ShouldProcessCWMAdditions=$true`
**Purpose:** Pull live Cove data, export files, then immediately push quantities to CWM in the same run.

Mode 3 is a sequential combination of Mode 1 (Cove pull) + Mode 2 Upload sub-path B mechanics, using the in-memory `$Script:CDPUsageCollection` instead of a CSV file. There is no GridView selection — all production customers are processed.

### Workflow Steps

```
[START]
   │
   ├─ [All Cove API pull steps from Mode 1 execute in full]
   │   (EnumerateAccountStatistics, HWM report, AncestorPartners, DeviceHistory)
   │   🔒 All Mode 1 gates apply: -IgnoreDeletedDevices, -TrialState, -DeviceCount,
   │      -DeviceFilter, -PhysServerFUGB, -VirtServerFUGB, -WorkstationFUGB, -Period/-Last
   │
   ├─ 💾 FILE  All Mode 1 output files generated:
   │           *_Statistics.csv, *_Statistics.xlsx (Trial + Production tabs),
   │           *_CDPUsageFile_yyyy-MM_.csv
   │
   ├─ [IF -ConnectCWM=$true]
   │   │
   │   ├─ AuthenticateCWM, LookupCWMProducts, [IF DebugCWM] LookupCWMAgreements
   │   │   (identical to Mode 1 CWM init — same CWM reads/files)
   │   │
   │   └─ FOR EACH customer in $Script:CDPUsageCollection  ← in-memory, NOT from CSV
   │       │   (same loop structure as Mode 2, but data source is live Cove data)
   │       │
   │       ├─ LookupCWMCompany → agreement lookup
   │       │
   │       └─ [ACTION] UpdateCWMQty
   │           │
   │           ├─ [IF $NoCWMUpdate=$true — default] — SUB-PATH A: READ-ONLY
   │           │   📥 READ  CWM additions
   │           │   Display Current vs Proposed (using live Cove data as proposed)
   │           │   💾 FILE  AuditLog: current + proposed
   │           │   return  [NO CWM WRITES]
   │           │
   │           └─ [IF $NoCWMUpdate=$false] — SUB-PATH B: LIVE WRITE ⚠️
   │               📤 WRITE all matched additions via Update-CWMAgreementAddition
   │               🔒 GATE: -CombineServers (same zeroing/combining logic as Mode 2)
   │               💾 FILE  AuditLog per addition
   │               [IF DebugCWM] 📥 READ after-quantities, display before/after table
   │
   └─ [END]
      💾 FILE  Console summary
```

### Mode 3 vs Mode 2 Key Differences

| | Mode 2 (Upload) | Mode 3 (Combined) |
|---|---|---|
| Data source for CWM push | CSV file (historical) | Live Cove API (current) |
| GridView customer selection | Yes (Interactive) / No (Unattended) | No — all production customers |
| IgnoreDeletedDevices effect | None (data already in CSV) | Yes — strips deleted before push ⚠️ |
| TrialState effect | None | Yes — can include trial customers ⚠️ |
| FairUseGB effect | None | Yes — affects CoveGBSelected pushed ⚠️ |
| CDPUsage CSV generated | No | Yes |
| XLS report generated | No | Yes |

### Mode 3 Risk Summary

| Risk | Condition | Severity |
|---|---|---|
| ⚠️ Live CWM quantities overwritten immediately after pull | `-NoCWMUpdate=$false` | **CRITICAL** |
| ⚠️ No review step between pull and push | No GridView in Mode 3 | High |
| ⚠️ Wrong period data pushed | `-Period`/`-Last` wrong | High |
| ⚠️ Deleted devices affect quantities | `-IgnoreDeletedDevices=$false` | Medium |
| ⚠️ Trial customers pushed to CWM | `-TrialState=$true` | Medium |
| ⚠️ CombineServers zeroes split additions | Default behavior | Medium — intended |

---

## Modifier Flags — Cross-Cutting Reference

### Flags that affect ALL three modes

| Flag | Default | Effect on Mode 1 | Effect on Mode 2 | Effect on Mode 3 |
|---|---|---|---|---|
| `-NoCWMUpdate` | `$true` | Read-only CWM snapshot (if ConnectCWM=$true); `$false` = no CWM reads at all | `$true` = snapshot only; `$false` = ⚠️ live CWM write | `$true` = snapshot; `$false` = ⚠️ live CWM write |
| `-ConnectCWM` | `$true` | `$false` = skip entire CWM block; no snapshot | `$false` = script exits at validation ⚠️ | `$false` = script exits at validation ⚠️ |
| `-CWMAgreementBehavior` | `"Type"` | Gates which agreements are read/snapshotted | Gates which agreements are written | Same |
| `-CWMAgreementTypes` | *(array of 4 types)* | Controls agreement type matching | Controls which companies are processed ⚠️ | Same |
| `-CWMAgreementName` | `"CoveDataProtection"` | Used when Behavior="Name" | Used when Behavior="Name" ⚠️ | Same |
| `-CWMAgreementSearchType` | `"Equals"` | Agreement name/type match mode | Broader values risk matching unintended ⚠️ | Same |
| `-CWMAdditionSearchType` | `"Equals"` | Affects which additions are snapshotted | `"StartsWith"` risks matching multiple ⚠️ | Same |
| `-CombineServers` | `$true` | Affects snapshot proposed values | Zeroes Phys/Virt; pushes combined ⚠️ | Same as Mode 2 |
| `-DebugCWM` | `$true` | `$true` = reads + exports CWMAgreements CSV | `$true` = reads + exports CWMAgreements CSV; extra after-read in live mode | Same |
| `-ClearCWMCredentials` | `$false` | Deletes CWM cred file at start | Same | Same |
| `-ClearCDPCredentials` | `$false` | Deletes Cove cred file at start | Same (Cove auth still required) | Same |
| `$CWMExcludedCompanies` | *(hardcoded array)* | Company skipped from snapshot | Company skipped from write ⚠️ | Same |

### Flags that affect Mode 1 and Mode 3 only (Pull path)

| Flag | Default | Effect |
|---|---|---|
| `-Period` | *(current date)* | Month for HWM report pull; wrong value = wrong period data ⚠️ |
| `-Last` | `0` | Months to count back (0=current, 1=prior month) |
| `-PartnerNameOverride` | `"Metrotech"` | Partner scope for all Cove API queries |
| `-DeviceCount` | `5000` | Max devices from EnumerateAccountStatistics; too low = missed devices ⚠️ |
| `-DeviceFilter` | `"AT==1 OR AT==2"` | Device types: AT==1 Backup Manager, AT==2 M365 |
| `-IgnoreDeletedDevices` | `$true` | `$true` = deleted devices excluded from quantities before CWM push ⚠️ |
| `-TrialState` | `$false` | `$true` = InTrial customers included in quantities and CWM push ⚠️ |
| `-PhysServerFUGB` | `0` | Fair Use GB for physical servers; 0 = full SelectedSizeGB as OverageGB |
| `-VirtServerFUGB` | `0` | Fair Use GB for virtual servers |
| `-WorkstationFUGB` | `0` | Fair Use GB for workstations |
| `-DebugCDP` | `$true` | Console verbosity only; no data reads/writes |
| `-Launch` | `$false` | Opens XLS/CSV after export; no data effect |
| `-Delimiter` | `","` | CSV delimiter for exported files |

### Flags that have NO effect on Mode 2 (Upload path)

`-Period`, `-Last`, `-DeviceCount`, `-DeviceFilter`, `-IgnoreDeletedDevices`, `-TrialState`, `-PhysServerFUGB`, `-VirtServerFUGB`, `-WorkstationFUGB`, `-DebugCDP`, `-Launch`

---

## AuditLog Column Behavior by Mode / Sub-path

The AuditLog CSV (`*_AuditLog.csv`) header is:
`Company Id, Company Name, Agreement Name, Agreement Type, Additions Product ID, Additions Description, Effective Date, Additions Quantity, Proposed/Updated Quantity`

The last column header is **dynamic**: `Proposed Quantity` when `NoCWMUpdate=$true` or `ShouldProcessCWMAdditions=$false`; `Updated Quantity` only when actively writing (`NoCWMUpdate=$false` with upload or combined mode).

| Mode | Sub-path | `Effective Date` | `Additions Quantity` | Last column (dynamic) |
|---|---|---|---|---|
| 1 (Pull) | `NoCWMUpdate=$true` (default) | Addition effectiveDate (yyyy-MM-dd) | Current CWM qty | **Proposed Quantity** (from live Cove data) |
| 1 (Pull) | `NoCWMUpdate=$false` | — | — | *(UpdateCWMQty not called; no audit rows)* |
| 2 (Upload) | `NoCWMUpdate=$true` (default) | Addition effectiveDate (yyyy-MM-dd) | Current CWM qty | **Proposed Quantity** (from CSV) |
| 2 (Upload) | `NoCWMUpdate=$false` ⚠️ | Addition effectiveDate (yyyy-MM-dd) | Current CWM qty (before) | **Updated Quantity** (new qty written to CWM) |
| 3 (Combined) | `NoCWMUpdate=$true` (default) | Addition effectiveDate (yyyy-MM-dd) | Current CWM qty | **Proposed Quantity** (from live Cove data) |
| 3 (Combined) | `NoCWMUpdate=$false` ⚠️ | Addition effectiveDate (yyyy-MM-dd) | Current CWM qty (before) | **Updated Quantity** (new qty written to CWM) |

For non-Cove-mapped additions (not in CDP2CWMProductMapping): `Updated Quantity` is always blank.

---

## Safe vs Unsafe Parameter Combinations

| Combination | Safe? | Notes |
|---|---|---|
| Mode 1, `NoCWMUpdate=$true`, `ConnectCWM=$true` | ✅ Safe | Default. Reads CWM for snapshot; no writes |
| Mode 1, `NoCWMUpdate=$false`, `ConnectCWM=$true` | ✅ Safe | No CWM interaction at all (UpdateCWMQty not called) |
| Mode 1, `ConnectCWM=$false` | ✅ Safe | No CWM block entered at all |
| Mode 2, `NoCWMUpdate=$true` | ✅ Safe | Reads CSV + CWM additions; shows proposed; no writes |
| Mode 2, `NoCWMUpdate=$false` | ⚠️ LIVE WRITE | Overwrites CWM addition quantities |
| Mode 3, `NoCWMUpdate=$true` | ✅ Safe | Pull + snapshot; no writes |
| Mode 3, `NoCWMUpdate=$false` | ⚠️ LIVE WRITE | Pull + immediate overwrite; no review step |
| Any mode, `CWMAgreementSearchType="Contains"` | ⚠️ | Risk of matching unintended agreements |
| Any mode, `CWMAdditionSearchType="StartsWith"` | ⚠️ | Risk of matching multiple additions |
| Mode 2/3, `CombineServers` changed vs CSV | ⚠️ | Zeroed additions may not match CSV generation state |
