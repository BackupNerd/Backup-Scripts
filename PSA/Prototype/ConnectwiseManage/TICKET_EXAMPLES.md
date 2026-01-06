# M365 Ticket Templates Preview

**Generated:** 2026-01-06  
**Purpose:** Review M365 ticket templates before deployment

---

## M365 New Ticket Template

```
Cove Data Protection Backup Alert

M365 Tenant          : example.onmicrosoft.com (ID:1234567)
Customer             : Example Corp
Reference            : CWM:12345:Example Corp
Notes                : Additional notes about this tenant
Severity             : Stale
Description          : Stale backup (~2d 6h ago)
Creation Date        : 2025-12-27 10:00:00 (EST) (~3d 2h ago)
Last Timestamp       : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
Oldest Problem       : 2025-12-29 14:32:15 (EST) (~2d 6h ago) | Exchange | CompletedWithErrors

─────────────────────────────────────────────────────────────────
TENANT DETAILS

Storage Usage:
  Selected Data       : 145.5 GB
  Used Storage        : 98.2 GB
  Storage Location    : United States
  Timezone Offset     : UTC-05:00

─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS

M365 Exchange Status: CompletedWithErrors
  Session Start       : 2025-12-29 14:30:01 (EST)
  Last Completed      : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
  Last Success        : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
  Duration            : 00:02:14 | Selected: 45.2 GB | Processed: 1.2 GB | Sent: 450 MB | Errors: 3
  Last Error          : 2025-12-29 14:32:10 (EST) (~2d 6h ago)
  Error Message       : Graph API mailbox access denied for user@example.com

M365 OneDrive Status: Completed
  Session Start       : 2025-12-29 14:35:00 (EST)
  Last Completed      : 2025-12-29 14:38:22 (EST) (~2d 5h ago)
  Last Success        : 2025-12-29 14:38:22 (EST) (~2d 5h ago)
  Duration            : 00:03:22 | Selected: 78.5 GB | Processed: 2.1 GB | Sent: 890 MB | Errors: 0

─────────────────────────────────────────────────────────────────
View Device: https://backup.management/#/device/1234567/cloud-properties/office365/history

This ticket was automatically created by Cove Data Protection Monitoring v10 @ 2026-01-06 00:03:41 (System Time)
```

---

## Systems New Ticket Template

```
Cove Data Protection Backup Alert

Device               : SRV-01 (ID:2345678)
Computer             : SRV-01.example.local
Alias                : Production File Server
Customer             : Example Corp
Reference            : CWM:12345:Example Corp
Notes                : Critical production server
Severity             : Failed
Description          : Backup failed - FS Failed
Creation Date        : 2025-12-20 08:00:00 (EST) (~16d 3h ago)
Last Timestamp       : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
Oldest Problem       : 2025-12-29 02:00:15 (EST) (~2d 18h ago) | File & Folders | Failed

─────────────────────────────────────────────────────────────────
DEVICE DETAILS

Hardware Information:
  OS                  : Windows Server 2019 Standard (17763), 64-bit
  Mfg|Model           : Dell Inc. | PowerEdge R740
  Device Type         : Physical Server | 16 Cores | 64 GB RAM
  IP Address          : 192.0.2.100
  External IP         : 203.0.113.25

Backup Configuration:
  Backup Profile      : Server Full Backup (ID: 123456)
  Retention Policy    : 30 Day Retention (ID: 789012)
  Storage Location    : United States
  Timezone Offset     : UTC-05:00

─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS

File & Folders Status: Failed
  Session Start       : 2025-12-29 02:00:00 (EST)
  Last Completed      : 2025-12-29 03:15:30 (EST) (~2d 8h ago)
  Last Success        : 2025-12-27 02:15:30 (EST) (~4d 8h ago)
  Duration            : 01:15:30 | Selected: 450.5 GB | Processed: 12.3 GB | Sent: 5.2 GB | Errors: 15
  Last Error          : 2025-12-29 03:14:22 (EST) (~2d 8h ago)
  Error Message       : Access denied - D:\Shares\Accounting\file.xlsx [15 errors]

System State Status: Completed
  Session Start       : 2025-12-29 02:00:00 (EST)
  Last Completed      : 2025-12-29 02:05:45 (EST) (~2d 18h ago)
  Last Success        : 2025-12-29 02:05:45 (EST) (~2d 18h ago)
  Duration            : 00:05:45 | Selected: 2.5 GB | Processed: 150 MB | Sent: 75 MB | Errors: 0

─────────────────────────────────────────────────────────────────
View Device: https://backup.management/#/backup/overview/view/1662(panel:device-properties/2345678/summary)

This ticket was automatically created by Cove Data Protection Monitoring v10 @ 2026-01-06 00:03:41 (System Time)
```
  Session Start       : 2025-12-29 03:15:35 (EST)
  Last Completed      : 2025-12-29 03:18:22 (EST) (~2d 8h ago)
  Last Success        : 2025-12-29 03:18:22 (EST) (~2d 8h ago)
  Duration            : 00:02:47 | Selected: 25.2 GB | Processed: 890 MB | Sent: 450 MB | Errors: 0

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-31 00:03:41 (System Time)
```

---

## M365 Update Ticket Template

```
Updated Status - 2025-12-31 11:04:51 (System Time)

Severity             : Stale
Description          : Stale backup (~2d 6h ago)
Last Timestamp       : 2025-12-29 14:32:15 (EST) | Completed
Last Backup          : 2025-12-29 14:32:15 (EST) | Completed
Oldest Problem       : 2025-12-29 14:32:15 (EST) (~2d 6h ago) | Exchange | Completed
Recent Success       : 2025-12-29 14:32:15 (EST) (~2d 6h ago) | OneDrive | Completed
Device Session       : 2025-12-29 14:32:15 (EST) (~2d 6h ago) | Exchange | Completed

Datasource Details:

M365 Exchange Status: Completed
  Session Start       : 2025-12-29 14:30:01 (EST)
  Last Completed      : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
  Last Success        : 2025-12-29 14:32:15 (EST) (~2d 6h ago)
  Duration            : 00:02:14 | Selected: 45.2 GB | Processed: 1.2 GB | Sent: 450 MB | Errors: 0

M365 OneDrive Status: Completed
  Session Start       : 2025-12-29 14:35:00 (EST)
  Last Completed      : 2025-12-29 14:38:22 (EST) (~2d 5h ago)
  Last Success        : 2025-12-29 14:38:22 (EST) (~2d 5h ago)
  Duration            : 00:03:22 | Selected: 78.5 GB | Processed: 2.1 GB | Sent: 890 MB | Errors: 0

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/1234567/summary)

This update was automatically generated by Cove Data Protection Monitoring v05 @ 2025-12-31 11:04:51 (System Time)
```

---

## M365 Close Ticket Template

```
Issue Resolved - 2025-12-31 11:39:15 (System Time)

The backup issue for example.onmicrosoft.com has been resolved.

Last Timestamp       : 2025-12-31 10:45:22 (EST) | Completed
Last Backup          : 2025-12-31 10:45:22 (EST) | Completed
Oldest Problem       : 2025-12-29 14:32:15 (EST) (~2d ago) | Exchange | Completed
Recent Success       : 2025-12-31 10:45:22 (EST) (~just now) | OneDrive | Completed
Device Session       : 2025-12-31 10:45:22 (EST) (~just now) | Exchange | Completed

Datasource Details:

M365 Exchange Status: Completed
  Session Start       : 2025-12-31 10:43:01 (EST)
  Last Completed      : 2025-12-31 10:45:22 (EST) (~just now)
  Last Success        : 2025-12-31 10:45:22 (EST) (~just now)
  Duration            : 00:02:21 | Selected: 45.3 GB | Processed: 850 MB | Sent: 320 MB | Errors: 0

M365 OneDrive Status: Completed
  Session Start       : 2025-12-31 10:48:00 (EST)
  Last Completed      : 2025-12-31 10:50:15 (EST) (~just now)
  Last Success        : 2025-12-31 10:50:15 (EST) (~just now)
  Duration            : 00:02:15 | Selected: 78.5 GB | Processed: 1.1 GB | Sent: 450 MB | Errors: 0

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/1234567/summary)

This ticket was automatically closed by Cove Data Protection Monitoring v05 @ 2025-12-31 11:39:15 (System Time)
```

---

## M365 InProcess Example (Completed -> InProcess)

```
M365 Exchange Status: Completed -> InProcess
  Session Start       : 2025-12-31 11:30:01 (EST)
  Last Completed      : 2025-12-31 10:45:22 (EST) (~45m ago)
  Last Success        : 2025-12-31 10:45:22 (EST) (~45m ago)
  Duration            : 00:02:21 | Selected: 45.3 GB | Processed: 850 MB | Sent: 320 MB | Errors: 0
```

---

## Key Differences from Systems Templates

### M365-Specific Changes:
1. **Header:** "M365 Tenant" instead of "Device"
2. **No Computer field** (M365 doesn't have computer name)
3. **Tenant Details section** instead of "Device Details"
4. **Storage Usage subsection** instead of "Hardware Information"
5. **Datasource names:** "M365 Exchange", "M365 OneDrive", "M365 SharePoint", "M365 Teams"
6. **Portal View ID:** 18439 (M365 specific)

### Maintained Nuances:
1. **InProcess status:** Shows "Completed -> InProcess" format when current session is in progress
2. **Status codes:** All session statuses displayed with names (Completed, Failed, etc.)
3. **Time ago format:** (~2d 6h ago) using existing Format-HoursAsRelativeTime function
4. **Error display:** When errors exist, shown as in current implementation
5. **Timezone display:** Respects UseLocalTime parameter (EST vs UTC)

---

## Error Examples from Historic Tickets

### Example 1: M365 Tenant with Multiple Datasource Errors

**Scenario:** achievementstherapy.com with 62 errors across multiple datasources

```
Cove Data Protection Backup Alert

M365 Tenant          : acme.onmicrosoft.com (ID:5131534)
Customer             : Acme Technologies | Site: 
Reference            : 
Severity             : Warning
Description          : Backup has 62 error(s) - not classified as successful

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/5131534/summary)

Last Timestamp       : 12/24/2025 04:01:36 (UTC) | Completed
Last Backup          : 12/24/2025 04:01:36 (UTC) | Completed
Oldest Problem       : 12/24/2025 02:42:04 (UTC) (just now) | M365 SharePoint | Completed
Recent Success       : 12/24/2025 04:01:36 (UTC) (just now) | M365 Teams | Completed
Device Session       : 12/24/2025 04:01:36 (UTC) (just now) | M365 Teams | Completed

Tenant Details:

Storage Usage:
  Selected Data       : 778.06 GB
  Used Storage        : 0.03 GB
  Storage Location    : United States
  Timezone Offset     : UTC+00:00

Datasource Details:

M365 SharePoint Status: Completed
  Session Start     : 12/24/2025 02:42:04 (UTC)
  Last Completed    : 12/24/2025 02:42:38 (UTC) (just now)
  Last Success      : 12/24/2025 02:42:38 (UTC) (just now)
  Duration          : 00:00:34 | Selected: 104.90 GB | Sent: 0.01 GB | Errors: 0

M365 Exchange Status: Completed
  Session Start     : 12/24/2025 03:15:22 (UTC)
  Last Completed    : 12/24/2025 03:18:30 (UTC) (just now)
  Last Success      : 12/24/2025 03:18:30 (UTC) (just now)
  Duration          : 00:03:08 | Selected: 502.31 GB | Sent: 0.01 GB | Errors: 0

M365 OneDrive Status: Completed
  Session Start     : 12/24/2025 03:16:23 (UTC)
  Last Completed    : 12/24/2025 03:16:35 (UTC) (just now)
  Last Success      : 12/24/2025 03:16:35 (UTC) (just now)
  Duration          : 00:00:12 | Selected: 71.84 GB | Sent: 0.01 GB | Errors: 0

M365 Teams Status: Completed
  Session Start     : 12/24/2025 04:01:08 (UTC)
  Last Completed    : 12/24/2025 04:01:36 (UTC) (just now)
  Last Success      : 12/24/2025 04:01:36 (UTC) (just now)
  Duration          : 00:00:28 | Selected: 0.01 GB | Sent: 0.00 GB | Errors: 0

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-24 02:16:24 (System Time)
```

**Key Points:**
- All datasources show "Completed" status but ticket has 62 errors
- This demonstrates error detection even when status is not "Failed"
- Errors field shows 0 for each datasource (error count may be in detailed logs)

---

### Example 2: M365 Tenant with High Error Count (229 errors)

**Scenario:** documax.com with 229 errors, large data volumes

```
Cove Data Protection Backup Alert

M365 Tenant          : datasafe.onmicrosoft.com (ID:5133022)
Customer             : DataSafe Solutions | Site: 
Reference            : 
Severity             : Warning
Description          : Backup has 229 error(s) - not classified as successful

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/5133022/summary)

Last Timestamp       : 12/24/2025 06:07:47 (UTC) | Completed
Last Backup          : 12/24/2025 06:07:46 (UTC) | Completed
Oldest Problem       : 12/24/2025 03:10:15 (UTC) (~3h ago) | M365 OneDrive | Completed
Recent Success       : 12/24/2025 06:07:46 (UTC) (just now) | M365 SharePoint | Completed
Device Session       : 12/24/2025 06:07:47 (UTC) (just now) | M365 SharePoint | Completed

Tenant Details:

Storage Usage:
  Selected Data       : 743.14 GB
  Used Storage        : 0.00 GB
  Storage Location    : United States
  Timezone Offset     : UTC+00:00

Datasource Details:

M365 SharePoint Status: Completed
  Session Start     : 12/24/2025 06:05:07 (UTC)
  Last Completed    : 12/24/2025 06:07:46 (UTC) (just now)
  Last Success      : 12/24/2025 06:07:46 (UTC) (just now)
  Duration          : 00:02:39 | Selected: 243.50 GB | Sent: 0.00 GB | Errors: 0

M365 Exchange Status: Completed
  Session Start     : 12/24/2025 05:17:19 (UTC)
  Last Completed    : 12/24/2025 05:23:37 (UTC) (just now)
  Last Success      : 12/24/2025 05:23:37 (UTC) (just now)
  Duration          : 00:06:18 | Selected: 286.40 GB | Sent: 0.00 GB | Errors: 0

M365 OneDrive Status: Completed
  Session Start     : 12/24/2025 03:10:15 (UTC)
  Last Completed    : 12/24/2025 03:10:20 (UTC) (~3h ago)
  Last Success      : 12/24/2025 03:10:20 (UTC) (~3h ago)
  Duration          : 00:00:05 | Selected: 213.09 GB | Sent: 0.00 GB | Errors: 0

M365 Teams Status: Completed
  Session Start     : 12/24/2025 03:44:48 (UTC)
  Last Completed    : 12/24/2025 03:45:58 (UTC) (~2h ago)
  Last Success      : 12/24/2025 03:45:58 (UTC) (~2h ago)
  Duration          : 00:01:10 | Selected: 0.15 GB | Sent: 0.00 GB | Errors: 0

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-24 02:16:44 (System Time)
```

**Key Points:**
- Even higher error count (229 errors)
- Large data volumes: 743.14 GB selected across all datasources
- "Oldest Problem" correctly identifies OneDrive as oldest completed session (~3h ago)
- Demonstrates time-based sorting of datasource sessions

---

### Example 3: M365 Tenant with .onmicrosoft.com Domain

**Scenario:** hcsresort.onmicrosoft.com with 37 errors, smaller tenant

```
Cove Data Protection Backup Alert

M365 Tenant          : techserv.onmicrosoft.com (ID:5181312)
Customer             : TechServ MSP | Site: 
Reference            : 
Severity             : Warning
Description          : Backup has 37 error(s) - not classified as successful

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/5181312/summary)

Last Timestamp       : 12/24/2025 07:11:17 (UTC) | Completed
Last Backup          : 12/24/2025 07:11:16 (UTC) | Completed
Oldest Problem       : 12/24/2025 00:22:01 (UTC) (~7h ago) | M365 Exchange | Completed
Recent Success       : 12/24/2025 07:11:16 (UTC) (just now) | M365 OneDrive | Completed
Device Session       : 12/24/2025 07:11:17 (UTC) (just now) | M365 OneDrive | Completed

Tenant Details:

Storage Usage:
  Selected Data       : 29.65 GB
  Used Storage        : 0.03 GB
  Storage Location    : United States
  Timezone Offset     : UTC+00:00

Datasource Details:

M365 SharePoint Status: Completed
  Session Start     : 12/24/2025 05:59:02 (UTC)
  Last Completed    : 12/24/2025 05:59:12 (UTC) (~1h ago)
  Last Success      : 12/24/2025 05:59:12 (UTC) (~1h ago)
  Duration          : 00:00:10 | Selected: 2.90 GB | Sent: 0.00 GB | Errors: 0

M365 Exchange Status: Completed
  Session Start     : 12/24/2025 00:22:01 (UTC)
  Last Completed    : 12/24/2025 00:23:14 (UTC) (~7h ago)
  Last Success      : 12/24/2025 00:23:14 (UTC) (~7h ago)
  Duration          : 00:01:13 | Selected: 24.19 GB | Sent: 0.03 GB | Errors: 0

M365 OneDrive Status: Completed
  Session Start     : 12/24/2025 07:11:14 (UTC)
  Last Completed    : 12/24/2025 07:11:16 (UTC) (just now)
  Last Success      : 12/24/2025 07:11:16 (UTC) (just now)
  Duration          : 00:00:02 | Selected: 2.56 GB | Sent: 0.00 GB | Errors: 0

M365 Teams Status: Completed
  Session Start     : 12/24/2025 05:53:58 (UTC)
  Last Completed    : 12/24/2025 05:54:25 (UTC) (~1h ago)
  Last Success      : 12/24/2025 05:54:25 (UTC) (~1h ago)
  Duration          : 00:00:27 | Selected: 0.00 GB | Sent: 0.00 GB | Errors: 0

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-24 02:17:14 (System Time)
```

**Key Points:**
- Lower error count (37 errors) compared to other examples
- Smaller tenant: 29.65 GB total selected data
- Exchange session is oldest (~7h ago) while OneDrive is most recent (just now)
- Shows datasource completion time spread across several hours
- Demonstrates .onmicrosoft.com domain format

---

### Example 4: UPDATE Template with Errors Example

**Scenario:** Follow-up update for acme.onmicrosoft.com showing error reduction

```
Cove Data Protection Backup Alert - Updated Status

Updated Status - 2025-12-24 04:30:15 (System Time)

Severity             : Warning
Description          : Backup has 45 error(s) - error count reduced from 62
Last Timestamp       : 12/24/2025 04:25:36 (UTC) | Completed
Last Backup          : 12/24/2025 04:25:36 (UTC) | Completed
Oldest Problem       : 12/24/2025 03:15:22 (UTC) (~1h ago) | M365 Exchange | Completed
Recent Success       : 12/24/2025 04:25:36 (UTC) (just now) | M365 Teams | Completed
Device Session       : 12/24/2025 04:25:36 (UTC) (just now) | M365 Teams | Completed

Datasource Details:

M365 SharePoint Status: Completed
  Session Start     : 12/24/2025 04:10:04 (UTC)
  Last Completed    : 12/24/2025 04:10:38 (UTC) (~15m ago)
  Last Success      : 12/24/2025 04:10:38 (UTC) (~15m ago)
  Duration          : 00:00:34 | Selected: 104.90 GB | Sent: 0.00 GB | Errors: 0

M365 Exchange Status: Completed
  Session Start     : 12/24/2025 03:15:22 (UTC)
  Last Completed    : 12/24/2025 03:18:30 (UTC) (~1h ago)
  Last Success      : 12/24/2025 03:18:30 (UTC) (~1h ago)
  Duration          : 00:03:08 | Selected: 502.31 GB | Sent: 0.00 GB | Errors: 0

M365 OneDrive Status: Completed
  Session Start     : 12/24/2025 04:20:23 (UTC)
  Last Completed    : 12/24/2025 04:20:35 (UTC) (~5m ago)
  Last Success      : 12/24/2025 04:20:35 (UTC) (~5m ago)
  Duration          : 00:00:12 | Selected: 71.84 GB | Sent: 0.00 GB | Errors: 0

M365 Teams Status: Completed
  Session Start     : 12/24/2025 04:25:08 (UTC)
  Last Completed    : 12/24/2025 04:25:36 (UTC) (just now)
  Last Success      : 12/24/2025 04:25:36 (UTC) (just now)
  Duration          : 00:00:28 | Selected: 0.01 GB | Sent: 0.00 GB | Errors: 0

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/5131534/summary)

This update was automatically generated by Cove Data Protection Monitoring v05 @ 2025-12-24 04:30:15 (System Time)
```

**Key Points:**
- Description shows error count reduction: "45 error(s) - error count reduced from 62"
- All datasources show completed status with recent sessions
- Relative time indicators help track when each datasource last ran
- Update format focuses on current status, not full device details

---

### Example 5: CLOSE Template with Errors Resolved Example

**Scenario:** Issue resolved after errors cleared for datasafe.onmicrosoft.com

```
Cove Data Protection Backup Alert - Issue Resolved

Issue Resolved - 2025-12-24 08:15:30 (System Time)

The backup issue for datasafe.onmicrosoft.com has been resolved.

Last Timestamp       : 12/24/2025 08:10:47 (UTC) | Completed
Last Backup          : 12/24/2025 08:10:47 (UTC) | Completed
Oldest Problem       : 12/24/2025 06:05:07 (UTC) (~2h ago) | M365 SharePoint | Completed
Recent Success       : 12/24/2025 08:10:47 (UTC) (just now) | M365 Teams | Completed
Device Session       : 12/24/2025 08:10:47 (UTC) (just now) | M365 Teams | Completed

Datasource Details:

M365 SharePoint Status: Completed
  Session Start     : 12/24/2025 06:05:07 (UTC)
  Last Completed    : 12/24/2025 06:07:46 (UTC) (~2h ago)
  Last Success      : 12/24/2025 06:07:46 (UTC) (~2h ago)
  Duration          : 00:02:39 | Selected: 243.50 GB | Sent: 0.00 GB | Errors: 0

M365 Exchange Status: Completed
  Session Start     : 12/24/2025 07:20:19 (UTC)
  Last Completed    : 12/24/2025 07:26:37 (UTC) (~45m ago)
  Last Success      : 12/24/2025 07:26:37 (UTC) (~45m ago)
  Duration          : 00:06:18 | Selected: 286.40 GB | Sent: 0.00 GB | Errors: 0

M365 OneDrive Status: Completed
  Session Start     : 12/24/2025 08:05:15 (UTC)
  Last Completed    : 12/24/2025 08:05:20 (UTC) (~5m ago)
  Last Success      : 12/24/2025 08:05:20 (UTC) (~5m ago)
  Duration          : 00:00:05 | Selected: 213.09 GB | Sent: 0.00 GB | Errors: 0

M365 Teams Status: Completed
  Session Start     : 12/24/2025 08:10:18 (UTC)
  Last Completed    : 12/24/2025 08:10:47 (UTC) (just now)
  Last Success      : 12/24/2025 08:10:47 (UTC) (just now)
  Duration          : 00:01:10 | Selected: 0.15 GB | Sent: 0.00 GB | Errors: 0

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/18439(panel:device-properties/5133022/summary)

This ticket was automatically closed by Cove Data Protection Monitoring v05 @ 2025-12-24 08:15:30 (System Time)
```

**Key Points:**
- Clear resolution message at top
- All datasources show successful completion
- Error count dropped from 229 to 0
- Timing information shows all datasources completed within last 2 hours
- Close template includes full datasource details for documentation

---

## Systems Device Error Examples from Historic Tickets

### Example 1: Critical Failure - Disk Space Error

**Scenario:** admin-pc_fq8hm with disk space error, System State failed

```
Cove Data Protection Backup Alert

Device               : ws-admin_abc123 (ID:3148735)
Computer Name        : WS-ADMIN
Customer             : Alpha Industries | Site: 
Reference            : 
Severity             : Critical
Description          : Backup failed - 2025-12-21 05:45:25 (UTC) - Insufficient disk space for backup (ID: 3996)

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/2109811(panel:device-properties/3148735/summary)

Last Timestamp       : 12/23/2025 15:43:34 (UTC) | Failed
Last Backup          : 12/11/2025 05:45:24 (UTC) | Completed
Oldest Problem       : 12/11/2025 05:45:05 (UTC) (~12d 10h ago) | System State | Failed
Recent Success       : 12/11/2025 05:45:24 (UTC) (~12d 10h ago) | File & Folders | Completed
Device Session       : 12/23/2025 15:43:34 (UTC) (~15m ago) | File & Folders | Completed with warnings

Device Details:

Hardware Information:
  OS                  : Windows 10 Pro (19045), 64-bit
  Manufacturer        : Dell Inc.
  Model               : Precision T1700
  Device Type         : Physical Workstation
  IP Address          : 192.0.2.80
  External IP         : 198.51.100.38

Backup Configuration:
  Backup Profile      : Full Backup (ID: 33843)
  Retention Policy    : Webistix Policy (ID: 1156313)
  Storage Location    : United States
  Timezone Offset     : UTC-05:00

Datasource Details:

File & Folders Status: Completed with warnings
  Session Start     : 12/23/2025 05:40:25 (UTC)
  Last Completed    : 12/23/2025 05:44:47 (UTC) (~10h ago)
  Last Success      : 12/11/2025 05:45:24 (UTC) (~12d 10h ago)
  Duration          : 00:04:22 | Selected: 193.01 GB | Sent: 0.18 GB | Errors: 1
  Error Session     : 3996 | 2025-12-21 05:45:25 (UTC)
  Last Error        : Insufficient disk space for backup

System State Status: Failed
  Session Start     : 12/11/2025 05:45:05 (UTC)
  Last Completed    : 12/11/2025 05:45:56 (UTC) (~12d 10h ago)
  Last Success      : 12/06/2025 05:43:56 (UTC) (~17d 10h ago)
  Duration          : 00:00:51 | Selected: 22.52 GB | Sent: 0.00 GB | Errors: 1
  Error Session     : 3986 | 2025-12-11 05:45:56 (UTC)
  Last Error        : There is no data available for the backup. Please check your settings and environment.

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-23 11:07:16 (System Time)
```

**Key Points:**
- Critical severity due to System State failure
- File & Folders showing "Completed with warnings" (not full success)
- Disk space error preventing proper backup
- System State failed 12 days ago (oldest problem)
- "Last Error" field shows specific error messages with session IDs

---

### Example 2: Critical Failure - macOS WebDAV Protocol Error

**Scenario:** adams-laptop.local_h9dhr with WebDAV protocol failure, 42 errors

```
Cove Data Protection Backup Alert

Device               : mac-laptop_xyz456 (ID:4858892)
Computer Name        : MAC-LAPTOP.local
Customer             : Beta Solutions | Site: 
Reference            : 
Severity             : Critical
Description          : Backup failed - 2025-12-19 18:58:55 (UTC) - [WebDAV Protocol] Failed to rename resource :  (404) (ID: 344)

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/2782386(panel:device-properties/4858892/summary)

Last Timestamp       : 12/19/2025 18:59:45 (UTC) | Failed
Last Backup          : 12/18/2025 18:46:36 (UTC) | Completed
Oldest Problem       : 12/19/2025 09:17:59 (UTC) (~2d 14h ago) | File & Folders | Failed
Recent Success       : 12/18/2025 18:46:36 (UTC) (~3d 5h ago) | File & Folders | Completed
Device Session       : 12/19/2025 18:59:45 (UTC) (~2d 5h ago) | File & Folders | Failed

Device Details:

Hardware Information:
  OS                  : macOS Tahoe (26.1.0), Intel-based
  Manufacturer        : Apple, Inc
  Model               : Mac15,13
  Device Type         : Physical Workstation
  IP Address          : 127.0.2.1;192.0.2.250
  External IP         : 203.0.113.32

Backup Configuration:
  Backup Profile      : 24 hour RPO (ID: 222354)
  Retention Policy    : 30 days (ID: 1118696)
  Storage Location    : United States
  Timezone Offset     : UTC-08:00

Datasource Details:

File & Folders Status: Failed
  Session Start     : 12/19/2025 09:17:59 (UTC)
  Last Completed    : 12/19/2025 18:59:45 (UTC) (~2d 5h ago)
  Last Success      : 12/18/2025 18:46:36 (UTC) (~3d 5h ago)
  Duration          : 09:41:46 | Selected: 79.92 GB | Sent: 1.41 GB | Errors: 42
  Error Session     : 344 | 2025-12-19 18:58:55 (UTC)
  Last Error        : [WebDAV Protocol] Failed to rename resource :  (404)

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-22 00:21:38 (System Time)
```

**Key Points:**
- macOS device with WebDAV protocol error
- Long session duration: 9 hours 41 minutes before failing
- 42 errors accumulated during session
- WebDAV protocol specific error (404 - resource not found)
- Multiple IP addresses shown (localhost + local network)

---

### Example 3: Stale Backup - macOS iMac

**Scenario:** ahabib-imac_g04ml with no backup for 137.9 hours (~5.7 days)

```
Cove Data Protection Backup Alert

Device               : imac-user_def789 (ID:3266350)
Computer Name        : IMAC-USER24
Customer             : Gamma Services | Site: 
Reference            : 
Severity             : Stale
Description          : No successful backup in 137.9 hours

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/2155090(panel:device-properties/3266350/summary)

Last Timestamp       : 12/17/2025 17:13:39 (UTC) | Completed
Last Backup          : 12/17/2025 17:13:38 (UTC) | Completed
Oldest Problem       : 12/17/2025 17:12:01 (UTC) (~5d 17h ago) | File & Folders | Completed
Recent Success       : 12/17/2025 17:13:38 (UTC) (~5d 17h ago) | File & Folders | Completed
Device Session       : 12/17/2025 17:13:39 (UTC) (~5d 17h ago) | File & Folders | Completed

Device Details:

Hardware Information:
  OS                  : macOS Sequoia (15.7.3), Intel-based
  Manufacturer        : Apple, Inc
  Model               : iMac21,1
  Device Type         : Physical Workstation
  IP Address          : 192.0.2.76
  External IP         : 198.51.100.15

Backup Configuration:
  Backup Profile      : JSP- OSX - Home Folders - excl One Drive (ID: 99901)
  Retention Policy    : Webistix Policy (ID: 1156313)
  Storage Location    : United States
  Timezone Offset     : UTC-05:00

Datasource Details:

File & Folders Status: Completed
  Session Start     : 12/17/2025 17:12:01 (UTC)
  Last Completed    : 12/17/2025 17:13:38 (UTC) (~5d 17h ago)
  Last Success      : 12/17/2025 17:13:38 (UTC) (~5d 17h ago)
  Duration          : 00:01:37 | Selected: 11.59 GB | Sent: 0.20 GB | Errors: 0

This ticket was automatically created by Cove Data Protection Monitoring v05 @ 2025-12-23 11:07:19 (System Time)
```

**Key Points:**
- Stale severity - last backup was successful but too old (137.9 hours ago)
- Last session completed successfully (no failures)
- macOS iMac device
- Small data set: 11.59 GB selected, only 0.20 GB sent (minimal changes)
- Quick backup duration: 1 minute 37 seconds

---

### Example 4: UPDATE Template - Systems Device Error Reduction

**Scenario:** Follow-up update for ws-admin_abc123 showing System State recovery

```
Cove Data Protection Backup Alert - Updated Status

Updated Status - 2025-12-24 06:15:30 (System Time)

Severity             : Warning
Description          : File & Folders completed with warnings - System State still failing
Last Timestamp       : 12/24/2025 05:44:47 (UTC) | Completed with warnings
Last Backup          : 12/23/2025 05:44:47 (UTC) | Completed with warnings
Oldest Problem       : 12/11/2025 05:45:05 (UTC) (~13d ago) | System State | Failed
Recent Success       : 12/23/2025 05:44:47 (UTC) (~25m ago) | File & Folders | Completed with warnings
Device Session       : 12/24/2025 05:44:47 (UTC) (just now) | File & Folders | Completed with warnings

Datasource Details:

File & Folders Status: Completed with warnings
  Session Start     : 12/24/2025 05:40:25 (UTC)
  Last Completed    : 12/24/2025 05:44:47 (UTC) (just now)
  Last Success      : 12/24/2025 05:44:47 (UTC) (just now)
  Duration          : 00:04:22 | Selected: 193.01 GB | Sent: 0.15 GB | Errors: 0

System State Status: Failed
  Session Start     : 12/11/2025 05:45:05 (UTC)
  Last Completed    : 12/11/2025 05:45:56 (UTC) (~13d ago)
  Last Success      : 12/06/2025 05:43:56 (UTC) (~18d ago)
  Duration          : 00:00:51 | Selected: 22.52 GB | Sent: 0.00 GB | Errors: 1
  Error Session     : 3986 | 2025-12-11 05:45:56 (UTC)
  Last Error        : There is no data available for the backup. Please check your settings and environment.

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/2109811(panel:device-properties/3148735/summary)

This update was automatically generated by Cove Data Protection Monitoring v05 @ 2025-12-24 06:15:30 (System Time)
```

**Key Points:**
- Severity reduced from Critical to Warning (partial recovery)
- File & Folders now completing (though with warnings)
- System State still failing - no recovery yet
- Shows mixed datasource status (one improving, one stuck)

---

### Example 5: CLOSE Template - Systems Device Issue Resolved

**Scenario:** Issue resolved after macOS laptop started backing up again

```
Cove Data Protection Backup Alert - Issue Resolved

Issue Resolved - 2025-12-24 18:30:15 (System Time)

The backup issue for mac-laptop_xyz456 has been resolved.

Last Timestamp       : 12/24/2025 18:25:45 (UTC) | Completed
Last Backup          : 12/24/2025 18:25:45 (UTC) | Completed
Oldest Problem       : 12/24/2025 09:17:59 (UTC) (~9h ago) | File & Folders | Completed
Recent Success       : 12/24/2025 18:25:45 (UTC) (just now) | File & Folders | Completed
Device Session       : 12/24/2025 18:25:45 (UTC) (just now) | File & Folders | Completed

Datasource Details:

File & Folders Status: Completed
  Session Start     : 12/24/2025 09:17:59 (UTC)
  Last Completed    : 12/24/2025 18:25:45 (UTC) (just now)
  Last Success      : 12/24/2025 18:25:45 (UTC) (just now)
  Duration          : 09:07:46 | Selected: 79.92 GB | Sent: 1.38 GB | Errors: 0

View Device in Cove Portal:
https://backup.management/#/backup/overview/view/2782386(panel:device-properties/4858892/summary)

This ticket was automatically closed by Cove Data Protection Monitoring v05 @ 2025-12-24 18:30:15 (System Time)
```

**Key Points:**
- Clear resolution message
- WebDAV error resolved - backup completed successfully
- Long duration session (9 hours) but no errors
- 1.38 GB sent (good data transfer)
- All datasources now showing Completed status

---

## Review Checklist

Please verify:

### M365 Templates:
- [ ] M365 Tenant field format acceptable (includes ID)
- [ ] No Computer field is correct for M365
- [ ] Tenant Details section structure appropriate
- [ ] Storage Usage fields are useful
- [ ] Datasource Details follow same format as Systems
- [ ] InProcess status display matches requirement ("Completed -> InProcess")
- [ ] Portal link structure correct (view ID: 18439)
- [ ] Footer messages appropriate
- [ ] Error examples realistic and helpful (62-229 errors)
- [ ] CREATE/UPDATE/CLOSE error scenarios cover common M365 cases

### Systems Templates:
- [ ] Device field format acceptable (includes ID)
- [ ] Computer Name field present and useful
- [ ] Device Details section structure appropriate
- [ ] Hardware Information fields comprehensive
- [ ] Backup Configuration section useful
- [ ] Datasource Details consistent with M365 format
- [ ] Portal link structure correct (view ID: 2782386/2109811/2155090)
- [ ] Footer messages appropriate
- [ ] Error examples realistic and helpful (disk space, protocol errors, stale)
- [ ] CREATE/UPDATE/CLOSE error scenarios cover common Systems cases

### Both Template Types:
- [ ] "Last Error" field only shows when errors exist (conditional display)
- [ ] Time ago format consistent (~X days Y hr ago)
- [ ] Status codes shown with names (Completed, Failed, etc.)
- [ ] InProcess hybrid check working ("Completed -> InProcess")
- [ ] Timezone display respects UseLocalTime parameter
- [ ] All timing fields with status indicators (| Completed, | Failed)
- [ ] Error examples show progression (CREATE → UPDATE → CLOSE)
- [ ] Real historic data used (not fabricated examples)

**If approved:** Changes are already implemented in script
**If changes needed:** Please specify adjustments

