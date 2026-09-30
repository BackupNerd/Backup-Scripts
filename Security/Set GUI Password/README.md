# Bulk Set GUI Password (v15)

PowerShell script to bulk **set** or **wipe** the Cove/N-able Backup Manager GUI password across
selected devices, using the N-able Backup.Management JSON-RPC API.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- A Backup.Management API user (email + password) with access to the target partner(s)/devices
- `Out-GridView` available - built into `Microsoft.PowerShell.Utility` on both Windows PowerShell
  5.1 and PowerShell 7+, no separate module install needed. Windows-only and requires a Windows
  Desktop session (won't work on Windows Server Core or Nano Server, or on non-Windows OSes).

## Quick start

```powershell
# Interactive - prompts for credentials (first run only), partner, devices, and password
.\BulkSetGUIPassword.v15.ps1

# Wipe the GUI password instead of setting one
.\BulkSetGUIPassword.v15.ps1 -WipeGUIPassword

# Target one specific device by exact name, skip partner selection entirely
.\BulkSetGUIPassword.v15.ps1 -DeviceName "NABLE-CS66194"

# Fully non-interactive (e.g. scheduled task) - supply the password directly
.\BulkSetGUIPassword.v15.ps1 -DeviceName "NABLE-CS16194_xdpr1" -GUIPassword 'P@ssw0rd'

# Fully non-interactive (e.g. scheduled task) - supply the password directly
.\BulkSetGUIPassword.v15.ps1 -DeviceName "NABLE-CS16194_xdpr1" -GUIPassword (ConvertTo-SecureString 'P@ssw0rd' -AsPlainText -Force)

```

## Parameters

| Parameter | Set | Description |
|---|---|---|
| `-SetGUIPassword` | SetGUIPW (default) | Prompts for a password to apply to selected devices. |
| `-GUIPassword <SecureString\|string>` | SetGUIPW | Supply the password directly - accepts a `SecureString` or a plain string (auto-converted). Skips the interactive masked/confirm prompt. |
| `-RestoreOnly` | SetGUIPW | Allow restores while the GUI password is set (`restore_only allow`). |
| `-WipeGUIPassword` | WipeGUIPW | Clears the GUI password from selected devices instead of setting one. |
| `-AllPartners` | both | Skip the interactive partner picker; operate directly on your own authenticated partner. |
| `-DeviceName <string>` | both | Target device(s) by **exact** Device Name match. Bypasses partner selection entirely - searches your whole authenticated partner tree (`RecurseSubPartners`). 0 matches re-prompts nothing (exits cleanly), 1 match auto-selects, 2+ matches opens a picker. |
| `-AllDevices` | both | Skip the interactive device picker; operate on every device returned by the partner/device-name query. |
| `-ClearCredentials` | any | Deletes the stored credential file and re-prompts for login. |
| `-DebugCDP` | any | Prints one masked line per device (`password ****`, never the real value) showing the command sent. **Do not use PowerShell's built-in `-Debug`/`-Verbose`** instead - see Security notes. |
| `-APICredentialFile <path>` | any | Override the credential file location. Default: `C:\ProgramData\MXB\<computer>_<user>_API_Credentials.Secure.xml`. |

## Typical flow

1. **Authenticate** - loads (or creates) a DPAPI-encrypted credential file, logs in, retries
   transient network errors automatically, and auto-refreshes the session (visa) if it's older
   than 10 minutes.
2. **Resolve partner** - either your own authenticated partner (`-AllPartners`), a device-name
   search (`-DeviceName`), or an interactive name search (case-insensitive substring, disambiguates
   multiple matches with a breadcrumb-style grid picker).
3. **Enumerate & select devices** - lists devices under the resolved scope; pick one or more via
   grid (or `-AllDevices` to skip picking).
4. **Enter/confirm password** (Set mode only) - masked entry (asterisks, not the real characters),
   typed twice, restarts from scratch on a mismatch. Skipped if `-GUIPassword` was supplied.
5. **Apply** - sends a per-device `SendAccountCommand` (`set gui password`) to each device's home
   node, with a progress bar.

## Security notes

- Credentials are stored via `Export-Clixml` (DPAPI-encrypted `SecureString`, tied to the current
  user + machine) - never in plaintext on disk.
- The GUI password is decrypted to plaintext only for the instant it's needed (building the API
  request body / the double-entry compare) and is never written to the console.
- **Do not run this script with PowerShell's built-in `-Debug` or `-Verbose` common parameters.**
  They also enable `Invoke-RestMethod`'s own internal HTTP tracing, which dumps the *raw, unmasked*
  request body (including the real password) and the `Authorization: Bearer <visa>` header. Use
  `-DebugCDP` instead - it only prints a masked, single-line summary per device.
- `-GUIPassword <plain string>` (or `-GUIPassword (ConvertTo-SecureString ... -AsPlainText -Force)`)
  is convenient for automation but puts the password in your shell history / process command line.
  Prefer `-GUIPassword (Read-Host -AsSecureString)` for interactive use, or source it from a
  pre-encrypted string when scripting.

## Known limitations

- `-DeviceName` matches on Device Name (`AN`) only, not Computer Name (`MN`), and is an **exact**
  match (no wildcards).
- Match counts from the partner-name search are "at least N" - the API appears to cap total
  filtered results (~50 seen in testing) regardless of `childrenLimit`, so refine your search text
  if the partner you want isn't listed.
- Requires `Out-GridView` for any interactive selection step (partner search disambiguation,
  device picker) - not usable unattended unless `-AllPartners`/`-DeviceName` + `-AllDevices` +
  `-GUIPassword` are all supplied together.

## Version history

- **v15** - Added `-DeviceName` (bypass partner selection, target by exact device name), masked
  double-entry password confirmation, optional `-GUIPassword`, `-DebugCDP` (decoupled from native
  `-Debug`). Fixed a `-WipeGUIPassword` ordering bug where the password was cleared *after* the
  request body was already built (a no-op left over from v13). Added `Documents`/M365 exclusion
  and quote-escaping to the `-DeviceName` filter. Fixed a leaked/never-zeroed BSTR buffer in the
  per-device password decrypt (now freed+zeroed immediately after use, matching the rest of the
  script's SecureString handling). Consolidated the Login and generic API retry/backoff logic into
  one shared helper; minor cleanup (removed a dead variable, parameterized `Send-RemoteCommand`
  instead of relying on an ambient loop variable).
- **v14** - Modernized credential storage (XML/SecureString, replacing the flat-text file) and
  replaced the deprecated `GetPartnerInfo(name=)` lookup with `GetPartnerInfoById`/`GetPartnerTree`,
  plus a colorized/progress-bar console experience.
- **v13** - Original release (2021).
