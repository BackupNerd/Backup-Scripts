<# ----- About: ----
    # CoveBackupNow - Backup Recently Accessed Files via ClientTool
    # Revision v02 - 2026-03-03
    # Author: Eric Harless, Head Backup Nerd - N-able
    # GitHub: https://github.com/BackupNerd
#>## end About

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties of
    # merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
#>## end Legal

<# ----- Compatibility: ----
    # For use with N-able Cove Data Protection Backup Manager (Backup Manager must be installed locally)
    # Requires: ClientTool PowerShell Module (https://github.com/BackupNerd/ClientTool)
    # Requires: PowerShell 5.1 or higher
#>## end Compatibility

<# ----- Behavior: ----
    # - Scans a specified path for files matching a chosen timestamp filter within the last N hours
    # - Supports three timestamp filter modes: Write (default), Access, Created
    # - On startup, checks NTFS LastAccessTime tracking status when -FilterBy Access is chosen
    # - Writes a timestamped plain-text selection list file (one path per line)
    # - Applies discovered file paths as FileSystem backup selections via ClientTool
    #   WITHOUT clearing existing selections (accumulates by default)
    # - Triggers an immediate FileSystem backup via ClientTool
    #
    # IMPORTANT Cove Retention & Multi-Run Safety:
    #   Cove Data Protection enhanced retention policies can retain the LAST backup session 
    #   of each day for up to 30 days. When intra-daily sessions expire, only the files present
    #   in the FINAL successful daily session are retained (up to 365 days) in the Daily retention tier.
    #
    #   If this script clears selections on every run and only re-adds files changed in the
    #   last N hours, then by end-of-day the selection only contains the most recently
    #   touched files. Files that were backed up in earlier runs but are NOT in the final
    #   run's selection will be ABSENT from the last session — and therefore will NOT be
    #   retained when intra-day sessions are purged. They are lost from daily retention.
    #
    #   To prevent this, selections are ACCUMULATED by default: new paths are ADDED to
    #   whatever is already selected, so the selection grows throughout the day and the
    #   final session of the day covers everything that changed since the last clear.
    #   Use -SmartClear (recommended for scheduled tasks) to automatically clear only on
    #   the first run of each new calendar day, accumulating on all subsequent runs.
    #   Use -ClearFirst to always force a clean slate regardless of day.
    #
    # Parameters:
    #   -ScanPath    [Required] Root path to scan for recently changed/accessed files
    #   -Hours       Lookback window in hours (default: 1, decimals supported e.g. 0.5 = 30 min)
    #   -FilterBy    Which file timestamp to filter on (default: Write)
    #                  Write   = LastWriteTime  — content was saved/changed        (RECOMMENDED)
    #                  Access  = LastAccessTime — file was opened or read          (UNRELIABLE on most systems — see note below)
    #                  Created = CreationTime   — file was created on this volume
    #
    #                NOTE on Access: Windows 8+ disables LastAccessTime updates by default
    #                to reduce disk I/O. If disabled, the Access filter will return zero or
    #                stale results. The script checks this automatically and warns you.
    #                To check status: fsutil behavior query disablelastaccess
    #                To enable:       fsutil behavior set disablelastaccess 0  (admin required)
    #
    #   -ExportPath  Where to save the selection list file (default: script directory)
    #   -BatchSize   Number of file paths per ClientTool selection call (default: 50)
    #   -SmartClear  Auto-clear schedule for selections:
    #                  Daily   = clear on first run of each calendar day (default)
    #                  Weekly  = clear on first run of each week (resets on Sunday)
    #                  Monthly = clear on first run of each month (resets on the 1st)
    #                  Off     = accumulate indefinitely, never auto-clear
    #                State is persisted in: CoveBackupNow_LastClear.dat (script folder)
    #   -ClearFirst  Always clear existing FileSystem selections before applying new ones
    #                (use for manual or one-off clean-slate runs)
    #   -NoBackup    Set selections only — do not trigger a backup job
    #   -NoGridView  Skip the interactive file review GridView (for unattended/scheduled runs)
    #   -FromFile    Load paths from a previously generated selection list file
    #   -AlwaysInclude  Pipe-separated list of paths to always include in every backup selection,
    #                regardless of timestamp scan results (e.g. 'C:\critical.db|D:\config\app.ini').
    #                Wildcards (* ?) are not permitted. Applied to ClientTool at -AlwaysIncludePriority.
    #   -AlwaysIncludePriority  Priority for AlwaysInclude paths (default: High)
    #
    # Usage Examples:
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Users\eric\Documents"
    #   .\CoveBackupNow.v01.ps1 -ScanPath "D:\Projects" -Hours 2
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -FilterBy Write
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -FilterBy Access
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -FilterBy Created
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -NoBackup
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -SmartClear Daily      # clear once per day (default)
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -SmartClear Weekly     # clear once per week (resets Sunday)
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -SmartClear Monthly    # clear on 1st of each month
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -SmartClear Off        # accumulate indefinitely
    #   .\.CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -ClearFirst -Priority High
    #   .\CoveBackupNow.v01.ps1 -FromFile "C:\Temp\CoveBackupNow_WriteFilter_SelectionList.txt"
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -WhatIf
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -NoGridView
    #   .\CoveBackupNow.v01.ps1 -ScanPath "C:\Data" -AlwaysInclude "C:\critical.db|D:\config\app.ini"
    #   .\CoveBackupNow.v01.ps1 -AlwaysInclude "C:\Data\important.docx"  # always include; uses default ScanPath too
#>## end Behavior

[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$false,  HelpMessage="Root path to scan for recently accessed files")]
    [string]$ScanPath = "Q:\",

    [Parameter(Mandatory=$false,  HelpMessage="Lookback window in hours for the timestamp filter (decimals supported, e.g. 0.5 = 30 minutes)")]
    [double]$Hours = 1,

    [Parameter(Mandatory=$false,  HelpMessage="Which file timestamp to filter on: Write=LastWriteTime (default/recommended), Access=LastAccessTime (unreliable on most Windows 8+ systems), Created=CreationTime")]
    [ValidateSet('Write','Access','Created')]
    [string]$FilterBy = 'Write',

    [Parameter(Mandatory=$false,  HelpMessage="Output directory for the selection list file")]
    [string]$ExportPath = "$PSScriptRoot",

    [Parameter(Mandatory=$false,  HelpMessage="Number of paths per ClientTool selection call (avoid CLI length limits)")]
    [ValidateRange(1, 200)]
    [int]$BatchSize = 50,

    [Parameter(Mandatory=$false,  HelpMessage="Backup selection priority")]
    [ValidateSet('Low','Normal','High')]
    [string]$Priority = 'Normal',

    [Parameter(Mandatory=$false,  HelpMessage="Clear existing FileSystem selections before applying new ones (default is to accumulate - safe for multiple runs per day)")]
    [bool]$ClearFirst=$false,

    [Parameter(Mandatory=$false,  HelpMessage="Auto-clear schedule: Daily=first run of each day, Weekly=first run of each week (resets Sunday), Monthly=first run of each month (resets 1st), Off=accumulate indefinitely")]
    [ValidateSet('Off','Daily','Weekly','Monthly')]
    [string]$SmartClear = 'Daily',

    [Parameter(Mandatory=$false,  HelpMessage="Apply selections but do not trigger a backup job")]
    [bool]$NoBackup=$false,

    [Parameter(Mandatory=$false,  HelpMessage="Skip the interactive file review GridView (for unattended/scheduled runs)")]
    [bool]$NoGridView=$true,

    [Parameter(Mandatory=$false,  HelpMessage="Load paths from a previously generated selection list file instead of scanning")]
    [string]$FromFile,

    [Parameter(Mandatory=$false,  HelpMessage="Pipe-separated list of paths to always include in selections regardless of timestamp filter results (e.g. 'C:\Data\critical.db|D:\Config\app.ini'). Wildcards (* ?) not allowed.")]
    [string]$AlwaysInclude = "",

    [Parameter(Mandatory=$false,  HelpMessage="Backup selection priority for AlwaysInclude paths (default: High - these are always-critical items)")]
    [ValidateSet('Low','Normal','High')]
    [string]$AlwaysIncludePriority = 'High',

    [Parameter(Mandatory=$false,  HelpMessage="Generate 2 random .txt files in the script directory each run (useful for testing that new files are picked up and backed up)")]
    [switch]$CreateTestData = $false
)

#region ----- Script Location Change ----
    if ($PSScriptRoot) { Set-Location $PSScriptRoot }
#endregion

#region ----- Environment, Variables, Names and Paths ----
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $ScriptVersion  = "02"
    $ScriptName     = "CoveBackupNow"
    # Search order: alongside script, parent\Modules, standard PS module paths
    $ModulePath     = Join-Path $PSScriptRoot "..\Modules\ClientTool\ClientTool.psd1"
    $ModuleCandidates = @(
        $ModulePath
        (Join-Path $PSScriptRoot "Modules\ClientTool\ClientTool.psd1")
        (Join-Path $PSScriptRoot "ClientTool\ClientTool.psd1")
    ) + ($env:PSModulePath -split ';' | ForEach-Object { Join-Path $_ "ClientTool\ClientTool.psd1" })
    $SelectionFile      = Join-Path $ExportPath "${ScriptName}_${FilterBy}Filter_SelectionList.txt"
    $SmartClearStateFile = Join-Path $PSScriptRoot "CoveBackupNow_LastClear.dat"
    $CutoffTime         = (Get-Date).AddHours(-$Hours)

    # Map -FilterBy choice to the actual FileInfo property name and a readable label
    $FilterByProperty   = switch ($FilterBy) {
        'Access'  { 'LastAccessTime' }
        'Write'   { 'LastWriteTime'  }
        'Created' { 'CreationTime'   }
    }
    $FilterByLabel      = switch ($FilterBy) {
        'Access'  { 'LastAccessTime  (last opened/read)'    }
        'Write'   { 'LastWriteTime   (last content change)' }
        'Created' { 'CreationTime    (file creation date)'  }
    }

    # Validate and create ExportPath if needed
    if ($ExportPath -match '^-') {
        Write-Error "ExportPath '$ExportPath' looks like a switch parameter, not a directory path. Check your parameter order."
        exit 1
    }
    if (-not (Test-Path $ExportPath -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
            Write-Host "  Created  | ExportPath directory: $ExportPath" -ForegroundColor DarkCyan
        } catch {
            Write-Error "ExportPath '$ExportPath' does not exist and could not be created: $_"
            exit 1
        }
    }

    # Validate mutual exclusivity
    if ($ClearFirst -and $SmartClear -ne 'Off') {
        Write-Error "-ClearFirst and -SmartClear cannot be used together. Use -SmartClear Off to disable auto-clear, or omit -ClearFirst for scheduled tasks."
        exit 1
    }
    if (-not $ScanPath -and -not $FromFile) {
        Write-Error "You must specify either -ScanPath or -FromFile."
        exit 1
    }
    if ($ScanPath -and $FromFile) {
        Write-Warning "-FromFile overrides -ScanPath. Using file: $FromFile"
        $ScanPath = $null
    }

    # Parse and validate -AlwaysInclude pipe-separated paths (wildcards not permitted)
    $AlwaysIncludePaths = @()
    if ($AlwaysInclude) {
        $AlwaysIncludePaths = $AlwaysInclude -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        foreach ($p in $AlwaysIncludePaths) {
            if ($p -match '[*?]') {
                Write-Error "-AlwaysInclude path '$p' contains a wildcard ('*' or '?'). Only exact paths are allowed."
                exit 1
            }
        }
        Write-Host "  AlwaysInclude | $($AlwaysIncludePaths.Count) fixed path(s) configured" -ForegroundColor DarkCyan
    }
#endregion

#region ----- Functions ----

    function Write-Banner {
        param([string]$Title, [string]$Color = "Cyan")
        $line = "=" * ($Title.Length + 4)
        Write-Host "`n$line" -ForegroundColor $Color
        Write-Host "  $Title  "  -ForegroundColor $Color
        Write-Host "$line`n"     -ForegroundColor $Color
    }

    function Import-ClientToolModule {
        if (Get-Module -Name ClientTool) {
            Write-Host "  Ready    | ClientTool module already loaded" -ForegroundColor Green
            return
        }

        # Try each candidate path in order
        $found = $null
        foreach ($candidate in $ModuleCandidates) {
            $resolved = Resolve-Path $candidate -ErrorAction SilentlyContinue
            if ($resolved) { $found = $resolved.Path; break }
        }

        # Also accept if already available on PSModulePath (no .psd1 path needed)
        if (-not $found -and (Get-Module -ListAvailable -Name ClientTool)) {
            $found = 'ClientTool'  # Import by name from module path
        }

        if (-not $found) {
            Write-Error "ClientTool module not found. Searched:"
            $ModuleCandidates | ForEach-Object { Write-Error "  $_" }
            Write-Error "Install from: https://github.com/BackupNerd/ClientTool"
            exit 1
        }

        Import-Module $found -Force -ErrorAction Stop
        Write-Host "  Loaded   | ClientTool module v$((Get-Module ClientTool).Version)" -ForegroundColor Green
    }

    function Test-NtfsLastAccessTracking {
        # NTFS Last Access Time (LAT) Tracking — Background
        # ────────────────────────────────────────────────────────────────────────
        # Windows maintains three timestamps per file on NTFS volumes:
        #   CreationTime   — when the file was first created on this volume
        #   LastWriteTime  — when the file content was last saved or changed
        #   LastAccessTime — when the file was last opened, read, or executed
        #
        # The LastAccessTime stamp requires a disk write on every file read,
        # which historically caused significant I/O overhead.  Microsoft began
        # disabling it by default in Windows Vista/2008 and made that the
        # permanent system default starting in Windows 8 / Server 2012.
        #
        # fsutil behavior query disablelastaccess returns one of these values:
        #   0 = User explicitly ENABLED  — tracking is ON  (reliable, stamp updates on reads)
        #   1 = User explicitly DISABLED — tracking is OFF (stamp is frozen)
        #   2 = System default DISABLED  — tracking is OFF (Windows 8/10/11/Server 2012+ default)
        #   3 = System + user DISABLED   — tracking is OFF
        #
        # If tracking is OFF and -FilterBy Access is used, the LastAccessTime
        # values on most files will be stale — often matching LastWriteTime or
        # the date the OS last had tracking enabled — so the filter will NOT
        # correctly identify recently opened files.
        #
        # To check:   fsutil behavior query disablelastaccess
        # To enable:  fsutil behavior set disablelastaccess 0  (admin, no reboot needed)
        # ────────────────────────────────────────────────────────────────────────
        try {
            $raw   = (& fsutil behavior query disablelastaccess 2>&1) | Out-String
            $value = if ($raw -match '=\s*(\d)') { [int]$Matches[1] } else { $null }
        }
        catch {
            return [PSCustomObject]@{ Value = $null; Enabled = $null; Raw = "fsutil unavailable: $_" }
        }

        return [PSCustomObject]@{
            Value   = $value
            Enabled = ($value -eq 0)   # Only 0 means tracking is ON
            Raw     = $raw.Trim()
        }
    }

    function Get-RecentlyAccessedFiles {
        param(
            [string]$Path,
            [datetime]$Since,
            [string]$Property      # FileInfo property: LastAccessTime, LastWriteTime, or CreationTime
        )

        if (-not (Test-Path $Path -PathType Container)) {
            Write-Error "Scan path not found or is not a directory: $Path"
            return $null
        }

        Write-Host "  Scanning | $Path" -ForegroundColor Cyan
        Write-Host "  Filter   | $Property >= $($Since.ToString('yyyy-MM-dd HH:mm:ss'))  (last $Hours hr)" -ForegroundColor Cyan

        $files = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.$Property -ge $Since }

        return $files
    }

    function Get-PathsFromFile {
        param([string]$FilePath)

        if (-not (Test-Path $FilePath -PathType Leaf)) {
            Write-Error "Selection list file not found: $FilePath"
            return $null
        }

        $paths = Get-Content $FilePath -ErrorAction Stop |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_ -PathType Leaf) }

        Write-Host "  Loaded   | $($paths.Count) valid path(s) from: $FilePath" -ForegroundColor Green
        return $paths
    }

    function Show-FileGridView {
        param(
            [string[]]$Paths,
            [string]$FilterProperty    # Which FileInfo property was used to filter (for column star-marking)
        )

        # Prepend [*] to the column that was the filter criterion so it stands out in the grid
        $writeLabel   = if ($FilterProperty -eq 'LastWriteTime')  { '[*] LastModified'  } else { 'LastModified'  }
        $accessLabel  = if ($FilterProperty -eq 'LastAccessTime') { '[*] LastAccessed'  } else { 'LastAccessed'  }
        $createdLabel = if ($FilterProperty -eq 'CreationTime')   { '[*] Created'       } else { 'Created'       }

        # Hydrate FileInfo objects to get rich metadata for display
        $displayItems = foreach ($path in $Paths) {
            $fi = Get-Item $path -ErrorAction SilentlyContinue
            if ($fi) {
                [PSCustomObject]@{
                    Name            = $fi.Name
                    'Size (KB)'     = [math]::Round($fi.Length / 1KB, 1)
                    $writeLabel     = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $accessLabel    = $fi.LastAccessTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $createdLabel   = $fi.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
                    Directory       = $fi.DirectoryName
                    FullPath        = $fi.FullName
                }
            }
        }

        Write-Host "  GridView | [*] marks the filter column  |  Ctrl+A selects all  |  OK = Confirm  |  X = Abort" -ForegroundColor Cyan

        $selected = $displayItems | Out-GridView `
            -Title "CoveBackupNow - $($displayItems.Count) file(s)  |  [*] filtered by: $FilterProperty  |  Ctrl+A = Select All  |  OK = Confirm  |  X = Abort" `
            -PassThru

        return $selected
    }

    function Write-SelectionListFile {
        param(
            [string[]]$Paths,
            [string]$OutputPath
        )

        $Paths | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "  Saved    | Selection list -> $OutputPath" -ForegroundColor Green
        Write-Host "  Files    | $($Paths.Count) path(s) written" -ForegroundColor Green
    }

    function Invoke-ClientToolFileSelections {
        param(
            [string[]]$Paths,
            [string]$SelectionPriority = $Priority
        )

        # Batch calls to avoid Windows CLI length limits.
        # Each path becomes: -include "path" (roughly path.Length + 11 chars)
        # With BatchSize=50 @ avg 80 chars/path = ~4,600 chars per call (well within 32KB limit)
        $totalBatches = [math]::Ceiling($Paths.Count / $BatchSize)
        $batchNum     = 0

        Write-Host "  Applying | $($Paths.Count) selection(s) in $totalBatches batch(es) of up to $BatchSize... [Priority: $SelectionPriority]" -ForegroundColor Cyan

        for ($i = 0; $i -lt $Paths.Count; $i += $BatchSize) {
            $batchNum++
            $end   = [math]::Min($i + $BatchSize - 1, $Paths.Count - 1)
            $batch = $Paths[$i..$end]

            Write-Host "  Batch    | $batchNum/$totalBatches ($($batch.Count) paths)" -ForegroundColor DarkCyan
            Set-ClientToolSelection -DataSource FileSystem -Include $batch -Priority $SelectionPriority
        }

        Write-Host "  Done     | All selections applied" -ForegroundColor Green
    }

#endregion

#region ----- Main Script ----

    # Optional: generate test data files so there is always something new to pick up
    if ($CreateTestData) {
        $testDir = Join-Path $ScanPath "TestData"
        if (-not (Test-Path $testDir)) { New-Item -ItemType Directory -Path $testDir -Force | Out-Null }
        1..2 | ForEach-Object {
            $fileName = "TestData_$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')_$([System.IO.Path]::GetRandomFileName().Replace('.','') ).txt"
            $filePath = Join-Path $testDir $fileName
            "Test data generated at $(Get-Date) -- run $_ of 2" | Out-File -FilePath $filePath -Encoding UTF8
            Write-Host "  TestData | Created: $filePath" -ForegroundColor DarkMagenta
            Start-Sleep -Milliseconds 10
        }
    }

    Write-Banner -Title "$ScriptName v$ScriptVersion - Backup Recently Accessed Files"

    # Step 1: Load ClientTool module
    Write-Host "[1/6] Loading ClientTool Module" -ForegroundColor White
    Import-ClientToolModule

    # Step 2: Collect file paths (scan or load from file)
    Write-Host "`n[2/6] Collecting File Paths" -ForegroundColor White

    if ($FromFile) {
        # Load from a previously generated selection list file
        $filePaths = Get-PathsFromFile -FilePath $FromFile
    } else {
        Write-Host "  Mode     | Filtering by: $FilterByLabel" -ForegroundColor Cyan

        # When -FilterBy Access is chosen, check whether NTFS LastAccessTime tracking is enabled.
        # On most Windows 8+ systems it is DISABLED by default, making this filter unreliable.
        if ($FilterBy -eq 'Access') {
            $latStatus = Test-NtfsLastAccessTracking

            if ($null -eq $latStatus.Enabled) {
                Write-Warning "Could not determine LastAccessTime tracking status (fsutil unavailable)."
                Write-Warning "Raw: $($latStatus.Raw)"
                Write-Warning "Results may be unreliable - consider using -FilterBy Write instead."
            }
            elseif (-not $latStatus.Enabled) {
                Write-Host ""
                Write-Host "  !! WARNING: LastAccessTime tracking is DISABLED on this system !!" -ForegroundColor Red
                Write-Host "  ---------------------------------------------------------------------" -ForegroundColor DarkRed
                Write-Host "  fsutil result : $($latStatus.Raw)"                                    -ForegroundColor DarkYellow
                Write-Host "  Value         : $($latStatus.Value)  (0=ON  1=user-OFF  2=system-OFF  3=both-OFF)" -ForegroundColor DarkYellow
                Write-Host ""
                Write-Host "  What this means:"                                                     -ForegroundColor Yellow
                Write-Host "  Windows 8 and later disable LastAccessTime updates by default to"     -ForegroundColor Yellow
                Write-Host "  reduce disk I/O. When disabled, LastAccessTime is NOT updated"        -ForegroundColor Yellow
                Write-Host "  when a file is opened or read - the stamp stays frozen at its"        -ForegroundColor Yellow
                Write-Host "  last known value, which may be months or years old."                 -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Effect on this script:"                                               -ForegroundColor Yellow
                Write-Host "  -FilterBy Access will likely return zero results or wrong files."    -ForegroundColor Yellow
                Write-Host "  Files you recently opened will NOT be detected."                     -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Recommended alternatives:"                                           -ForegroundColor Cyan
                Write-Host "  -FilterBy Write   -> catches recently saved/modified files (RECOMMENDED)" -ForegroundColor Cyan
                Write-Host "  -FilterBy Created -> catches newly created files"                     -ForegroundColor Cyan
                Write-Host ""
                Write-Host "  To enable LastAccessTime tracking (admin required, no reboot needed):" -ForegroundColor Cyan
                Write-Host "    fsutil behavior set disablelastaccess 0"                            -ForegroundColor Green
                Write-Host "  ---------------------------------------------------------------------" -ForegroundColor DarkRed
                Write-Host ""
                Write-Warning "Continuing with -FilterBy Access as requested - results are likely unreliable."
                Write-Host ""
            }
            else {
                Write-Host "  Tracking | LastAccessTime updates ENABLED (fsutil value=0) - filter is reliable" -ForegroundColor Green
            }
        }

        # Scan the path using the selected timestamp filter
        $recentFiles = Get-RecentlyAccessedFiles -Path $ScanPath -Since $CutoffTime -Property $FilterByProperty

        if (-not $recentFiles -or $recentFiles.Count -eq 0) {
            if ($AlwaysIncludePaths.Count -eq 0) {
                Write-Warning "No files found under '$ScanPath' where $FilterByProperty >= $($CutoffTime.ToString('yyyy-MM-dd HH:mm:ss'))"
                Write-Warning "Nothing to back up. Exiting."
                exit 0
            }
            Write-Warning "No recently changed files found - continuing with $($AlwaysIncludePaths.Count) AlwaysInclude path(s) only."
            $filePaths = @()
        } else {
            Write-Host "  Found    | $($recentFiles.Count) file(s) matched ($FilterByProperty in last $Hours hr)" -ForegroundColor Yellow
            $filePaths = $recentFiles | Select-Object -ExpandProperty FullName
        }
    }

    # Validate AlwaysInclude paths (warn if not found on disk, but still pass to ClientTool)
    if ($AlwaysIncludePaths.Count -gt 0) {
        foreach ($p in $AlwaysIncludePaths) {
            if (-not (Test-Path $p)) {
                Write-Warning "AlwaysInclude path not found on disk (still added to selection): $p"
            }
        }
        Write-Host "  AlwaysInclude | $($AlwaysIncludePaths.Count) path(s) will be applied at [$AlwaysIncludePriority] priority" -ForegroundColor Cyan
    }

    if ((-not $filePaths -or $filePaths.Count -eq 0) -and $AlwaysIncludePaths.Count -eq 0) {
        Write-Error "No valid file paths to process. Exiting."
        exit 1
    }

    # Step 3: GridView review — let user confirm/trim the file list
    Write-Host "`n[3/6] Review & Confirm Files" -ForegroundColor White

    if ($NoGridView) {
        Write-Host "  Skipping | GridView review (-NoGridView specified)" -ForegroundColor DarkYellow
    } else {
        $gridSelection = Show-FileGridView -Paths $filePaths -FilterProperty $FilterByProperty

        if (-not $gridSelection -or $gridSelection.Count -eq 0) {
            Write-Warning "No files confirmed in GridView. Backup cancelled."
            exit 0
        }

        $filePaths = $gridSelection | Select-Object -ExpandProperty FullPath
        Write-Host "  Confirmed| $($filePaths.Count) file(s) selected for backup" -ForegroundColor Green
    }

    # Step 4: Write selection list file (for re-use / audit)
    Write-Host "`n[4/6] Writing Selection List File" -ForegroundColor White
    if ($PSCmdlet.ShouldProcess($SelectionFile, "Write selection list")) {
        $allPathsForLog = @($filePaths) + @($AlwaysIncludePaths) | Select-Object -Unique
        Write-SelectionListFile -Paths $allPathsForLog -OutputPath $SelectionFile
    }

    # Step 5: Apply selections to ClientTool
    Write-Host "`n[5/6] Applying FileSystem Selections" -ForegroundColor White

    if ($ClearFirst) {
        Write-Host "  Clearing | Existing FileSystem selections (-ClearFirst specified)..." -ForegroundColor Cyan
        if ($PSCmdlet.ShouldProcess("FileSystem", "Clear existing selections")) {
            Clear-ClientToolSelection -DataSource FileSystem -Confirm:$false
            # Stamp the state file so -SmartClear runs later the same day know today's clear already happened
            [System.IO.File]::WriteAllText($SmartClearStateFile, (Get-Date).ToString('yyyy-MM-dd'))
        }
    } elseif ($SmartClear -ne 'Off') {
        # Auto-clear on schedule: Daily / Weekly (Sunday reset) / Monthly (1st of month)
        # Last-clear date persisted in CoveBackupNow_LastClear.dat (yyyy-MM-dd)
        $today         = Get-Date
        $todayStr      = $today.ToString('yyyy-MM-dd')
        $lastClearDate = if (Test-Path $SmartClearStateFile) { (Get-Content $SmartClearStateFile -Raw).Trim() } else { '' }

        $shouldClear = switch ($SmartClear) {
            'Daily' {
                # Clear on first run of each new calendar day
                $lastClearDate -ne $todayStr
            }
            'Weekly' {
                # Clear on first run of each new week — week boundary is Sunday
                $daysSinceSunday = [int]$today.DayOfWeek   # 0=Sun, 1=Mon ... 6=Sat
                $thisSunday      = $today.AddDays(-$daysSinceSunday).ToString('yyyy-MM-dd')
                $lastClearSunday = if ($lastClearDate) {
                    $lcd = [datetime]::ParseExact($lastClearDate,'yyyy-MM-dd',$null)
                    $lcd.AddDays(-[int]$lcd.DayOfWeek).ToString('yyyy-MM-dd')
                } else { '' }
                $lastClearSunday -ne $thisSunday
            }
            'Monthly' {
                # Clear on first run of each calendar month
                $thisMonth      = $today.ToString('yyyy-MM')
                $lastClearMonth = if ($lastClearDate.Length -ge 7) { $lastClearDate.Substring(0,7) } else { '' }
                $lastClearMonth -ne $thisMonth
            }
        }

        if ($shouldClear) {
            $scheduleLabel = switch ($SmartClear) {
                'Daily'   { 'calendar day' }
                'Weekly'  { 'week (Sunday reset)' }
                'Monthly' { 'month (1st reset)' }
            }
            Write-Host "  SmartClear | First run of new $scheduleLabel (last clear: $(if ($lastClearDate) { $lastClearDate } else { 'never' })) - clearing selections..." -ForegroundColor Cyan
            if ($PSCmdlet.ShouldProcess("FileSystem", "Clear existing selections (SmartClear - $scheduleLabel)")) {
                Clear-ClientToolSelection -DataSource FileSystem -Confirm:$false
                [System.IO.File]::WriteAllText($SmartClearStateFile, $todayStr)
            }
        } else {
            $accumLabel = switch ($SmartClear) {
                'Daily'   { "already cleared today ($todayStr)" }
                'Weekly'  { 'already cleared this week' }
                'Monthly' { 'already cleared this month' }
            }
            Write-Host "  SmartClear | $accumLabel - accumulating selections" -ForegroundColor DarkCyan
        }
    } else {
        Write-Host "  Accumulating | Adding to existing FileSystem selections (use -SmartClear Daily/Weekly/Monthly or -ClearFirst to reset)" -ForegroundColor DarkCyan
    }

    if ($filePaths -and $filePaths.Count -gt 0) {
        Write-Host "  Scan paths  | $($filePaths.Count) path(s) at [$Priority] priority" -ForegroundColor Cyan
        Invoke-ClientToolFileSelections -Paths $filePaths -SelectionPriority $Priority
    }

    if ($AlwaysIncludePaths.Count -gt 0) {
        Write-Host "  AlwaysInclude | $($AlwaysIncludePaths.Count) path(s) at [$AlwaysIncludePriority] priority" -ForegroundColor Cyan
        Invoke-ClientToolFileSelections -Paths $AlwaysIncludePaths -SelectionPriority $AlwaysIncludePriority
    }

    # Step 6: Trigger backup
    Write-Host "`n[6/6] Starting Backup" -ForegroundColor White

    if ($NoBackup) {
        Write-Warning "Backup skipped (-NoBackup specified). Selections are applied and ready."
        Write-Host "  Run manually:  Start-ClientToolBackup -DataSource FileSystem -NonInteractive" -ForegroundColor DarkCyan
    } else {
        Write-Host "  Starting | FileSystem backup job..." -ForegroundColor Cyan
        if ($PSCmdlet.ShouldProcess("FileSystem", "Start-ClientToolBackup")) {
            Start-ClientToolBackup -DataSource FileSystem -NonInteractive
        }
        Write-Host "  Submitted| Backup job dispatched to Backup Manager" -ForegroundColor Green
    }

    # Summary
    Write-Banner -Title "Complete - $($filePaths.Count) file(s) selected  |  Selection list: $(Split-Path $SelectionFile -Leaf)" -Color Green

#endregion
