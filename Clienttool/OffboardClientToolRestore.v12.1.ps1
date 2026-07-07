<# ----- About: ----
    # Cove Data Protection | Offline Backup Session Restore
    # Revision v12.1 - 2026-07-07
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email: eric.harless@n-able.com
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with the standalone edition of N-able | Cove Data Protection
    # Requires Backup Manager installed locally with ClientTool.exe present
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Validates availability of the target restore volume and base directory
    # Retrieves completed backup sessions from the local Backup Manager via ClientTool.exe
    # Displays sessions in a GridView for interactive selection (sorted oldest to newest)
    # Supports All | Daily | Weekly | Monthly session filtering
    # Supports multiple -RestoreSelection paths (each emits a separate -selection argument)
    # Initiates restore via ClientTool.exe control.restore.start for each selected session
    # Monitors restore progress using two independent signals:
    #   - ClientTool.exe control.status.get  (live engine state: Idle / Backup / Restore)
    #   - SessionReport.xml session records  (persistent; written at session end regardless of duration)
    # Detects sub-second/fast restores that would be missed by status polling alone
    # Skips sessions that fail to start; never hangs on a failed session
    # Logs restore start/finish details to a per-datasource log file in RestoreBase
    #
    # State monitoring timeouts:
    #   - Idle wait before restore:      120 minutes  (ClientTool poll every 5s)
    #   - Restore start confirmation:     60 seconds  (SessionReport.xml + ClientTool poll every 1s)
    #   - Restore completion wait:        Unbounded   (SessionReport.xml poll every 5s; exits on session record)
    #
    # Parameters:
    #   -RestoreSelection  One or more source paths to restore (string[])
    #   -RestoreBase       Base directory for restored output
    #   -CombinedRestore   Combine all session restores into one directory (default: true)
    #   -ExistingFileRestorePolicy   Overwrite | Skip (default: Overwrite)
    #   -OutdatedFileRestorePolicy   CheckContentOfOutdatedFilesOnly | CheckContentOfAllFiles
    #   -DataSource        FileSystem | NetworkShares | VssMsSql (default: FileSystem)
    #   -SessionType       All | Daily | Weekly | Monthly (default: All)
    #   -Weekday           Day filter for Weekly mode (default: Friday)
    #   -IncludedStates    Session states to include (default: Completed, CompletedWithErrors)
    #   -AddSessionTimestampSuffix   Deprecated suffix flag
    #
    # Examples:
    #   .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\"
    #   .\offline.ps1 -RestoreSelection "C:\Data","C:\Logs" -RestoreBase "D:\RestoredData\" -DataSource FileSystem
    #   .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -SessionType Monthly
    #   .\offline.ps1 -RestoreSelection "/" -RestoreBase "D:\RestoredData\" -DataSource VssMsSql -SessionType All
    #   .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -SessionType Weekly -Weekday Friday
# -----------------------------------------------------------#>  ## Behavior

<#
.SYNOPSIS
    Automates restoration of backup sessions using the local Cove | Backup Manager | ClientTool.exe command-line utility.

.DESCRIPTION
    The script performs the following tasks:
    1. Validates the availability of the target restore volume and base directory.
    2. Retrieves a list of completed backup sessions from the Backup Manager.
    3. Prompts the user to select one or more backup sessions to restore (sorted oldest to newest).
    4. Initiates the restore process for each selected session using two independent monitoring signals:
       - ClientTool.exe control.status.get for live engine state (Idle / Backup / Restore)
       - SessionReport.xml session records for persistent completion detection, including sub-second restores
    5. Skips sessions that fail to start or time out; never hangs on a failed session.
    6. Logs the restore process details to a log file.

.PARAMETER RestoreSelection
    One or more paths to restore. Each path is passed as a separate -selection argument to ClientTool.exe.
    Example: -RestoreSelection "C:\Data"  or  -RestoreSelection "C:\Data","C:\Logs"

.PARAMETER RestoreBase
    Base directory where the restored files will be placed.

.PARAMETER CombinedRestore
    If $true (default), combines all session restores into a single directory under RestoreBase\DataSource.
    If $false, creates a timestamped subdirectory per session.

.PARAMETER ExistingFileRestorePolicy
    How to handle files that already exist at the restore destination. Valid values: Overwrite, Skip.

.PARAMETER OutdatedFileRestorePolicy
    How to handle outdated files during restore. Valid values: CheckContentOfOutdatedFilesOnly, CheckContentOfAllFiles.

.PARAMETER DataSource
    Data source type to restore. Valid values: FileSystem, NetworkShares, VssMsSql.

.PARAMETER SessionType
    Filters which sessions appear in the GridView selector. Valid values: All, Daily, Weekly, Monthly.

.PARAMETER Weekday
    Day-of-week filter used when SessionType is Weekly. Valid values: Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday.

.PARAMETER IncludedStates
    Backup session states to include. Valid values: Completed, CompletedWithErrors, Aborted, Failed.

.PARAMETER AddSessionTimestampSuffix
    Deprecated. When set, appends a timestamp suffix to restored files.

.EXAMPLE
    .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\"
    Restores all qualifying sessions for C:\Data to D:\RestoredData\.

.EXAMPLE
    .\offline.ps1 -RestoreSelection "C:\Data","C:\Logs" -RestoreBase "D:\RestoredData\" -DataSource FileSystem
    Restores two selections per session using separate -selection arguments.

.EXAMPLE
    .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -SessionType Monthly
    Restores the last session of each month for C:\Data.

.EXAMPLE
    .\offline.ps1 -RestoreSelection "/" -RestoreBase "D:\RestoredData\" -DataSource VssMsSql -SessionType All
    Restores all VssMsSql sessions.

.EXAMPLE
    .\offline.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -SessionType Weekly -Weekday Friday
    Restores Friday sessions only.

.NOTES
    Author:  Eric Harless, Head Backup Nerd - N-able
    Twitter: @Backup_Nerd  Email: eric.harless@n-able.com
    Version: 12.1 - 2026-07-07
#>

param (
        [string[]]$RestoreSelection = @("C:\temp","C:\output"),                                     ## Path(s) to restore; multiple paths each get their own -selection argument
        [string]$RestoreBase = "d:\test restores\",                                        ## Base directory where the restored files will be placed 
        [switch]$CombinedRestore = $true,                                               ## ( if $true ) Combine multiple date restores into a single directory

        [ValidateSet("Overwrite", "Skip")]
        [string]$ExistingFileRestorePolicy = "Overwrite",                                     ## Restore policy for existing files
                
        [ValidateSet("CheckContentOfOutdatedFilesOnly", "CheckContentOfAllFiles")]
        [string]$OutdatedFileRestorePolicy = "CheckContentOfOutdatedFilesOnly",          ## Restore only outdated files
        
        [ValidateSet("FileSystem", "NetworkShares", "VssMsSql")]
        [string]$DataSource = "FileSystem",                                              ## Data source type to be restored

        [ValidateSet("All", "Daily", "Weekly", "Monthly")]
        [string]$SessionType = "All",                                                    ## Type of sessions to display be restored
        
        [ValidateSet("Sunday", "Monday", "Tuesday",
        "Wednesday", "Thursday", "Friday", "Saturday")][string]$Weekday = "Friday",      ## Day of the week for weekly restore sessions

        [ValidateSet("Completed", "CompletedWithErrors", "Aborted", "Failed")]
        [string[]]$IncludedStates = @("Completed", "CompletedWithErrors"),              ## Backup session states to include in the restore process

        [switch]$AddSessionTimestampSuffix = $false                                     ## (Depricated) Add session timestamp as suffix to restored files [format: file@yyyy-MM-dd_HH-mm-ss@.ext]
)


clear-host
Write-Output "Script Parameter Syntax:`n  $(Get-Command $PSCommandPath -Syntax)"
Push-Location -Path (Split-Path -Path $MyInvocation.MyCommand.Path)

# Output a summary of the defined script parameters
Write-Host "Script Parameters Summary:"
Write-Host "Restore Selection:            $($RestoreSelection -join ', ')"
Write-Host "Restore Base:                 $RestoreBase"
Write-Host "Combined Restore:             $CombinedRestore"
Write-Host "Existing File Restore Policy: $ExistingFileRestorePolicy"
Write-Host "Outdated File Restore Policy: $OutdatedFileRestorePolicy"
Write-Host "Data Source:                  $DataSource"
Write-Host "Session Type:                 $SessionType"
Write-Host "Weekday:                      $Weekday"
Write-Host "Included States:              $($IncludedStates -join ', ')"
Write-Host "Add Session Timestamp Suffix: $AddSessionTimestampSuffix"

$clienttool     = "C:\Program Files\Backup Manager\ClientTool.exe"              ## Path to the Backup Manager ClientTool executable
if ($RestoreBase.Length -lt 2) { Write-Warning "RestoreBase path is too short to extract a volume letter. Exiting."; exit }
$volumeLetter   = $RestoreBase.Substring(0, 2)                                  ## Extract the volume letter from the RestoreBase path
$logFile        = Join-Path -Path $RestoreBase -ChildPath "$($datasource)_Restore_log.txt"

# Check if the target restore volume is available
if (-Not (Test-Path -Path $volumeLetter)) {
        Write-Warning "The target restore volume $volumeLetter is not available. Exiting script."; exit
# Check if the RestoreBase directory exists, if not, create it
} elseif (-Not (Test-Path -Path $RestoreBase)) {
        New-Item -ItemType Directory -Path $RestoreBase -Force | Out-Null
}

# Retrieve a list of completed backup sessions from the Backup Manager
$BackupSessions = & $clienttool -machine-readable control.session.list -delimiter "`t" | ConvertFrom-Csv -Delimiter "`t" | Select-Object * | Where-Object { 

        ($IncludedStates -contains $_.STATE) -and          ## May exclude failed, skipped, aborted, completed and/or completed with errors sessions unless specified in $IncludedStates
        $_.TYPE -eq "Backup" -and 
        $_.DSRC -eq "$DataSource" -and 
        (($_.FLAGS -match "^A-") -or ($_.FLAGS -match "^--"))   ## Excludes cleaned sessions
} | ForEach-Object {
        $_.START = [datetime]::ParseExact($_.START, "yyyy-MM-dd HH:mm:ss", $null)
        $_.END = [datetime]::ParseExact($_.END, "yyyy-MM-dd HH:mm:ss", $null)
        $_ | Add-Member -MemberType NoteProperty -Name Weekday -Value $_.START.DayOfWeek -PassThru
}

# Prompt the user to select one or more backup sessions to restore
# Update the title of the PowerShell window
Write-Host "`nGetting Backup Sessions`nUse the Grid-View Popup and Select One or More Backup Sessions to Restore"
$host.ui.RawUI.WindowTitle = "Find the Grid-View Popup and Select One or More Backup Sessions to Restore"

# Ensure $BackupSession is always treated as an array

# This switch statement handles different types of backup session selections based on the value of $SessionType.
switch ($SessionType) {
        "All" {
            # If $SessionType is "All", display all backup sessions in an Out-GridView window for the user to select one or more sessions to restore.
            $BackupSession = @($BackupSessions | Sort-Object START | Out-GridView -Title "Select One or More Backup Sessions to Restore | Includes $($IncludedStates -join ', ')" -OutputMode Multiple)
        }
        "Daily" {
            # If $SessionType is "Daily", filter the backup sessions to only include the most recent session for each day.
            # Display the filtered sessions in an Out-GridView window for the user to select one or more sessions to restore.
            $BackupSession = @($BackupSessions | Sort-Object START -Descending | Group-Object { $_.START.ToString("yyyy-MM-dd") } | ForEach-Object { $_.Group | Select-Object -First 1 } | Sort-Object START | Out-GridView -Title "Select One or More EOD Backup Sessions to Restore | Includes $($IncludedStates -join ', ')" -OutputMode Multiple)
        }
        "Weekly" {
            # If $SessionType is "Weekly", filter the backup sessions to only include those that match the $Weekday script parameter
            # Display the filtered sessions in an Out-GridView window for the user to select one or more sessions to restore.
            $BackupSession = @($BackupSessions | Where-Object { $_.Weekday -eq $Weekday } | Sort-Object START | Out-GridView -Title "Select One or More [$Weekday] Backup Sessions to Restore | Includes $($IncludedStates -join ', ')" -OutputMode Multiple)
        }
        "Monthly" {
            # If $SessionType is "Monthly", sort the backup sessions by their start date in descending order.
            # Group the sessions by year and month, and select the first session from each group.
            # Display the grouped sessions in an Out-GridView window for the user to select one or more end-of-month (EOM) sessions to restore.
            $BackupSession = @($BackupSessions | Sort-Object START -Descending | Group-Object { $_.START.ToString("yyyy-MM") } | ForEach-Object { $_.Group | Select-Object -First 1 } | Sort-Object START | Out-GridView -Title "Select One or More EOM Backup Sessions to Restore | Includes $($IncludedStates -join ', ')" -OutputMode Multiple)
        }
}

$sessionReportPath = "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"
$DataSourcePlugin  = @{
    "FileSystem"    = "FsBackupPlugin"
    "NetworkShares" = "NetworkSharesBackupPlugin"
    "VssMsSql"      = "VssMsSqlBackupPlugin"
}
$pluginName = $DataSourcePlugin[$DataSource]

$counter = 1

# Process each selected backup session
Foreach ($Session in $BackupSession) {
    $formattedStart = $Session.START.ToString("yyyy-MM-dd HH:mm:ss")        
    Write-Host "`nNumber of sessions selected: $($BackupSession.Count)"
    Write-Host "`nProcessing session $counter of $($BackupSession.Count): from $formattedStart"
    $status = & $clienttool control.status.get
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Current status: $status"

    # Update the title of the PowerShell window
    $host.ui.RawUI.WindowTitle = "Backup Manager restoring session $counter of $($BackupSession.Count): from $formattedStart"

    $counter++

    if ($status -ne "Idle") {
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Status is $status — waiting up to 120 minutes for Idle"
    }

    # Wait for the Backup Manager to become idle before starting the restore.
    # StatusReport.xml is not reliably updated on Idle transitions — ClientTool is the authoritative source here.
    $idleTimeout = [datetime]::Now.AddMinutes(120)
    $idleTimedOut = $false
    do {
        Start-Sleep -Seconds 5
        $status = & $clienttool control.status.get
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Idle"
        if ([datetime]::Now -gt $idleTimeout) {
            Write-Warning "Timed out waiting for Idle state after 120 minutes (current status: $status). Skipping session $formattedStart."
            $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Idle wait timeout | $DataSource | $formattedStart | status: $status | Skipping session"
            Add-Content -Path $logFile -Value $logEntry
            $idleTimedOut = $true
            break
        }
    } until ($status -eq "Idle")
    if ($idleTimedOut) { continue }
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Idle confirmed — ready to start restore"

    if ($combinedrestore) {
        $RestorePath = Join-Path -Path $RestoreBase -ChildPath $DataSource
    }
    else {
        $RestorePath = Join-Path -Path $RestoreBase -ChildPath "$DataSource\$($formattedStart.Replace(':', '-').Replace(' ', '_'))"
    }

    # Create the restore directory if it does not exist
    if (-Not (Test-Path -Path $RestorePath)) {
        New-Item -ItemType Directory -Path $RestorePath -Force | Out-Null
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Created restore directory: $RestorePath"
    } else {
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Restore directory exists: $RestorePath"
    }

    # Snapshot the highest restore session Id before starting so we can identify the new session record
    try {
        [xml]$sr = Get-Content $sessionReportPath -Raw -ErrorAction Stop
        $priorSessions = $sr.SessionStatistics.Session | Where-Object { $_.Plugin -eq $pluginName -and $_.Type -eq "Restore" }
        $snapshotId = if ($priorSessions) { ($priorSessions | ForEach-Object { [int]$_.Id } | Measure-Object -Maximum).Maximum } else { 0 }
    } catch { $snapshotId = 0 }
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Session snapshot Id: $snapshotId (new restore will have Id > this)"

    # Start the restore process for the selected session
    $restoreArgs = @(
        "control.restore.start"
        "-datasource", $DataSource
        "-time", $formattedStart
    )
    foreach ($sel in $RestoreSelection) {
        $restoreArgs += @("-selection", $sel)
    }
    $restoreArgs += @(
        "-restore-to", $RestorePath
        "-existing-files-restore-policy", $ExistingFileRestorePolicy
        "-outdated-files-restore-policy", $OutdatedFileRestorePolicy
    )
    if ($AddSessionTimestampSuffix) {
        $sessionSuffix = $Session.START.ToString("yyyy-MM-dd_HH-mm-ss")
        $restoreArgs += "-add-suffix", $sessionSuffix
    }
    
    # Display the exact command line being executed
    $displayArgs = $restoreArgs | ForEach-Object { if ($_ -match '\s') { "'$_'" } else { $_ } }
    Write-Host "`n[COMMAND] & `"$clienttool`" $($displayArgs -join ' ')" -ForegroundColor Yellow
    
    & $clienttool @restoreArgs

    # Capture the exit code
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        Write-Warning "Error Starting $DataSource Restore for Session $formattedStart - Exit code: $exitCode - Skipping to next session"
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Error Starting $DataSource Restore for | $formattedStart | Exit code: $exitCode | Skipping session"
        Add-Content -Path $logFile -Value $logEntry
        continue
    }

    # Phase 1: Confirm restore started (up to 60 seconds)
    # Two independent signals checked on every tick — either one is sufficient to confirm the restore:
    #   1. SessionReport.xml: new session record with Id > snapshotId (catches sub-second/fast completes
    #      that no poll interval could ever see via status alone — record is persistent, written at session end)
    #   2. ClientTool control.status.get: StatusCode = Restore (catches long-running restores still in progress)
    # StatusReport.xml is NOT used here — it does not update reliably on status transitions.
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Waiting for restore to register (snapshot Id: $snapshotId)"
    $startTimeout = [datetime]::Now.AddSeconds(60)
    $newSession = $null
    $restoreActive = $false
    do {
        Start-Sleep -Seconds 1
        try {
            [xml]$sr = Get-Content $sessionReportPath -Raw -ErrorAction Stop
            $newSession = $sr.SessionStatistics.Session |
                Where-Object { $_.Plugin -eq $pluginName -and $_.Type -eq "Restore" -and [int]$_.Id -gt $snapshotId } |
                Sort-Object { [int]$_.Id } -Descending | Select-Object -First 1
        } catch { $newSession = $null }
        $status = & $clienttool control.status.get
        if ($newSession) {
            Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - SessionReport: new record (Id: $($newSession.Id)) | ClientTool: $status — fast complete detected"
            break
        }
        if ($status -eq "Restore") {
            Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - SessionReport: no record yet | ClientTool: $status — long-running restore confirmed active"
            $restoreActive = $true
            break
        }
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - SessionReport: no record | ClientTool: $status — waiting ($([int]($startTimeout - [datetime]::Now).TotalSeconds)s remaining)"
    } until ([datetime]::Now -gt $startTimeout)

    if (-not $newSession -and -not $restoreActive) {
        # Final check before skipping — session record may have arrived in the last poll window
        try {
            [xml]$sr = Get-Content $sessionReportPath -Raw -ErrorAction Stop
            $newSession = $sr.SessionStatistics.Session |
                Where-Object { $_.Plugin -eq $pluginName -and $_.Type -eq "Restore" -and [int]$_.Id -gt $snapshotId } |
                Sort-Object { [int]$_.Id } -Descending | Select-Object -First 1
        } catch { $newSession = $null }
    }

    if (-not $newSession -and -not $restoreActive) {
        Write-Warning "Restore did not register within 60 seconds (status: $status). Skipping session $formattedStart."
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Restore start timeout | $DataSource | $formattedStart | status: $status | Skipping session"
        Add-Content -Path $logFile -Value $logEntry
        continue
    }

    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $DataSource Restore Started for Session  | $formattedStart | $($RestoreSelection -join ';') | $RestorePath"
    Add-Content -Path $logFile -Value $logEntry

    # Phase 2: Wait for session record (only entered when restore is still running after Phase 1)
    if ($restoreActive) {
        Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Restore is running — polling session report for completion"
        do {
            Start-Sleep -Seconds 5
            try {
                [xml]$sr = Get-Content $sessionReportPath -Raw -ErrorAction Stop
                $newSession = $sr.SessionStatistics.Session |
                    Where-Object { $_.Plugin -eq $pluginName -and $_.Type -eq "Restore" -and [int]$_.Id -gt $snapshotId } |
                    Sort-Object { [int]$_.Id } -Descending | Select-Object -First 1
            } catch { $newSession = $null }
            if (-not $newSession) {
                $status = & $clienttool control.status.get
                Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ClientTool: $status — restore in progress, waiting for session record"
            }
        } until ($newSession)
    }

    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Restore complete | Status: $($newSession.Status) | Session Id: $($newSession.Id) | $($newSession.StartTimeUTC) → $($newSession.EndTimeUTC)"


    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $DataSource Restore Finished for Session | $formattedStart | $($RestoreSelection -join ';') | $RestorePath"
    Add-Content -Path $logFile -Value $logEntry
    Write-Host "`nLog file path: $logFile"

    # Wait for 60 seconds or until a key is pressed
    Write-Host "`nWaiting for 60 seconds or until a key is pressed..."
    for ($i = 60; $i -ge 0; $i--) {
        if ($Host.UI.RawUI.KeyAvailable) {
            Write-Host "`nKey pressed. Exiting wait."
            break
        }
        Write-Host "$i seconds remaining..."
        Start-Sleep -Seconds 1
    }
}

Write-host "`nRestore process completed. Exiting script."

