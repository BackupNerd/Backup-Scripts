<# ----- About: ----
    # Cove Data Protection | Offline Backup Session Restore
    # Revision v12.0 - 2026-06-24
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
    # Logs restore start/finish details to a per-datasource log file in RestoreBase
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
    4. Initiates the restore process for each selected session.
    5. Logs the restore process details to a log file.

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
    Version: 12.0 - 2026-06-24
#>

param (
        [string[]]$RestoreSelection = @("C:\temp","C:\output"),                                     ## Path(s) to restore; multiple paths each get their own -selection argument
        [string]$RestoreBase = "X:\test restores\",                                        ## Base directory where the restored files will be placed 
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

$counter = 1

# Process each selected backup session
Foreach ($Session in $BackupSession) {
    $formattedStart = $Session.START.ToString("yyyy-MM-dd HH:mm:ss")        
    Write-Host "`nNumber of sessions selected: $($BackupSession.Count)"
    Write-Host "`nProcessing session $counter of $($BackupSession.Count): from $formattedStart"
    $status = & $clienttool control.status.get
     Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Restore to Start"

    # Update the title of the PowerShell window
    $host.ui.RawUI.WindowTitle = "Backup Manager restoring session $counter of $($BackupSession.Count): from $formattedStart"

    $counter++

    # Check the status of the Backup Manager and wait for it to be idle

    do {
        Start-Sleep -Seconds 5
        $status = & $clienttool control.status.get
    } until ($status -eq "Idle")
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Idle to Start Restore"

    if ($combinedrestore) {
        $RestorePath = Join-Path -Path $RestoreBase -ChildPath $DataSource
    }
    else {
        $RestorePath = Join-Path -Path $RestoreBase -ChildPath "$DataSource\$($formattedStart.Replace(':', '-').Replace(' ', '_'))"
    }

    # Create the restore directory if it does not exist
    if (-Not (Test-Path -Path $RestorePath)) {
        New-Item -ItemType Directory -Path $RestorePath -Force | Out-Null
    }

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
        #$sessionSuffix = $Session.START.ToString("yyyy-MM-dd_HH-mm-ss")
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

    # Wait for the restore process to start
    $status = & $clienttool control.status.get
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Restore to Start"
    $timeout = [datetime]::Now.AddSeconds(30)
    do {
        Start-Sleep -Seconds 2
        $status = & $clienttool control.status.get
    } until ($status -eq "Restore" -or ($status -eq "Idle" -and [datetime]::Now -gt $timeout))
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Restore to Complete"


    # Log the restore process details
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $DataSource Restore Started for Session  | $formattedStart | $RestoreSelection | $RestorePath "
    Add-Content -Path $logFile -Value $logEntry

    # Wait for the restore process to complete
    $status = & $clienttool control.status.get
    do {
        $status = & $clienttool control.status.get
        Start-Sleep -Seconds 5
    } while ($status -eq "Restore")
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Restore Complete"


    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $DataSource Restore Finished for Session | $formattedStart | $RestoreSelection | $RestorePath "
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

