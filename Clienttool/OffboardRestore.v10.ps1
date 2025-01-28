<#
.SYNOPSIS
        This script automates the restoration of backup sessions using the Cove | Backup Manager | Clienttool.exe command-line utility.

.DESCRIPTION
        The script performs the following tasks:
        1. Validates the availability of the target restore volume and base directory.
        2. Retrieves a list of completed backup sessions from the Backup Manager.
        3. Prompts the user to select one or more backup sessions to restore.
        4. Initiates the restore process for each selected session.
        5. Logs the restore process details to a log file.

.PARAMETER RestoreSelection
        Specifies the path to the directory or file selection to be restored. Default is "C:\Dell".

.PARAMETER RestoreBase
        Specifies the base directory where the restored files will be placed. Default is "E:\Restore\".

.NOTES
        Author: Eric Harless
        Version: 10.0
        Date: 2025.01.28

.EXAMPLE
        .\OffboardRestore.v10.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\"
        This command restores the backup sessions for the selection "C:\Data" to the base directory "D:\RestoredData\".

        .\OffboardRestore.v10.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -DataSource "FileSystem"
        This command restores the backup sessions for the selection "C:\Data" from the data source "FileSystem" to the base directory "D:\RestoredData\".

        .\OffboardRestore.v10.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -SessionType "Monthly"
        This command restores the monthly backup sessions for the selection "C:\Data" to the base directory "D:\RestoredData\".

        .\OffboardRestore.v10.ps1 -RestoreSelection "/" -RestoreBase "D:\RestoredData\" -DataSource "VssMsSql" -SessionType "All"
        This command restores all backup sessions for the selection "/" from the data source "VssMsSql" to the base directory "D:\RestoredData\".

        .\OffboardRestore.v10.ps1 -RestoreSelection "C:\Data" -RestoreBase "D:\RestoredData\" -DataSource "FileSystem" -SessionType "Weekly" -Weekday "Friday"
        This command restores the weekly backup sessions for the selection "C:\Data" from the data source "FileSystem" on Fridays to the base directory "D:\RestoredData\".

#>

param (
        [string]$RestoreSelection = "c:\dell",                                          # Path to the directory or file selection to be restored
        [string]$RestoreBase = "E:\Restore\",                                           # Base directory where the restored files will be placed        
        
        [ValidateSet("FileSystem", "NetworkShares","VssMsSql")]
        [string]$DataSource = "FileSystem",                                             # Data source type to be restored

        [ValidateSet("All", "Weekly", "Monthly")][string]$SessionType = "Weekly",       # Type of sessions to be restored
        
        [ValidateSet("Sunday", "Monday", "Tuesday",
        "Wednesday", "Thursday", "Friday", "Saturday")][string]$Weekday = "Friday"      # Day of the week for weekly restore sessions

)

$clienttool     = "C:\Program Files\Backup Manager\ClientTool.exe"  # Path to the Backup Manager ClientTool executable
$volumeLetter   = $RestoreBase.Substring(0, 2)  # Extract the volume letter from the RestoreBase path
$logFile        = Join-Path -Path $RestoreBase -ChildPath "restore_log.txt"

# Check if the target restore volume is available
if (-Not (Test-Path -Path $volumeLetter)) {
        Write-Warning "The target restore volume $volumeLetter is not available. Exiting script."; exit
# Check if the RestoreBase directory exists, if not, create it
} elseif (-Not (Test-Path -Path $RestoreBase)) {
        New-Item -ItemType Directory -Path $RestoreBase -Force | Out-Null
}

# Retrieve a list of completed backup sessions from the Backup Manager
$BackupSessions = & $clienttool -machine-readable control.session.list -delimiter "`t" | ConvertFrom-Csv -Delimiter "`t" | Select-Object * | Where-Object { 
        $_.STATE -eq "Completed" -and 
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
Write-Host "Getting Backup Sessions`nUse the Grid-View Popup and Select One or More Backup Sessions to Restore"
$host.ui.RawUI.WindowTitle = "Find the Grid-View Popup and Select One or More Backup Sessions to Restore"

# This switch statement handles different types of backup session selections based on the value of $SessionType.
switch ($SessionType) {
        "All" {
                # If $SessionType is "All", display all backup sessions in an Out-GridView window for the user to select one or more sessions to restore.
                $BackupSession = $BackupSessions | Out-GridView -Title "Select One or More Backup Sessions to Restore" -OutputMode Multiple
        }
        "Weekly" {
                # If $SessionType is "Weekly", filter the backup sessions to only include those that match the $Weekday script parameter
                # Display the filtered sessions in an Out-GridView window for the user to select one or more sessions to restore.
                $BackupSession = $BackupSessions | Where-Object { $_.Weekday -eq $Weekday } | Out-GridView -Title "Select One or More [$Weekday] Backup Sessions to Restore" -OutputMode Multiple
        }
        "Monthly" {
                # If $SessionType is "Monthly", sort the backup sessions by their start date in descending order.
                # Group the sessions by year and month, and select the first session from each group.
                # Display the grouped sessions in an Out-GridView window for the user to select one or more end-of-month (EOM) sessions to restore.
                $BackupSession = $BackupSessions | Sort-Object START -Descending | Group-Object { $_.START.ToString("yyyy-MM") } | ForEach-Object { $_.Group | Select-Object -First 1 } | Out-GridView -Title "Select One or More Backup EOM Sessions to Restore" -OutputMode Multiple
        }
}

$counter = 1

# Process each selected backup session
Foreach ($Session in $BackupSession) {
    $formattedStart = $Session.START.ToString("yyyy-MM-dd HH:mm:ss")        
    Write-Host "`nNumber of sessions selected: $($BackupSession.Count)"
    Write-Host "`nProcessing session $counter of $($BackupSession.Count): from $formattedStart"

    # Update the title of the PowerShell window
    $host.ui.RawUI.WindowTitle = "Backup Manager restoring session $counter of $($BackupSession.Count): from $formattedStart"

    $counter++

    # Check the status of the Backup Manager and wait for it to be idle

    do {
        Start-Sleep -Seconds 5
        $status = & $clienttool control.status.get
    } until ($status -eq "Idle")
    Write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $status - Waiting for Idle to Start Restore"

    $RestorePath = $RestoreBase + "\" + $DataSource + "\" + $formattedStart.Replace(":", "-").Replace(" ", "_") 

    # Create the restore directory if it does not exist
    if (-Not (Test-Path -Path $RestorePath)) {
        New-Item -ItemType Directory -Path $RestorePath -Force | Out-Null
    }

    # Start the restore process for the selected session
   & $clienttool control.restore.start -datasource $DataSource -time $formattedStart -selection "$RestoreSelection" -restore-to $RestorePath -existing-files-restore-policy Overwrite

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
    do {
        Start-Sleep -Seconds 2
        $status = & $clienttool control.status.get
    } until ($status -eq "Restore")
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

