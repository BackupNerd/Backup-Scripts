<# ----- About: ----
    # ConnectWise Manage - Interrogate Ticket Configuration Options
    # Revision v10 - 2026-01-08
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
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
    # For use with N-able | Cove Data Protection and ConnectWise Manage
    # Requires ConnectWiseManageAPI PowerShell module
    # Credentials are stored using Windows DPAPI encryption and can only be 
    # decrypted by the same user account on the same machine where created.
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Load stored ConnectWise Manage API credentials (DPAPI encrypted)
    # Connect to ConnectWise Manage API
    # Discover and enumerate:
    #   - Service Boards (active boards only)
    #   - Service Statuses (grouped by board)
    #   - Service Priorities (with color codes)
    # Export discovered configuration to CSV files in Boards subfolder
    # Display recommendations for monitoring script configuration
    # Interactive mode: Allow selection of configuration via GridView
    # Auto-update Cove monitoring script with selected configuration
    # Create timestamped backup before modifying monitoring script
    # Undo mode: Restore monitoring script from previous backup
    #
    # Use the -AllowInsecureSSL parameter to bypass SSL validation (default=$true, staging/dev only)
    # Use the -Interactive parameter to enable interactive mode (default=$true, set to $false for discovery/export only with no prompts/changes)
    # Use the -MonitoringScriptPath parameter to specify monitoring script (auto-detects if omitted)
    # Use the -Undo parameter to restore monitoring script from backup (default=$false, interactive GridView selection)
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://github.com/christaylorcodes/ConnectWiseManageAPI
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$False)][bool]$AllowInsecureSSL = $true,     ## Bypass SSL certificate validation (for staging/dev)
    [Parameter(Mandatory=$False)][bool]$Interactive = $true,          ## Enable interactive GridView selection and script updates ($false = discovery/export only, no prompts, no changes)
    [Parameter(Mandatory=$False)][string]$MonitoringScriptPath = "",  ## Auto-detect latest version if not specified
    [Parameter(Mandatory=$False)][bool]$Undo = $false                 ## Restore monitoring script from backup (interactive GridView selection)
)

#Requires -Version 7.0

# PowerShell 7 version check with helpful error message
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "Download PowerShell 7: https://aka.ms/powershell" -ForegroundColor Cyan
    exit 1
}

# Auto-detect monitoring script if not specified
if (-not $MonitoringScriptPath) {
    $detectedScript = Get-ChildItem -Path $PSScriptRoot -Filter "Cove2CWM-SyncTickets.v*.ps1" | 
        Where-Object { $_.Name -notmatch 'backup' } | 
        Sort-Object Name -Descending | 
        Select-Object -First 1
    
    if ($detectedScript) {
        $MonitoringScriptPath = $detectedScript.FullName
    } else {
        Write-Host "ERROR: Could not find Cove2CWM-SyncTickets script in $PSScriptRoot" -ForegroundColor Red
        Write-Host "Specify path manually with -MonitoringScriptPath parameter" -ForegroundColor Yellow
        exit 1
    }
}

#region ----- Environment Setup ----

# Handle Undo operation first (before Clear-Host)
if ($Undo) {
    Clear-Host
    Write-Host "`n  ConnectWise Manage - Undo Configuration Changes`n" -ForegroundColor Cyan
    
    $backupFolder = Join-Path (Split-Path $MonitoringScriptPath) "Backups"
    
    if (Test-Path $backupFolder) {
        # Get all backup files
        $backupFiles = Get-ChildItem -Path $backupFolder -Filter "*.backup" | Sort-Object LastWriteTime -Descending
        
        if ($backupFiles.Count -gt 0) {
            # Parse backup files to get configuration details
            $backupList = @()
            foreach ($file in $backupFiles) {
                $content = Get-Content $file.FullName -Raw
                $timestamp = $file.LastWriteTime
                
                $board = if ($content -match '\$TicketBoard\s*=\s*"([^"]*)"') { $matches[1] } else { "N/A" }
                $newStatus = if ($content -match '\$TicketStatus\s*=\s*"([^"]*)"') { $matches[1] } else { "N/A" }
                $closedStatus = if ($content -match '\$TicketClosedStatus\s*=\s*"([^"]*)"') { $matches[1] } else { "N/A" }
                $priority = if ($content -match '\$TicketPriority\s*=\s*"([^"]*)"') { $matches[1] } else { "N/A" }
                
                $backupList += [PSCustomObject]@{
                    Timestamp = $timestamp.ToString("yyyy-MM-dd HH:mm:ss")
                    Board = $board
                    NewStatus = $newStatus
                    ClosedStatus = $closedStatus
                    Priority = $priority
                    FilePath = $file.FullName
                    FileName = $file.Name
                }
            }
            
            # Show GridView to select backup to restore
            $selectedBackup = $backupList | Select-Object Timestamp, Board, NewStatus, ClosedStatus, Priority | 
                Out-GridView -Title "Select Configuration Backup to Restore (Most Recent at Top)" -OutputMode Single
            
            if ($selectedBackup) {
                $backupToRestore = $backupList | Where-Object { $_.Timestamp -eq $selectedBackup.Timestamp } | Select-Object -First 1
                
                if ($PSCmdlet.ShouldProcess($MonitoringScriptPath, "Restore from backup: $($backupToRestore.Timestamp)")) {
                    Copy-Item $backupToRestore.FilePath $MonitoringScriptPath -Force
                    Write-Host "  SUCCESS: Monitoring script restored from backup!" -ForegroundColor Green
                    Write-Host "  Backup Date: $($backupToRestore.Timestamp)" -ForegroundColor Gray
                    Write-Host "  File: $MonitoringScriptPath" -ForegroundColor Gray
                    
                    Write-Host "`n  Restored Configuration:" -ForegroundColor Cyan
                    Write-Host "  Board: $($backupToRestore.Board)" -ForegroundColor White
                    Write-Host "  New Status: $($backupToRestore.NewStatus)" -ForegroundColor White
                    Write-Host "  Closed Status: $($backupToRestore.ClosedStatus)" -ForegroundColor White
                    Write-Host "  Priority: $($backupToRestore.Priority)" -ForegroundColor White
                    Write-Host ""
                }
            } else {
                Write-Host "  Restore cancelled - no backup selected." -ForegroundColor Yellow
                Write-Host ""
            }
        } else {
            Write-Warning "No backup files found in: $backupFolder"
            Write-Host "  Cannot undo - no previous configurations saved." -ForegroundColor Yellow
            Write-Host "  Backups are created automatically when you apply configuration changes." -ForegroundColor Gray
            Write-Host ""
        }
    } else {
        Write-Warning "Backup folder not found at: $backupFolder"
        Write-Host "  Cannot undo - no previous configuration saved." -ForegroundColor Yellow
        Write-Host "  Backups are created automatically when you apply configuration changes." -ForegroundColor Gray
        Write-Host ""
    }
    return
}

Clear-Host
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "`n  ConnectWise Manage - Ticket Configuration Discovery`n" -ForegroundColor Cyan

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Bypass SSL certificate validation if requested
if ($AllowInsecureSSL) {
    Write-Warning "SSL certificate validation is disabled - use only for staging/dev environments!"
    try {
        Add-Type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;
            public class TrustAllCertsPolicy : ICertificatePolicy {
                public bool CheckValidationResult(
                    ServicePoint srvPoint, X509Certificate certificate,
                    WebRequest request, int certificateProblem) {
                    return true;
                }
            }
"@ -ErrorAction SilentlyContinue
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    } catch {
        # Type already loaded, ignore
    }
}

#endregion

#region ----- Install/Import Module ----

if (Get-Module -ListAvailable -Name "ConnectWiseManageAPI") {
    Write-Host "  ConnectWise Manage PowerShell Module Found" -ForegroundColor Green
} else {
    Write-Host "  Installing ConnectWise Manage PowerShell Module..." -ForegroundColor Yellow
    Install-Module -Name ConnectWiseManageAPI -Force -AllowClobber
    Write-Host "  Module Installed" -ForegroundColor Green
}

Import-Module ConnectWiseManageAPI -ErrorAction Stop

#endregion

#region ----- Get Credentials ----

$CWMAPICredsFile = "C:\ProgramData\MXB\${env:computername}_${env:username}_CWM_Ticketing_Credentials.Secure.xml"

if (Test-Path $CWMAPICredsFile) {
    Write-Host "`n  Loading stored ConnectWise credentials..." -ForegroundColor Cyan
    $CWMAPICreds = Import-Clixml -Path $CWMAPICredsFile
    
    # Decrypt sensitive fields
    $privateKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
            ($CWMAPICreds.privateKey | ConvertTo-SecureString)
        )
    )
    $pubKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
            ($CWMAPICreds.pubKey | ConvertTo-SecureString)
        )
    )
    $clientId = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
            ($CWMAPICreds.clientId | ConvertTo-SecureString)
        )
    )
    
    Write-Host "  Server: $($CWMAPICreds.Server)" -ForegroundColor White
    Write-Host "  Company: $($CWMAPICreds.Company)" -ForegroundColor White
} else {
    Write-Error "Credential file not found: $CWMAPICredsFile"
    Write-Host "Run the monitoring script first to create credentials." -ForegroundColor Yellow
    Exit
}

#endregion

#region ----- Connect to ConnectWise ----

Write-Host "`n  Connecting to ConnectWise Manage..." -ForegroundColor Cyan
Write-Host "  Server: $($CWMAPICreds.Server)" -ForegroundColor Gray
Write-Host "  Company: $($CWMAPICreds.Company)" -ForegroundColor Gray

$connectionParams = @{
    Server = $CWMAPICreds.Server
    Company = $CWMAPICreds.Company
    pubkey = $pubKey
    privatekey = $privateKey
    clientId = $clientId
}

try {
    $connectionResult = Connect-CWM @connectionParams -ErrorAction Stop
    
    # The module stores connection in $Global:CWMServerConnection
    $Script:CWMServerConnection = $Global:CWMServerConnection
    
    if ($Script:CWMServerConnection) {
        Write-Host "  Connected Successfully!" -ForegroundColor Green
        Write-Host "  Server: $($Script:CWMServerConnection.Server)" -ForegroundColor Gray
    } else {
        Write-Warning "Connection object not found - attempting to continue"
    }
}
catch {
    if (-not $AllowInsecureSSL) {
        Write-Error "Failed to connect to ConnectWise Manage: $_"
        Exit
    }
    # Continue anyway for staging/dev with SSL bypass
    Write-Warning "Connection warning (SSL bypass active): $_"
}

#endregion

#region ----- Discover Service Boards ----

Write-Host "`n" -NoNewline
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  SERVICE BOARDS" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan

try {
    if ($AllowInsecureSSL) {
        $boards = Get-CWMServiceBoard -all -ErrorAction SilentlyContinue 2>$null
    } else {
        $boards = Get-CWMServiceBoard -all
    }
    
    $boards | Select-Object id, Name, inactiveFlag | Sort-Object Name | ForEach-Object {
        $statusText = if ($_.inactiveFlag) { " (INACTIVE)" } else { "" }
        Write-Host "  ID: $($_.id.ToString().PadRight(5)) | Name: $($_.name)$statusText" -ForegroundColor $(if ($_.inactiveFlag) { "Gray" } else { "White" })
    }
    
    Write-Host "`n  Total Boards Found: $($boards.Count)" -ForegroundColor Yellow
    
} catch {
    Write-Warning "Error retrieving service boards: $_"
}

#endregion

#region ----- Discover Service Statuses ----

Write-Host "`n" -NoNewline
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  SERVICE STATUSES (by Board)" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan

foreach ($board in ($boards | Where-Object { -not $_.inactiveFlag } | Sort-Object Name)) {
    Write-Host "`n  Board: $($board.name) (ID: $($board.id))" -ForegroundColor Cyan
    Write-Host "  " -NoNewline
    Write-Host ("-" * 78) -ForegroundColor DarkGray
    
    try {
        if ($AllowInsecureSSL) {
            $statuses = Get-CWMBoardStatus -parentId $board.id -all -ErrorAction SilentlyContinue 2>$null
        } else {
            $statuses = Get-CWMBoardStatus -parentId $board.id -all
        }
        
        $statuses | Sort-Object sortOrder | ForEach-Object {
            $statusInfo = "  ID: $($_.id.ToString().PadRight(5)) | Name: $($_.name.PadRight(30))"
            
            if ($_.closedStatus) {
                $statusInfo += " [CLOSED]"
                Write-Host $statusInfo -ForegroundColor Green
            } elseif ($_.defaultFlag) {
                $statusInfo += " [DEFAULT]"
                Write-Host $statusInfo -ForegroundColor Yellow
            } else {
                Write-Host $statusInfo -ForegroundColor White
            }
        }
        
    } catch {
        Write-Warning "    Error retrieving statuses for board $($board.name): $_"
    }
}

#endregion

#region ----- Discover Service Priorities ----

Write-Host "`n" -NoNewline
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  SERVICE PRIORITIES" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan

try {
    if ($AllowInsecureSSL) {
        $priorities = Get-CWMPriority -all -ErrorAction SilentlyContinue 2>$null
    } else {
        $priorities = Get-CWMPriority -all
    }
    
    $priorities | Sort-Object sort | ForEach-Object {
        $priorityInfo = "  ID: $($_.id.ToString().PadRight(5)) | Name: $($_.name.PadRight(35))"
        
        if ($_.name -match "Critical|Priority 1") {
            Write-Host $priorityInfo -ForegroundColor Red
        } elseif ($_.name -match "High|Priority 2") {
            Write-Host $priorityInfo -ForegroundColor Yellow
        } else {
            Write-Host $priorityInfo -ForegroundColor White
        }
    }
    
    Write-Host "`n  Total Priorities Found: $($priorities.Count)" -ForegroundColor Yellow
    
} catch {
    Write-Warning "Error retrieving priorities: $_"
}

#endregion

#region ----- Discover Service Types ----

# Skip service types - not required for monitoring script configuration
# Service types require board-specific parentId and are rarely changed

#endregion

#region ----- Export Results ----

Write-Host "`n" -NoNewline
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  EXPORTING RESULTS" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan

# Create Boards subfolder for output files
$exportPath = Join-Path $PSScriptRoot "Boards"
if (-not (Test-Path $exportPath)) {
    New-Item -Path $exportPath -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export Boards
if ($boards) {
    $boardsFile = Join-Path $exportPath "CWM_Boards_$timestamp.csv"
    $boards | Select-Object id, Name, inactiveFlag | Export-Csv -Path $boardsFile -NoTypeInformation
    Write-Host "  Boards exported to: $boardsFile" -ForegroundColor Green
}

# Export all statuses with board info
if ($boards) {
    $allStatuses = @()
    foreach ($board in $boards) {
        try {
            if ($AllowInsecureSSL) {
                $statuses = Get-CWMBoardStatus -parentId $board.id -all -ErrorAction SilentlyContinue 2>$null
            } else {
                $statuses = Get-CWMBoardStatus -parentId $board.id -all
            }
            $statuses | ForEach-Object {
                $allStatuses += [PSCustomObject]@{
                    BoardId = $board.id
                    BoardName = $board.name
                    StatusId = $_.id
                    StatusName = $_.name
                    SortOrder = $_.sortOrder
                    DefaultFlag = $_.defaultFlag
                    ClosedStatus = $_.closedStatus
                }
            }
        } catch { }
    }
    
    if ($allStatuses.Count -gt 0) {
        $statusesFile = Join-Path $exportPath "CWM_Statuses_$timestamp.csv"
        $allStatuses | Sort-Object BoardName, SortOrder | Export-Csv -Path $statusesFile -NoTypeInformation
        Write-Host "  Statuses exported to: $statusesFile" -ForegroundColor Green
    }
}

# Export Priorities
if ($priorities) {
    $prioritiesFile = Join-Path $exportPath "CWM_Priorities_$timestamp.csv"
    $priorities | Select-Object id, Name, sort, color | Export-Csv -Path $prioritiesFile -NoTypeInformation
    Write-Host "  Priorities exported to: $prioritiesFile" -ForegroundColor Green
}

# Cleanup old CSV exports - keep only last 10 of each type
Write-Host "`n  Cleaning up old CSV exports..." -ForegroundColor Gray
$csvTypes = @('CWM_Boards_*.csv', 'CWM_Statuses_*.csv', 'CWM_Priorities_*.csv')
foreach ($pattern in $csvTypes) {
    $oldFiles = Get-ChildItem -Path $exportPath -Filter $pattern -ErrorAction SilentlyContinue | 
                Sort-Object LastWriteTime -Descending | 
                Select-Object -Skip 10
    
    if ($oldFiles) {
        $oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Host "    Removed $($oldFiles.Count) old $($pattern -replace '\*','')" -ForegroundColor DarkGray
    }
}

#endregion

#region ----- Summary Recommendations ----

Write-Host "`n" -NoNewline
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "  RECOMMENDATIONS FOR MONITORING SCRIPT" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan

# Initialize selection object
$Script:SelectedConfig = @{
    Board = $null
    NewStatus = $null
    ClosedStatus = $null
    PriorityServer = $null
    PriorityWorkstation = $null
    PriorityM365 = $null
}

# Find "Service Desk" board
$serviceDeskBoard = $boards | Where-Object { $_.name -eq "Service Desk" -and -not $_.inactiveFlag } | Select-Object -First 1

if ($serviceDeskBoard) {
    Write-Host "`n  Recommended Board: '$($serviceDeskBoard.name)' (ID: $($serviceDeskBoard.id))" -ForegroundColor Green
    
    # Get statuses for this board
    try {
        if ($AllowInsecureSSL) {
            $serviceDeskStatuses = Get-CWMBoardStatus -parentId $serviceDeskBoard.id -all -ErrorAction SilentlyContinue 2>$null
        } else {
            $serviceDeskStatuses = Get-CWMBoardStatus -parentId $serviceDeskBoard.id -all
        }
        
        $newStatus = $serviceDeskStatuses | Where-Object { $_.name -like "*New*" -and -not $_.closedStatus } | Select-Object -First 1
        $closedStatus = $serviceDeskStatuses | Where-Object { $_.closedStatus -eq $true } | Select-Object -First 1
        
        if ($newStatus) {
            Write-Host "  Recommended New Status: '$($newStatus.name)' (ID: $($newStatus.id))" -ForegroundColor Green
        }
        if ($closedStatus) {
            Write-Host "  Recommended Closed Status: '$($closedStatus.name)' (ID: $($closedStatus.id))" -ForegroundColor Green
        }
    } catch { }
}

# Find priority options
$priority1 = $priorities | Where-Object { $_.name -match "Priority 1|Critical" } | Select-Object -First 1
$priority2 = $priorities | Where-Object { $_.name -match "Priority 2|High" } | Select-Object -First 1
$priority3 = $priorities | Where-Object { $_.name -match "Priority 3|Normal" } | Select-Object -First 1

Write-Host "`n  Recommended Priority Mapping:" -ForegroundColor Yellow
if ($priority1) { Write-Host "    Critical Issues: '$($priority1.name)'" -ForegroundColor Red }
if ($priority2) { Write-Host "    Warning Issues:  '$($priority2.name)'" -ForegroundColor Yellow }
if ($priority3) { Write-Host "    Stale Backups:   '$($priority3.name)'" -ForegroundColor White }

Write-Host "`n  Script Parameter Suggestions:" -ForegroundColor Cyan
if ($serviceDeskBoard) {
    Write-Host "    -TicketBoard '$($serviceDeskBoard.name)'" -ForegroundColor Gray
}
if ($newStatus) {
    Write-Host "    -TicketStatus '$($newStatus.name)'" -ForegroundColor Gray
}
if ($closedStatus) {
    Write-Host "    -TicketClosedStatus '$($closedStatus.name)'" -ForegroundColor Gray
}
if ($priority3) {
    Write-Host "    -TicketPriority '$($priority3.name)'" -ForegroundColor Gray
}

#region ----- Interactive Selection ----

if ($Interactive) {
    Write-Host "`n" -NoNewline
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "  INTERACTIVE CONFIGURATION" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan
    
    # Read current values from monitoring script if it exists
    $currentBoard = $null
    $currentNewStatus = $null
    $currentClosedStatus = $null
    $currentPriorityServer = $null
    $currentPriorityWorkstation = $null
    $currentPriorityM365 = $null
    
    if (Test-Path $MonitoringScriptPath) {
        Write-Host "`n  Reading current configuration from monitoring script..." -ForegroundColor Cyan
        $scriptContent = Get-Content $MonitoringScriptPath -Raw
        
        if ($scriptContent -match '\$TicketBoard\s*=\s*"([^"]*)"') {
            $currentBoard = $matches[1]
            Write-Host "  Current Board: $currentBoard" -ForegroundColor Gray
        }
        if ($scriptContent -match '\$TicketStatus\s*=\s*"([^"]*)"') {
            $currentNewStatus = $matches[1]
            Write-Host "  Current New Status: $currentNewStatus" -ForegroundColor Gray
        }
        if ($scriptContent -match '\$TicketClosedStatus\s*=\s*"([^"]*)"') {
            $currentClosedStatus = $matches[1]
            Write-Host "  Current Closed Status: $currentClosedStatus" -ForegroundColor Gray
        }
        if ($scriptContent -match '\$TicketPriorityServer\s*=\s*"([^"]*)"') {
            $currentPriorityServer = $matches[1]
            Write-Host "  Current Server Priority: $currentPriorityServer" -ForegroundColor Gray
        }
        if ($scriptContent -match '\$TicketPriorityWorkstation\s*=\s*"([^"]*)"') {
            $currentPriorityWorkstation = $matches[1]
            Write-Host "  Current Workstation Priority: $currentPriorityWorkstation" -ForegroundColor Gray
        }
        if ($scriptContent -match '\$TicketPriorityM365\s*=\s*"([^"]*)"') {
            $currentPriorityM365 = $matches[1]
            Write-Host "  Current M365 Priority: $currentPriorityM365" -ForegroundColor Gray
        }
    }
    
    # Select Service Board
    Write-Host "`n  Step 1: Select Service Board" -ForegroundColor Yellow
    $selectedBoard = $boards | Sort-Object id | Select-Object @{N='Board Name';E={$_.name}}, 
        @{N='Status';E={if($_.inactiveFlag){'INACTIVE'}else{'Active'}}},
        @{N='Current';E={if($_.name -eq $currentBoard){'✓'}else{''}}},
        @{N='Recommended';E={if($_.name -eq $serviceDeskBoard.name){'★'}else{''}}},
        @{N='Board ID';E={$_.id}} | 
        Out-GridView -Title "Select Service Board | Current: $currentBoard | Recommended: $($serviceDeskBoard.name)" -OutputMode Single
    
    if ($selectedBoard) {
        $Script:SelectedConfig.Board = $selectedBoard.'Board Name'
        Write-Host "  Selected Board: $($selectedBoard.'Board Name')" -ForegroundColor Green
        
        # Get statuses for selected board
        if ($AllowInsecureSSL) {
            $boardStatuses = Get-CWMBoardStatus -parentId $selectedBoard.'Board ID' -all -ErrorAction SilentlyContinue 2>$null
        } else {
            $boardStatuses = Get-CWMBoardStatus -parentId $selectedBoard.'Board ID' -all
        }
        
        # Select New Ticket Status
        Write-Host "`n  Step 2: Select New Ticket Status" -ForegroundColor Yellow
        $openStatuses = $boardStatuses | Where-Object { -not $_.closedStatus } | Sort-Object sortOrder
        $recommendedNew = $openStatuses | Where-Object { $_.name -like "*New*" } | Select-Object -First 1
        
        $selectedNewStatus = $openStatuses | Select-Object @{N='Status Name';E={$_.name}},
            @{N='Current';E={if($_.name -eq $currentNewStatus){'✓'}else{''}}},
            @{N='Recommended';E={if($_.id -eq $recommendedNew.id){'★'}else{''}}},
            @{N='Default';E={if($_.defaultFlag){'★'}else{''}}},
            @{N='Status ID';E={$_.id}} | 
            Out-GridView -Title "Select New Ticket Status | Current: $currentNewStatus | Recommended: $($recommendedNew.name)" -OutputMode Single
        
        if ($selectedNewStatus) {
            $Script:SelectedConfig.NewStatus = $selectedNewStatus.'Status Name'
            Write-Host "  Selected New Status: $($selectedNewStatus.'Status Name')" -ForegroundColor Green
        }
        
        # Select Closed Ticket Status
        Write-Host "`n  Step 3: Select Closed Ticket Status" -ForegroundColor Yellow
        $closedStatuses = $boardStatuses | Where-Object { $_.closedStatus } | Sort-Object sortOrder
        
        if ($closedStatuses.Count -eq 0) {
            Write-Warning "    No closed statuses found for board '$($Script:SelectedConfig.Board)'"
            Write-Host "    Available statuses (select any as closed status):" -ForegroundColor Yellow
            
            # Find a reasonable recommendation - prefer status with "closed" in name
            $recommendedClosed = $boardStatuses | Where-Object { $_.name -like '*closed*' } | Select-Object -First 1
            if (-not $recommendedClosed) {
                $recommendedClosed = $boardStatuses | Sort-Object sortOrder | Select-Object -First 1
            }
            
            # Show all statuses if no closed ones found
            $selectedClosedStatus = $boardStatuses | Sort-Object sortOrder |
                Select-Object @{N='Status Name';E={$_.name}},
                    @{N='Type';E={if($_.closedStatus){'Closed'}elseif($_.defaultFlag){'Default'}else{'Open'}}},
                    @{N='Current';E={if($_.name -eq $currentClosedStatus){'✓'}else{''}}},
                    @{N='Recommended';E={if($_.id -eq $recommendedClosed.id){'★'}else{''}}},
                    @{N='Status ID';E={$_.id}} | 
                Out-GridView -Title "Select Closed Ticket Status (No closed statuses found - select any) | Current: $currentClosedStatus | Recommended: $($recommendedClosed.name)" -OutputMode Single
        } else {
            $recommendedClosed = $closedStatuses | Select-Object -First 1
            
            $selectedClosedStatus = $closedStatuses | Select-Object @{N='Status Name';E={$_.name}},
                @{N='Current';E={if($_.name -eq $currentClosedStatus){'✓'}else{''}}},
                @{N='Recommended';E={if($_.id -eq $recommendedClosed.id){'★'}else{''}}},
                @{N='Status ID';E={$_.id}} | 
                Out-GridView -Title "Select Closed Ticket Status | Current: $currentClosedStatus | Recommended: $($recommendedClosed.name)" -OutputMode Single
        }
        
        if ($selectedClosedStatus) {
            $Script:SelectedConfig.ClosedStatus = $selectedClosedStatus.'Status Name'
            Write-Host "  Selected Closed Status: $($selectedClosedStatus.'Status Name')" -ForegroundColor Green
        } else {
            Write-Warning "    No closed status selected - this parameter is required!"
        }
    }
    
    # Select Server Priority
    Write-Host "`n  Step 4a: Select Server Priority" -ForegroundColor Yellow
    $selectedPriorityServer = $priorities | Sort-Object sort | 
        Select-Object @{N='Priority Name';E={$_.name}},
            @{N='Recommended';E={if($_.id -eq $priority1.id){'★ Critical/Emergency'}elseif($_.id -eq $priority2.id){'Urgent'}else{''}}},
            @{N='Current';E={if($_.name -eq $currentPriorityServer){'✓'}else{''}}},
            @{N='Priority ID';E={$_.id}} |
        Out-GridView -Title "Select Server Ticket Priority | Current: $currentPriorityServer | Recommended: $($priority1.name)" -OutputMode Single
    
    if ($selectedPriorityServer) {
        $Script:SelectedConfig.PriorityServer = $selectedPriorityServer.'Priority Name'
        Write-Host "  Selected Server Priority: $($selectedPriorityServer.'Priority Name')" -ForegroundColor Green
    }
    
    # Select Workstation Priority
    Write-Host "`n  Step 4b: Select Workstation Priority" -ForegroundColor Yellow
    $selectedPriorityWorkstation = $priorities | Sort-Object sort | 
        Select-Object @{N='Priority Name';E={$_.name}},
            @{N='Recommended';E={if($_.id -eq $priority3.id){'★ Normal Response'}elseif($_.id -eq $priority4.id){'Low/Scheduled'}else{''}}},
            @{N='Current';E={if($_.name -eq $currentPriorityWorkstation){'✓'}else{''}}},
            @{N='Priority ID';E={$_.id}} |
        Out-GridView -Title "Select Workstation Ticket Priority | Current: $currentPriorityWorkstation | Recommended: $($priority3.name)" -OutputMode Single
    
    if ($selectedPriorityWorkstation) {
        $Script:SelectedConfig.PriorityWorkstation = $selectedPriorityWorkstation.'Priority Name'
        Write-Host "  Selected Workstation Priority: $($selectedPriorityWorkstation.'Priority Name')" -ForegroundColor Green
    }
    
    # Select M365 Priority
    Write-Host "`n  Step 4c: Select M365 Priority" -ForegroundColor Yellow
    $selectedPriorityM365 = $priorities | Sort-Object sort | 
        Select-Object @{N='Priority Name';E={$_.name}},
            @{N='Recommended';E={if($_.id -eq $priority2.id){'★ Quick Response'}elseif($_.id -eq $priority3.id){'Normal'}else{''}}},
            @{N='Current';E={if($_.name -eq $currentPriorityM365){'✓'}else{''}}},
            @{N='Priority ID';E={$_.id}} |
        Out-GridView -Title "Select M365 Ticket Priority | Current: $currentPriorityM365 | Recommended: $($priority2.name)" -OutputMode Single
    
    if ($selectedPriorityM365) {
        $Script:SelectedConfig.PriorityM365 = $selectedPriorityM365.'Priority Name'
        Write-Host "  Selected M365 Priority: $($selectedPriorityM365.'Priority Name')" -ForegroundColor Green
    }
    
    # Display final configuration
    Write-Host "`n" -NoNewline
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "  FINAL CONFIGURATION" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "  Service Board         : $($Script:SelectedConfig.Board)" -ForegroundColor White
    Write-Host "  New Ticket Status     : $($Script:SelectedConfig.NewStatus)" -ForegroundColor White
    Write-Host "  Closed Ticket Status  : $($Script:SelectedConfig.ClosedStatus)" -ForegroundColor White
    Write-Host "  Server Priority       : $($Script:SelectedConfig.PriorityServer)" -ForegroundColor White
    Write-Host "  Workstation Priority  : $($Script:SelectedConfig.PriorityWorkstation)" -ForegroundColor White
    Write-Host "  M365 Priority         : $($Script:SelectedConfig.PriorityM365)" -ForegroundColor White
    
    # Ask to update monitoring script
    Write-Host "`n  Would you like to update the monitoring script with these settings?" -ForegroundColor Yellow
    $updateScript = Read-Host "  Enter 'Y' to update $($MonitoringScriptPath.Split('\')[-1]) (or use -WhatIf to preview)"
    
    if ($updateScript -eq 'Y' -or $updateScript -eq 'y') {
        if (Test-Path $MonitoringScriptPath) {
            try {
                # Create backup folder if it doesn't exist
                $backupFolder = Join-Path (Split-Path $MonitoringScriptPath) "Backups"
                if (-not (Test-Path $backupFolder)) {
                    New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
                }
                
                # Create timestamped backup file
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $backupFileName = "CoveMonitoring_Backup_$timestamp.backup"
                $backupFile = Join-Path $backupFolder $backupFileName
                
                if ($PSCmdlet.ShouldProcess($MonitoringScriptPath, "Create backup and update configuration")) {
                    Copy-Item $MonitoringScriptPath $backupFile -Force
                    Write-Host "`n  Backup created: $backupFile" -ForegroundColor Gray
                    
                    $scriptContent = Get-Content $MonitoringScriptPath -Raw
                    
                    # Update parameters with selected values
                    if ($Script:SelectedConfig.Board) {
                        $newBoardValue = $Script:SelectedConfig.Board
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketBoard\s*=\s*")[^"]*(")' , "`${1}$newBoardValue`${2}"
                    }
                    if ($Script:SelectedConfig.NewStatus) {
                        $newStatusValue = $Script:SelectedConfig.NewStatus
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketStatus\s*=\s*")[^"]*(")' , "`${1}$newStatusValue`${2}"
                    }
                    if ($Script:SelectedConfig.ClosedStatus) {
                        $newClosedValue = $Script:SelectedConfig.ClosedStatus
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketClosedStatus\s*=\s*")[^"]*(")' , "`${1}$newClosedValue`${2}"
                    }
                    
                    # Debug: Show priority values before write
                    Write-Host "`n  DEBUG - Priority values before write:" -ForegroundColor Magenta
                    Write-Host "    PriorityServer       : '$($Script:SelectedConfig.PriorityServer)'" -ForegroundColor Magenta
                    Write-Host "    PriorityWorkstation  : '$($Script:SelectedConfig.PriorityWorkstation)'" -ForegroundColor Magenta
                    Write-Host "    PriorityM365         : '$($Script:SelectedConfig.PriorityM365)'" -ForegroundColor Magenta
                    
                    # Build replacement strings separately to avoid backreference conflicts with "Priority 1"
                    if ($Script:SelectedConfig.PriorityServer) {
                        $newServerValue = $Script:SelectedConfig.PriorityServer
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketPriorityServer\s*=\s*")[^"]*(")', "`${1}$newServerValue`${2}"
                    }
                    if ($Script:SelectedConfig.PriorityWorkstation) {
                        $newWorkstationValue = $Script:SelectedConfig.PriorityWorkstation
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketPriorityWorkstation\s*=\s*")[^"]*(")', "`${1}$newWorkstationValue`${2}"
                    }
                    if ($Script:SelectedConfig.PriorityM365) {
                        $newM365Value = $Script:SelectedConfig.PriorityM365
                        $scriptContent = $scriptContent -replace '(\[string\]\$TicketPriorityM365\s*=\s*")[^"]*(")', "`${1}$newM365Value`${2}"
                    }
                    
                    # Save updated script
                    $scriptContent | Set-Content $MonitoringScriptPath -Force
                    
                    # Validate the updates were applied correctly
                    Write-Host "`n  Validating parameter updates..." -ForegroundColor Cyan
                    $validationContent = Get-Content $MonitoringScriptPath -Raw
                    $validationErrors = @()
                    
                    if ($Script:SelectedConfig.Board) {
                        if ($validationContent -match '\$TicketBoard\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.Board) {
                                Write-Host "  ✓ TicketBoard: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketBoard: Expected '$($Script:SelectedConfig.Board)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketBoard mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketBoard: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketBoard not found"
                        }
                    }
                    
                    if ($Script:SelectedConfig.NewStatus) {
                        if ($validationContent -match '\$TicketStatus\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.NewStatus) {
                                Write-Host "  ✓ TicketStatus: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketStatus: Expected '$($Script:SelectedConfig.NewStatus)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketStatus mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketStatus: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketStatus not found"
                        }
                    }
                    
                    if ($Script:SelectedConfig.ClosedStatus) {
                        if ($validationContent -match '\$TicketClosedStatus\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.ClosedStatus) {
                                Write-Host "  ✓ TicketClosedStatus: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketClosedStatus: Expected '$($Script:SelectedConfig.ClosedStatus)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketClosedStatus mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketClosedStatus: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketClosedStatus not found"
                        }
                    }
                    
                    if ($Script:SelectedConfig.PriorityServer) {
                        if ($validationContent -match '\$TicketPriorityServer\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.PriorityServer) {
                                Write-Host "  ✓ TicketPriorityServer: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketPriorityServer: Expected '$($Script:SelectedConfig.PriorityServer)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketPriorityServer mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketPriorityServer: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketPriorityServer not found"
                        }
                    }
                    
                    if ($Script:SelectedConfig.PriorityWorkstation) {
                        if ($validationContent -match '\$TicketPriorityWorkstation\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.PriorityWorkstation) {
                                Write-Host "  ✓ TicketPriorityWorkstation: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketPriorityWorkstation: Expected '$($Script:SelectedConfig.PriorityWorkstation)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketPriorityWorkstation mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketPriorityWorkstation: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketPriorityWorkstation not found"
                        }
                    }
                    
                    if ($Script:SelectedConfig.PriorityM365) {
                        if ($validationContent -match '\$TicketPriorityM365\s*=\s*"([^"]*)"') {
                            $actualValue = $matches[1]
                            if ($actualValue -eq $Script:SelectedConfig.PriorityM365) {
                                Write-Host "  ✓ TicketPriorityM365: '$actualValue'" -ForegroundColor Green
                            } else {
                                Write-Host "  ✗ TicketPriorityM365: Expected '$($Script:SelectedConfig.PriorityM365)' but got '$actualValue'" -ForegroundColor Red
                                $validationErrors += "TicketPriorityM365 mismatch"
                            }
                        } else {
                            Write-Host "  ✗ TicketPriorityM365: Parameter not found in file" -ForegroundColor Red
                            $validationErrors += "TicketPriorityM365 not found"
                        }
                    }
                    
                    if ($validationErrors.Count -eq 0) {
                        Write-Host "`n  SUCCESS: All parameters validated successfully!" -ForegroundColor Green
                    } else {
                        Write-Host "`n  WARNING: $($validationErrors.Count) validation error(s) detected!" -ForegroundColor Yellow
                        Write-Host "  Check the monitoring script manually or restore from backup" -ForegroundColor Yellow
                    }
                    
                    Write-Host "`n  File: $MonitoringScriptPath" -ForegroundColor Gray
                    Write-Host "  Backup: $backupFileName" -ForegroundColor Gray
                    Write-Host "`n  To undo this change, run: .\Cove2CWM-SetTicketsConfig.v10.ps1 -Undo" -ForegroundColor Cyan
                    
                    # Clean up old backups (keep last 10)
                    $allBackups = Get-ChildItem -Path $backupFolder -Filter "*.backup" | Sort-Object LastWriteTime -Descending
                    if ($allBackups.Count -gt 10) {
                        $allBackups | Select-Object -Skip 10 | Remove-Item -Force
                        Write-Host "  (Cleaned up old backups - keeping most recent 10)" -ForegroundColor DarkGray
                    }
                } else {
                    Write-Host "`n  WHATIF: Would update monitoring script with:" -ForegroundColor Yellow
                    Write-Host "    -TicketBoard '$($Script:SelectedConfig.Board)'" -ForegroundColor Gray
                    Write-Host "    -TicketStatus '$($Script:SelectedConfig.NewStatus)'" -ForegroundColor Gray
                    Write-Host "    -TicketClosedStatus '$($Script:SelectedConfig.ClosedStatus)'" -ForegroundColor Gray
                    Write-Host "    -TicketPriorityServer '$($Script:SelectedConfig.PriorityServer)'" -ForegroundColor Gray
                    Write-Host "    -TicketPriorityWorkstation '$($Script:SelectedConfig.PriorityWorkstation)'" -ForegroundColor Gray
                    Write-Host "    -TicketPriorityM365 '$($Script:SelectedConfig.PriorityM365)'" -ForegroundColor Gray
                    Write-Host "`n  Backup would be created in: $backupFolder" -ForegroundColor Gray
                    Write-Host "  Backup filename: $backupFileName" -ForegroundColor Gray
                }
                
            } catch {
                Write-Warning "Error updating monitoring script: $_"
            }
        } else {
            Write-Warning "Monitoring script not found at: $MonitoringScriptPath"
            Write-Host "  You can manually copy these values:" -ForegroundColor Yellow
            Write-Host "    -TicketBoard '$($Script:SelectedConfig.Board)'" -ForegroundColor Gray
            Write-Host "    -TicketStatus '$($Script:SelectedConfig.NewStatus)'" -ForegroundColor Gray
            Write-Host "    -TicketClosedStatus '$($Script:SelectedConfig.ClosedStatus)'" -ForegroundColor Gray
            Write-Host "    -TicketPriorityServer '$($Script:SelectedConfig.PriorityServer)'" -ForegroundColor Gray
            Write-Host "    -TicketPriorityWorkstation '$($Script:SelectedConfig.PriorityWorkstation)'" -ForegroundColor Gray
            Write-Host "    -TicketPriorityM365 '$($Script:SelectedConfig.PriorityM365)'" -ForegroundColor Gray
        }
    } else {
        Write-Host "`n  Configuration not applied. Manual parameter values:" -ForegroundColor Yellow
        Write-Host "    -TicketBoard '$($Script:SelectedConfig.Board)'" -ForegroundColor Gray
        Write-Host "    -TicketStatus '$($Script:SelectedConfig.NewStatus)'" -ForegroundColor Gray
        Write-Host "    -TicketClosedStatus '$($Script:SelectedConfig.ClosedStatus)'" -ForegroundColor Gray
        Write-Host "    -TicketPriorityServer '$($Script:SelectedConfig.PriorityServer)'" -ForegroundColor Gray
        Write-Host "    -TicketPriorityWorkstation '$($Script:SelectedConfig.PriorityWorkstation)'" -ForegroundColor Gray
        Write-Host "    -TicketPriorityM365 '$($Script:SelectedConfig.PriorityM365)'" -ForegroundColor Gray
    }
}

#endregion ----- Interactive Selection ----

Write-Host ""

#endregion


