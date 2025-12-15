#Requires -Version 5.1
<# ----- About: ----
    # N-able Backup Manager - Intelligent Backup Failure Monitor
    # Revision v02 - 2025-12-15
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
#>

<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able Cove Data Protection
    # Requires PowerShell 5.1 or later
    # Requires ClientTool PowerShell module (auto-installs if missing)
#>

<# ----- Behavior: ----
    # Monitors local N-able Backup Manager for datasource backup failures with intelligent alerting
    #
    # Features:
    # - Automatically detects device type (Server vs Workstation) from OS
    # - Uses different failure thresholds based on device type
    # - Analyzes historical backup frequency patterns per datasource
    # - For datasources with multiple daily backups: Alerts after 3 consecutive failures
    # - For datasources with daily or less frequent backups: Alerts after 1 failure if age exceeds threshold
    # - Monitors Local SpeedVault (LSV) synchronization status and failures
    # - Auto-installs and updates ClientTool module from GitHub
    # - Captures last error message from failed sessions
    # - Returns appropriate exit codes for RMM/monitoring integration
    #
    # Exit Codes:
    #   0 = Success - All datasources healthy
    #   1 = Warning - Datasource failures detected
    #   2 = Critical - No recent successful backups or errors during execution
    #
    # Parameters:
    #   -ServerDays (default=2) - Days before alerting on server datasource failures
    #   -WorkstationDays (default=3) - Days before alerting on workstation datasource failures
    #   -HistoryDays (default=14) - Days of history to analyze for backup frequency patterns
    #   -MultiDailyThreshold (default=3) - Number of consecutive failures before alerting on multi-daily datasources
    #   -Verbose - Show detailed analysis information
    #
    # Examples:
    #   .\Monitor-BackupStatus.v01.ps1
    #   .\Monitor-BackupStatus.v01.ps1 -ServerDays 1 -WorkstationDays 2
    #   .\Monitor-BackupStatus.v01.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$ServerDays = 2,
    
    [Parameter(Mandatory=$false)]
    [int]$WorkstationDays = 3,
    
    [Parameter(Mandatory=$false)]
    [int]$HistoryDays = 14,
    
    [Parameter(Mandatory=$false)]
    [int]$MultiDailyThreshold = 3,
    
    [Parameter(Mandatory=$false)]
    [switch]$OutputJson,
    
    [Parameter(Mandatory=$false)]
    [switch]$DumpFailureJson
)

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $scriptVersion = "2.0"
    $scriptStartTime = Get-Date
    $statusXmlPath = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml"
    
    Write-Host "============================================================================================================" -ForegroundColor Cyan
    Write-Host "  N-able Backup Manager - Intelligent Backup Failure Monitor v$scriptVersion" -ForegroundColor White
    Write-Host "============================================================================================================" -ForegroundColor Cyan
    Write-Host ""
#endregion

#region ----- Functions ----

#region ----- Module Management ----
Function Install-ClientToolModule {
    <#
    .SYNOPSIS
        Ensures ClientTool module is installed and up to date
    #>
    [CmdletBinding()]
    param()
    
    $moduleLoaded = Get-Command Get-ClientToolSession -ErrorAction SilentlyContinue

    if ($moduleLoaded) {
        # Check if we should update (check version from GitHub)
        try {
            Write-Host "Checking for ClientTool module updates..." -ForegroundColor Cyan
            $currentModule = Get-Module ClientTool
            if (-not $currentModule) {
                $currentModule = Get-Module ClientTool -ListAvailable | Select-Object -First 1
            }
            
            # Get version from GitHub manifest
            $manifestContent = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/BackupNerd/ClientTool/main/ClientTool.psd1" -ErrorAction SilentlyContinue
            if ($manifestContent -match "ModuleVersion\s*=\s*'([^']+)'") {
                $githubVersion = [version]$matches[1]
                $currentVersion = [version]$currentModule.Version
                
                Write-Host "  Current version: $currentVersion" -ForegroundColor White
                Write-Host "  Cloud version:   $githubVersion" -ForegroundColor White
                
                if ($githubVersion -gt $currentVersion) {
                    Write-Host "  Status: Update available" -ForegroundColor Yellow
                    Write-Host "  Updating ClientTool module..." -ForegroundColor Yellow
                    Remove-Module ClientTool -Force -ErrorAction SilentlyContinue
                    Invoke-RestMethod https://raw.githubusercontent.com/BackupNerd/ClientTool/main/Install-Module.ps1 | Invoke-Expression
                    Write-Host "  ✓ Module updated successfully to v$githubVersion!" -ForegroundColor Green
                }
                else {
                    Write-Host "  Status: Up to date ✓" -ForegroundColor Green
                }
            }
            else {
                Write-Host "  Current version: $($currentModule.Version)" -ForegroundColor White
                Write-Host "  Unable to check cloud version. Using installed version." -ForegroundColor Gray
            }
        }
        catch {
            Write-Host "  Unable to check for updates: $_" -ForegroundColor Gray
            Write-Host "  Using installed version: $($currentModule.Version)" -ForegroundColor Gray
        }
    }
    else {
        # Module not loaded, install from GitHub
        Write-Host "ClientTool module not found. Installing from GitHub..." -ForegroundColor Yellow
        try {
            Invoke-RestMethod https://raw.githubusercontent.com/BackupNerd/ClientTool/main/Install-Module.ps1 | Invoke-Expression
            if (-not (Get-Command Get-ClientToolSession -ErrorAction SilentlyContinue)) {
                Write-Error "ClientTool module installation failed. Please install manually."
                return $false
            }
            $installedModule = Get-Module ClientTool
            Write-Host "ClientTool module installed successfully (v$($installedModule.Version))." -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to install ClientTool module: $_"
            return $false
        }
    }
    
    return $true
}
#endregion

#region ----- Device Detection ----
Function Get-DeviceType {
    <#
    .SYNOPSIS
        Determines if device is a Server or Workstation from StatusReport.xml or WMI
    .OUTPUTS
        String - "Server" or "Workstation"
    #>
    
    # Try to get from StatusReport.xml first (faster and more reliable)
    $statusXmlPath = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml"
    if (Test-Path $statusXmlPath) {
        try {
            [xml]$statusXml = Get-Content $statusXmlPath -ErrorAction Stop
            if ($statusXml.Statistics.OsVersion) {
                $osVersion = $statusXml.Statistics.OsVersion
                
                # Check if OS string contains "Server"
                if ($osVersion -match 'Server') {
                    return "Server"
                }
                else {
                    return "Workstation"
                }
            }
        }
        catch {
            # Fall through to WMI method
        }
    }
    
    # Fallback to WMI if StatusReport.xml unavailable or parsing failed
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        
        if ($osInfo.ProductType -eq 1) {
            return "Workstation"
        }
        else {
            return "Server"
        }
    }
    catch {
        Write-Warning "Unable to detect device type, defaulting to Workstation"
        return "Workstation"
    }
}
#endregion

#region ----- Backup Analysis ----
Function Get-DatasourceBackupPattern {
    <#
    .SYNOPSIS
        Analyzes historical backup frequency for a datasource
    .PARAMETER Sessions
        Array of backup sessions for a specific datasource
    .PARAMETER HistoryDays
        Number of days of history being analyzed
    .OUTPUTS
        Hashtable with pattern analysis results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Sessions,
        
        [Parameter(Mandatory=$false)]
        [int]$HistoryDays = 14
    )
    
    if ($Sessions.Count -eq 0) {
        return @{
            IsMultiDaily = $false
            AveragePerDay = 0
            DaysSinceLastSuccess = 999
            ConsecutiveFailures = 0
            Pattern = "No History"
        }
    }
    
    # Get successful sessions only for pattern analysis
    $successfulSessions = @($Sessions | Where-Object { $_.STATE -eq 'Completed' })
    
    # Get last session (most recent regardless of status)
    $lastSession = $Sessions | 
        Sort-Object { [DateTime]::Parse($_.START) } -Descending | 
        Select-Object -First 1
    
    # Find last successful backup
    $lastSuccess = $successfulSessions | 
        Sort-Object { [DateTime]::Parse($_.START) } -Descending | 
        Select-Object -First 1
    
    if ($lastSuccess) {
        $daysSinceSuccess = ((Get-Date) - [DateTime]::Parse($lastSuccess.START)).TotalDays
    }
    else {
        $daysSinceSuccess = 999
    }
    
    # Calculate consecutive failures (sessions after last success)
    $consecutiveFailures = 0
    $sortedSessions = $Sessions | Sort-Object { [DateTime]::Parse($_.START) } -Descending
    
    foreach ($session in $sortedSessions) {
        if ($session.STATE -eq 'Completed') {
            break
        }
        $consecutiveFailures++
    }
    
    # Determine backup frequency pattern
    if ($successfulSessions.Count -ge 3) {
        # Group by date to count backups per day
        $sessionsByDate = $successfulSessions | Group-Object { ([DateTime]::Parse($_.START)).Date }
        
        # Calculate average backups per day (only for days with backups)
        $totalBackupsPerDay = ($sessionsByDate | Measure-Object -Property Count -Average).Average
        
        # Calculate total span of days
        $oldestSession = $successfulSessions | 
            Sort-Object { [DateTime]::Parse($_.START) } | 
            Select-Object -First 1
        $newestSession = $successfulSessions | 
            Sort-Object { [DateTime]::Parse($_.START) } -Descending | 
            Select-Object -First 1
        
        $daySpan = ([DateTime]::Parse($newestSession.START) - [DateTime]::Parse($oldestSession.START)).TotalDays
        if ($daySpan -eq 0) { $daySpan = 1 }
        
        # Average backups per calendar day
        $avgPerDay = $successfulSessions.Count / $daySpan
        
        # If average is > 1.5 backups per day, consider it multi-daily
        $isMultiDaily = $avgPerDay -gt 1.5
        
        $pattern = if ($isMultiDaily) { 
            "Multi-Daily (avg $([math]::Round($avgPerDay, 1))/day)" 
        } 
        else { 
            "Daily or Less (avg $([math]::Round($avgPerDay, 1))/day)" 
        }
    }
    else {
        # Not enough successful sessions to determine pattern
        # Default to daily pattern (less aggressive alerting)
        $isMultiDaily = $false
        $avgPerDay = 0
        
        if ($successfulSessions.Count -eq 0) {
            $pattern = "No Successful Backups (${HistoryDays}d history)"
        }
        elseif ($Sessions.Count -lt 3) {
            $pattern = "Limited History ($($Sessions.Count) sessions in ${HistoryDays}d)"
        }
        else {
            $pattern = "Mostly Failures ($($successfulSessions.Count) of $($Sessions.Count) OK in ${HistoryDays}d)"
        }
    }
    
    # Get last error message if session failed
    $lastError = "N/A"
    if ($lastSession -and ($lastSession.STATE -eq 'Failed' -or $lastSession.STATE -eq 'FailedBlocked')) {
        try {
            # Map datasource name to proper format for Get-ClientToolSessionError
            $datasourceName = $lastSession.DSRC
            $sessionTime = [DateTime]::Parse($lastSession.START)
            
            # Get errors from this specific session (limit to 1 for last error)
            $errors = Get-ClientToolSessionError -DataSource $datasourceName -Time $sessionTime -Limit 1 -ErrorAction SilentlyContinue
            
            if ($errors -and $errors.Count -gt 0) {
                $lastError = $errors[0].CONTENT
                # Truncate if too long
                if ($lastError.Length -gt 150) {
                    $lastError = $lastError.Substring(0, 147) + "..."
                }
            }
        }
        catch {
            # Silently fail if error retrieval doesn't work
            $lastError = "[Error retrieval failed]"
        }
    }
    
    return @{
        IsMultiDaily = $isMultiDaily
        AveragePerDay = [math]::Round($avgPerDay, 2)
        DaysSinceLastSuccess = [math]::Round($daysSinceSuccess, 1)
        ConsecutiveFailures = $consecutiveFailures
        Pattern = $pattern
        TotalSessions = $Sessions.Count
        SuccessfulSessions = $successfulSessions.Count
        LastSessionStart = if ($lastSession) { $lastSession.START } else { "N/A" }
        LastSessionState = if ($lastSession) { $lastSession.STATE } else { "N/A" }
        LastSessionEnd = if ($lastSession -and $lastSession.END) { $lastSession.END } else { "N/A" }
        LastError = $lastError
    }
}

Function Test-DatasourceHealth {
    <#
    .SYNOPSIS
        Determines if a datasource should trigger an alert based on pattern and thresholds
    .OUTPUTS
        Hashtable with alert status and details
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Pattern,
        
        [Parameter(Mandatory=$true)]
        [int]$DayThreshold,
        
        [Parameter(Mandatory=$true)]
        [int]$MultiDailyFailureThreshold,
        
        [Parameter(Mandatory=$false)]
        [int]$HistoryDays = 14
    )
    
    $shouldAlert = $false
    $alertReason = ""
    $severity = "OK"
    
    # No successful backups ever
    if ($Pattern.SuccessfulSessions -eq 0) {
        $shouldAlert = $true
        $alertReason = "No successful backups found in last ${HistoryDays} days"
        $severity = "CRITICAL"
    }
    # Multi-daily pattern: Alert after consecutive failures threshold
    elseif ($Pattern.IsMultiDaily) {
        if ($Pattern.ConsecutiveFailures -ge $MultiDailyFailureThreshold) {
            $shouldAlert = $true
            $alertReason = "Multi-daily datasource: $($Pattern.ConsecutiveFailures) consecutive failures (threshold: $MultiDailyFailureThreshold)"
            $severity = "WARNING"
        }
    }
    # Daily or less pattern: Alert if age exceeds threshold OR first failure
    else {
        if ($Pattern.DaysSinceLastSuccess -ge $DayThreshold) {
            $shouldAlert = $true
            $alertReason = "Last successful backup $([math]::Round($Pattern.DaysSinceLastSuccess, 1)) days ago (threshold: $DayThreshold days)"
            $severity = "WARNING"
        }
        elseif ($Pattern.ConsecutiveFailures -ge 1) {
            $shouldAlert = $true
            $alertReason = "Recent failure detected ($($Pattern.ConsecutiveFailures) failed session(s))"
            $severity = "WARNING"
        }
    }
    
    return @{
        ShouldAlert = $shouldAlert
        AlertReason = $alertReason
        Severity = $severity
    }
}
#endregion

#endregion

#region ----- Main Script ----

# Install/Update ClientTool Module
$moduleStatus = Install-ClientToolModule
if (-not $moduleStatus) {
    Write-Host "`n[CRITICAL] Cannot proceed without ClientTool module" -ForegroundColor Red
    exit 2
}

Write-Host ""

# Detect device type
$deviceType = Get-DeviceType
$dayThreshold = if ($deviceType -eq "Server") { $ServerDays } else { $WorkstationDays }

Write-Host "Device Information:" -ForegroundColor Yellow
Write-Host "  Type: $deviceType" -ForegroundColor White
Write-Host "  Failure Threshold: $dayThreshold days" -ForegroundColor White
Write-Host "  Multi-Daily Failure Count: $MultiDailyThreshold consecutive failures" -ForegroundColor White
Write-Host ""

# Get device name, customer name, and machine name from StatusReport.xml or fallback
$deviceName = $env:COMPUTERNAME
$customerName = "Unknown"
$machineName = $env:COMPUTERNAME

if (Test-Path $statusXmlPath) {
    try {
        [xml]$statusXml = Get-Content $statusXmlPath -ErrorAction SilentlyContinue
        if ($statusXml.Statistics.Account) {
            $deviceName = $statusXml.Statistics.Account
        }
        if ($statusXml.Statistics.PartnerName) {
            $customerName = $statusXml.Statistics.PartnerName
        }
        if ($statusXml.Statistics.MachineName) {
            $machineName = $statusXml.Statistics.MachineName
        }
    }
    catch {
        Write-Verbose "Unable to read StatusReport.xml, using defaults"
    }
}

Write-Host "Analyzing Backup Status for: $deviceName" -ForegroundColor Cyan
Write-Host ""

# Check LSV Status
$lsvEnabled = $false
$lsvStatus = "N/A"
$lsvAlert = $null

if (Test-Path $statusXmlPath) {
    try {
        [xml]$statusXml = Get-Content $statusXmlPath -ErrorAction SilentlyContinue
        $lsvEnabled = if ($statusXml.Statistics.LocalSpeedVaultEnabled) { [int]$statusXml.Statistics.LocalSpeedVaultEnabled -eq 1 } else { $false }
        
        if ($lsvEnabled) {
            $lsvStatus = if ($statusXml.Statistics.LocalSpeedVaultSynchronizationStatus) { 
                $statusXml.Statistics.LocalSpeedVaultSynchronizationStatus 
            } else { 
                "Unknown" 
            }
            
            # Check for LSV sync issues
            if ($lsvStatus -notin @('Synchronized', 'N/A')) {
                $lsvAlert = [PSCustomObject]@{
                    Component = "Local SpeedVault"
                    Status = $lsvStatus
                    Severity = "WARNING"
                    Reason = "LSV synchronization status: $lsvStatus"
                }
            }
            
            Write-Host "Local SpeedVault Status:" -ForegroundColor Yellow
            $statusColor = if ($lsvStatus -eq 'Synchronized') { "Green" } else { "Yellow" }
            Write-Host "  Enabled: Yes" -ForegroundColor White
            Write-Host "  Sync Status: $lsvStatus" -ForegroundColor $statusColor
            Write-Host ""
        }
    }
    catch {
        Write-Verbose "Unable to check LSV status from StatusReport.xml"
    }
}

# Gather session data
Write-Host "Gathering session data (last $HistoryDays days)..." -ForegroundColor Yellow
$cutoffDate = (Get-Date).AddDays(-$HistoryDays)

try {
    $allSessions = Get-ClientToolSession | Where-Object { 
        $_.TYPE -eq 'Backup' -and
        $_.START -and 
        $_.START -ne '' -and
        [DateTime]::Parse($_.START) -ge $cutoffDate
    }
    
    Write-Host "  Found $($allSessions.Count) backup sessions" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to retrieve backup sessions: $_" -ForegroundColor Red
    exit 2
}

if ($allSessions.Count -eq 0) {
    Write-Host "[CRITICAL] No backup sessions found in last $HistoryDays days!" -ForegroundColor Red
    exit 2
}

# Group sessions by datasource
$datasourceGroups = $allSessions | Group-Object DSRC

Write-Host "`nDatasource Analysis:" -ForegroundColor Yellow
Write-Host ("=" * 120) -ForegroundColor Gray
Write-Host ("{0,-30} {1,-35} {2,-12} {3,-10} {4,-15}" -f "Datasource", "Pattern", "Last OK", "Failures", "Status") -ForegroundColor Cyan
Write-Host ("=" * 120) -ForegroundColor Gray

$alerts = @()
$healthyCount = 0
$warningCount = 0

foreach ($dsGroup in $datasourceGroups | Sort-Object Name) {
    $datasourceName = $dsGroup.Name
    $sessions = $dsGroup.Group
    
    # Analyze pattern
    $pattern = Get-DatasourceBackupPattern -Sessions $sessions -HistoryDays $HistoryDays
    
    # Test health
    $health = Test-DatasourceHealth -Pattern $pattern -DayThreshold $dayThreshold -MultiDailyFailureThreshold $MultiDailyThreshold -HistoryDays $HistoryDays
    
    # Format days since last success
    $daysSinceStr = if ($pattern.DaysSinceLastSuccess -eq 999) { 
        "Never" 
    } 
    else { 
        "$([math]::Round($pattern.DaysSinceLastSuccess, 1))d ago" 
    }
    
    # Color coding based on severity
    $statusColor = switch ($health.Severity) {
        "CRITICAL" { "Red" }
        "WARNING" { "Yellow" }
        default { "Green" }
    }
    
    $statusText = if ($health.ShouldAlert) { 
        $health.Severity 
    } 
    else { 
        "OK" 
    }
    
    # Display row
    $line = "{0,-30} {1,-35} {2,-12} {3,-10} {4,-15}" -f `
        $datasourceName.Substring(0, [Math]::Min(29, $datasourceName.Length)),
        $pattern.Pattern.Substring(0, [Math]::Min(34, $pattern.Pattern.Length)),
        $daysSinceStr,
        $pattern.ConsecutiveFailures,
        $statusText
    
    Write-Host $line -ForegroundColor $statusColor
    
    # Track alerts
    if ($health.ShouldAlert) {
        $warningCount++
        $alerts += [PSCustomObject]@{
            Datasource = $datasourceName
            Pattern = $pattern.Pattern
            DaysSinceSuccess = $pattern.DaysSinceLastSuccess
            ConsecutiveFailures = $pattern.ConsecutiveFailures
            Reason = $health.AlertReason
            Severity = $health.Severity
            LastSessionStart = $pattern.LastSessionStart
            LastSessionState = $pattern.LastSessionState
            LastSessionEnd = $pattern.LastSessionEnd
            LastError = $pattern.LastError
        }
    }
    else {
        $healthyCount++
    }
    
    # Verbose output
    if ($VerbosePreference -eq 'Continue') {
        Write-Host "    └─ Total: $($pattern.TotalSessions) sessions | Successful: $($pattern.SuccessfulSessions) | Avg/Day: $($pattern.AveragePerDay)" -ForegroundColor DarkGray
        Write-Host "    └─ Last Session: $($pattern.LastSessionStart) | State: $($pattern.LastSessionState)" -ForegroundColor DarkGray
        if ($health.ShouldAlert) {
            Write-Host "    └─ ALERT: $($health.AlertReason)" -ForegroundColor $statusColor
        }
    }
}

Write-Host ("=" * 120) -ForegroundColor Gray
Write-Host ""

# Summary
Write-Host "`nSummary:" -ForegroundColor Yellow
Write-Host "  Healthy Datasources: $healthyCount" -ForegroundColor Green
Write-Host "  Datasources with Alerts: $warningCount" -ForegroundColor $(if ($warningCount -gt 0) { "Yellow" } else { "Green" })
if ($lsvEnabled) {
    $lsvColor = if ($lsvAlert) { "Yellow" } else { "Green" }
    $lsvStatusText = if ($lsvAlert) { "Issue Detected" } else { "Healthy" }
    Write-Host "  Local SpeedVault: $lsvStatusText" -ForegroundColor $lsvColor
}
Write-Host ""

# Detailed alerts
if ($alerts.Count -gt 0 -or $lsvAlert) {
    Write-Host ""
    Write-Host "⚠ ALERTS DETECTED:" -ForegroundColor Yellow
    Write-Host ("=" * 120) -ForegroundColor Gray
    
    # LSV alert first if present
    if ($lsvAlert) {
        Write-Host "  [WARNING] $($lsvAlert.Component)" -ForegroundColor Yellow
        Write-Host "    Reason: $($lsvAlert.Reason)" -ForegroundColor White
        Write-Host "    Current Status: $($lsvAlert.Status)" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Datasource alerts
    foreach ($alert in $alerts) {
        $severityColor = if ($alert.Severity -eq "CRITICAL") { "Red" } else { "Yellow" }
        
        Write-Host ""
        Write-Host "  [$($alert.Severity)] $($alert.Datasource)" -ForegroundColor $severityColor
        Write-Host "    Reason: $($alert.Reason)" -ForegroundColor White
        Write-Host "    Pattern: $($alert.Pattern)" -ForegroundColor Gray
        Write-Host "    Last Session: $($alert.LastSessionStart) | State: $($alert.LastSessionState)" -ForegroundColor Gray
        if ($alert.LastError -and $alert.LastError -ne "N/A") {
            Write-Host "    Last Error: $($alert.LastError)" -ForegroundColor DarkYellow
        }
    }
    Write-Host ""
    Write-Host ("=" * 120) -ForegroundColor Gray
}

# Generate PSA-friendly output object
if ($OutputJson -or ($alerts.Count -gt 0 -or $lsvAlert)) {
    $psaOutput = [PSCustomObject]@{
        DeviceName = $deviceName
        CustomerName = $customerName
        MachineName = $machineName
        DeviceType = $deviceType
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ExitCode = 0  # Will be updated below
        Status = "OK"
        TotalDatasources = $datasourceGroups.Count
        HealthyDatasources = $healthyCount
        DatasourcesWithAlerts = $warningCount
        CriticalIssues = @()
        WarningIssues = @()
        LSVEnabled = $lsvEnabled
        LSVStatus = if ($lsvEnabled) { $lsvStatus } else { "N/A" }
        Summary = ""
    }
    
    # Add LSV alert if present
    if ($lsvAlert) {
        $psaOutput.WarningIssues += [PSCustomObject]@{
            Type = "LocalSpeedVault"
            Component = $lsvAlert.Component
            Status = $lsvAlert.Status
            Reason = $lsvAlert.Reason
            Severity = $lsvAlert.Severity
        }
    }
    
    # Add datasource alerts
    foreach ($alert in $alerts) {
        $alertObj = [PSCustomObject]@{
            Type = "Datasource"
            Datasource = $alert.Datasource
            Pattern = $alert.Pattern
            DaysSinceSuccess = $alert.DaysSinceSuccess
            ConsecutiveFailures = $alert.ConsecutiveFailures
            Reason = $alert.Reason
            Severity = $alert.Severity
            LastSessionStart = $alert.LastSessionStart
            LastSessionState = $alert.LastSessionState
            LastSessionEnd = $alert.LastSessionEnd
            LastError = $alert.LastError
        }
        
        if ($alert.Severity -eq "CRITICAL") {
            $psaOutput.CriticalIssues += $alertObj
        } else {
            $psaOutput.WarningIssues += $alertObj
        }
    }
    
    # Determine exit code and status
    if ($psaOutput.CriticalIssues.Count -gt 0) {
        $psaOutput.ExitCode = 2
        $psaOutput.Status = "CRITICAL"
        $psaOutput.Summary = "$($psaOutput.CriticalIssues.Count) critical issue(s) detected"
    }
    elseif ($psaOutput.WarningIssues.Count -gt 0) {
        $psaOutput.ExitCode = 1
        $psaOutput.Status = "WARNING"
        $psaOutput.Summary = "$($psaOutput.WarningIssues.Count) warning(s) detected"
    }
    else {
        $psaOutput.ExitCode = 0
        $psaOutput.Status = "OK"
        $psaOutput.Summary = "All $($psaOutput.TotalDatasources) datasource(s) healthy"
    }
    
    # Output JSON for PSA consumption
    if ($OutputJson) {
        Write-Host "`n" -NoNewline
        Write-Host "PSA Output (JSON):" -ForegroundColor Cyan
        Write-Host ("=" * 120) -ForegroundColor Gray
        $psaOutput | ConvertTo-Json -Depth 10
        Write-Host ("=" * 120) -ForegroundColor Gray
        Write-Host ""
    }
    
    # Dump failure JSON if requested (only when failures exist)
    if ($DumpFailureJson -and ($psaOutput.CriticalIssues.Count -gt 0 -or $psaOutput.WarningIssues.Count -gt 0)) {
        $jsonOutput = $psaOutput | ConvertTo-Json -Depth 10
        Write-Output $jsonOutput
    }
}

# Execution time
$scriptEndTime = Get-Date
$elapsedTime = $scriptEndTime - $scriptStartTime
Write-Host "Execution completed in $([math]::Round($elapsedTime.TotalSeconds, 2)) seconds" -ForegroundColor Gray
Write-Host ""

# Exit with appropriate code
if ($alerts | Where-Object { $_.Severity -eq "CRITICAL" }) {
    Write-Host "Exiting with code 2 (CRITICAL)" -ForegroundColor Red
    exit 2
}
elseif ($warningCount -gt 0 -or $lsvAlert) {
    Write-Host "Exiting with code 1 (WARNING)" -ForegroundColor Yellow
    exit 1
}
else {
    Write-Host "Exiting with code 0 (SUCCESS)" -ForegroundColor Green
    exit 0
}

#endregion
