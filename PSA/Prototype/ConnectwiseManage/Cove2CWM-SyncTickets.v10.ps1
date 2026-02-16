<# ----- About: ----
    # N-able Cove Data Protection Monitoring with ConnectWise Manage Ticket Integration
    # Revision v10.1 - 2026-02-16 - Authentication & Credential Handling Fixes
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # https://github.com/backupNerd
    #
    # v10.1 Changes (2026-02-16):
    #   - Fixed Set-APICredentials to use XML format (Export-Clixml) instead of text file format
    #   - Added try/catch error handling for corrupted credential files with auto-recovery
    #   - Enhanced Send-APICredentialsCookie with detailed error messages (API errors vs network errors)
    #   - Fixed case sensitivity in authentication visa token validation
    #   - Improved debug output for authentication troubleshooting
    #
    # v10 Changes (2026-01-01):
    #   - Added EndCustomer → CWM Company cache to eliminate redundant lookups for same customer
    #   - Fixed company creation timing - delay moved BEFORE CWM query (was after, causing ticket creation failures)
    #   - Enhanced placeholder pattern with wait logic for concurrent company creation
    #   - Expected performance: Eliminates duplicate CWM API calls when multiple devices share same EndCustomer
    #
    # v09 Changes:
    #   - Added status filter to query only failed/error sessions (Status==2 or Status==8)
    #   - Eliminated duplicate QueryErrors API call - saves 50% of error queries
    #   - Reduced session query range from 30 to 5 sessions
    #
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
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials for Cove Data Protection API
    # Check/ Get/ Store secure credentials for ConnectWise Manage API
    # Credentials are stored using Windows DPAPI encryption and can only be decrypted by the same user account on the same machine where created.
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Get Device Statistics and Last Session Status
    # Identify backup failures, warnings, and stale backups
    # Create or update ConnectWise Manage Service Tickets for issues
    # Close tickets automatically when issues are resolved
    #
    # Use the -DaysBack parameter to define how many days to look back for backup failures (default=30)
    #   Devices with a Timestamp older than -DaysBack days are presumed orpahaned or obsolate and this will be ignored regarless of error status
    # Use the -DeviceCount parameter to define the maximum number of devices to query (default=5000)
    # Use the -PartnerName parameter to specify exact Cove partner/customer name to monitor (e.g., "MYPartner Inc (bob@myparter.com)") (optional - defaults to authenticated partner)
    # Use the -TestDeviceName parameter to filter to a single device for testing (e.g., "desktop-pmb")
    # Use the -MonitorSystems parameter to enable/disable monitoring of servers and workstations (default=$true)
    # Use the -StaleHoursServers parameter to define hours since last successful backup before considering stale for servers (default=26)
    # Use the -StaleHoursWorkstations parameter to define hours since last successful backup before considering stale for workstations (default=240)
    # 
    # Use the -MonitorM365 parameter to enable/disable monitoring of Microsoft 365 tenants (default=$true)
    # Use the -StaleHoursM365 parameter to define hours since last successful backup before considering stale for M365 tenants (default=12)
    #
    # Use the -CreateTickets parameter to create tickets in ConnectWise Manage (default=$true, set to $false for test mode)
    # Use the -UpdateTickets parameter to update existing open tickets with new information (default=$true, set to $false to skip updates)
    # Use the -CloseResolvedTickets parameter to close tickets when issues are resolved (default=$true, set to $false to keep tickets open)
    # Use the -AutoCreateCompanies parameter to auto-create missing companies in ConnectWise Manage (default=$true, set to $false to disable)
    # Use the -UpdateCoveReferences parameter to append Cove PartnerReference/ExternalCode with CWM Company ID (default=$false, set to $true to enable)
    # Use the -UseDevicePartner parameter to create tickets at device partner level instead of End Customer (default=$false) - PROTOTYPE
    # Use the -UseLocalTime parameter to display timestamps in local time instead of UTC (default=$true)
    #
    # Use the -TicketBoard parameter to specify the ConnectWise Service Board name (default="Service Desk")
    # Use the -TicketType parameter to specify ticket type (default="ServiceTicket")
    # Use the -TicketStatus parameter to specify new ticket status (default="New Support Issue")
    # Use the -TicketPriorityServer parameter to specify ticket priority for servers (default="Priority 1 - Emergency Response")
    # Use the -TicketPriorityWorkstation parameter to specify ticket priority for workstations (default="Priority 3 - Normal Response")
    # Use the -TicketPriorityM365 parameter to specify ticket priority for M365 tenants (default="Priority 2 - Quick Response")
    # Use the -TicketClosedStatus parameter to specify closed ticket status (default="Closed")
    # Use the -TicketCompany parameter to specify the default company for tickets (optional) - PROTOTYPE
    #
    # Use the -ExportPath parameter to specify CSV file path for results (default=$PSScriptRoot)
    # Use the -ClearCDPCredentials parameter to remove stored Cove API credentials (default=$false)
    # Use the -ClearCWMCredentials parameter to remove stored ConnectWise Manage API credentials (default=$false)
    # Use the -AllowInsecureSSL parameter to bypass SSL certificate validation (default=$true, staging/dev environments only)
    # Use the -CleanupCoveTickets parameter to delete all Cove-created tickets from last 24 hours (default=$false, WARNING: Cannot be undone!)
    # Use the -TestMode parameter to simulate 50% of devices as successful for ticket closure testing (default=$false)
    # Use the -DebugCDP parameter to display debug info for Cove Data Protection queries (default=$false)
    # Use the -DebugCWM parameter to display debug info for ConnectWise Manage operations (default=$false)
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    #
    # Special thanks to Chris Taylor for the PS Module for ConnectWise Manage.
    # https://www.powershellgallery.com/packages/ConnectWiseManageAPI
    # https://github.com/christaylorcodes/ConnectWiseManageAPI

# -----------------------------------------------------------#>  ## Behavior

<#
.SYNOPSIS
Sync N-able Cove Data Protection backup issues to ConnectWise Manage tickets.

.DESCRIPTION
Monitors Cove backup devices (servers, workstations, M365 tenants) for failures, errors, and stale backups.
Automatically creates ConnectWise Manage tickets for issues and closes them when resolved.

.PARAMETER DaysBack
Number of days to look back for backup failures. Devices with timestamps older than this are considered orphaned/obsolete and ignored regardless of status. Default: 30 days.

.PARAMETER DeviceCount
Maximum number of devices to query from Cove API. Use lower values for testing. Default: 20 devices.

.PARAMETER PartnerName
Exact Cove partner/customer name to monitor (e.g., "MyPartner Inc (bob@mypartner.com)"). If omitted, defaults to authenticated partner or prompts for selection.

.PARAMETER TestDeviceName
Filter to a single device for testing purposes (e.g., "desktop-ph5hqmb"). Useful for development and troubleshooting.

.PARAMETER MonitorSystems
Enable monitoring of servers and workstations. Set to $false to skip system monitoring. Default: $true.

.PARAMETER StaleHoursServers
Hours since last successful backup before a server is considered stale. Default: 26 hours.

.PARAMETER StaleHoursWorkstations
Hours since last successful backup before a workstation is considered stale. Default: 240 hours (10 days).

.PARAMETER MonitorM365
Enable monitoring of Microsoft 365 tenants (Exchange, OneDrive, SharePoint, Teams). Set to $false to skip M365. Default: $true.

.PARAMETER StaleHoursM365
Hours since last successful backup before an M365 tenant is considered stale. Default: 12 hours.

.PARAMETER CreateTickets
Create new tickets in ConnectWise Manage for detected issues. Set to $false for test mode (no tickets created). Default: $true.

.PARAMETER UpdateTickets
Update existing open tickets with new information when device status changes. Set to $false to skip updates. Default: $true.

.PARAMETER CloseResolvedTickets
Automatically close tickets when issues are resolved (successful backup detected). Set to $false to keep tickets open. Default: $true.

.PARAMETER AutoCreateCompanies
Automatically create missing companies in ConnectWise Manage when Cove partner has no CWM match. Set to $false to disable auto-creation. Default: $true.

.PARAMETER UpdateCoveReferences
Append Cove PartnerReference/ExternalCode field with ConnectWise Company ID for future lookups. Set to $true to enable cross-referencing. Default: $false.

.PARAMETER UseDevicePartner
Create tickets at device partner level instead of End Customer level. PROTOTYPE feature. Default: $false.

.PARAMETER UseLocalTime
Display timestamps in local time zone instead of UTC. Default: $true.

.PARAMETER TicketBoard
ConnectWise Service Board name where tickets will be created. Default: "Service Desk".

.PARAMETER TicketType
ConnectWise ticket type identifier. Default: "ServiceTicket".

.PARAMETER TicketStatus
Status to assign to newly created tickets. Default: "New Support Issue".

.PARAMETER TicketPriorityServer
Priority level for server backup failure tickets. Default: "Priority 1 - Emergency Response".

.PARAMETER TicketPriorityWorkstation
Priority level for workstation backup failure tickets. Default: "Priority 3 - Normal Response".

.PARAMETER TicketPriorityM365
Priority level for Microsoft 365 backup failure tickets. Default: "Priority 2 - Quick Response".

.PARAMETER TicketClosedStatus
Status to set when closing resolved tickets. Default: "Closed".

.PARAMETER TicketCompany
Default company identifier for tickets. PROTOTYPE feature. Leave empty for auto-detection.

.PARAMETER ExportPath
Directory path for CSV export files. Default: Script directory ($PSScriptRoot).

.PARAMETER ClearCDPCredentials
Remove stored Cove Data Protection API credentials from encrypted credential file. Forces re-authentication on next run.

.PARAMETER ClearCWMCredentials
Remove stored ConnectWise Manage API credentials from encrypted credential file. Forces re-authentication on next run.

.PARAMETER AllowInsecureSSL
Bypass SSL certificate validation. Only use for staging/development environments. Default: $true.

.PARAMETER CleanupCoveTickets
Delete ALL Cove-created tickets from the last 24 hours. WARNING: Cannot be undone! Use for testing cleanup only.

.PARAMETER TestMode
Simulate 50% of devices as successful for testing ticket closure logic. Does not affect real device status.

.PARAMETER DebugCDP
Enable verbose debug output for Cove Data Protection API queries. Shows API calls, filters, and response data.

.PARAMETER DebugCWM
Enable verbose debug output for ConnectWise Manage operations. Shows ticket creation, updates, and company lookups.

.EXAMPLE
.\Cove2CWM-SyncTickets.v10.ps1
Run with default settings - monitor all devices, create/update/close tickets.

.EXAMPLE
.\Cove2CWM-SyncTickets.v10.ps1 -PartnerName "Acme Corp (admin@acme.com)" -DaysBack 7
Monitor specific partner for issues in the last 7 days.

.EXAMPLE
.\Cove2CWM-SyncTickets.v10.ps1 -TestDeviceName "srv-backup01" -CreateTickets:$false
Test mode - monitor single device, don't create tickets.

.EXAMPLE
.\Cove2CWM-SyncTickets.v10.ps1 -MonitorSystems:$false -MonitorM365:$true -StaleHoursM365 24
Monitor only M365 tenants with 24-hour stale threshold.

.NOTES
Requires PowerShell 7.0 or later.
Credentials stored using Windows DPAPI encryption (user/machine specific).
ConnectWiseManageAPI module required: Install-Module -Name ConnectWiseManageAPI

.LINK
https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm

.LINK
https://github.com/christaylorcodes/ConnectWiseManageAPI
#>

[CmdletBinding()]
Param (
    [int]$DaysBack = 30,                                         ## Days to look back for backup failures
    [int]$DeviceCount = 20,                                      ## Maximum number of devices to query
    [string]$PartnerName = "SEDEMO",                                   ## Exact Cove partner/customer name to monitor (optional - defaults to authenticated partner)
    [string]$TestDeviceName = "",                                ## Filter to single device for testing (e.g., "desktop-ph5hqmb")

    # Servers/Workstations Monitoring
    [bool]$MonitorSystems = $true,                               ## Monitor servers and workstations
    [int]$StaleHoursServers = 26,                                ## Hours since last backup to consider stale (servers)
    [int]$StaleHoursWorkstations = 72,                          ## Hours since last backup to consider stale (workstations)
    
    # M365 Monitoring
    [bool]$MonitorM365 = $true,                                  ## Monitor Microsoft 365 tenants
    [int]$StaleHoursM365 = 2,                                   ## Hours since last backup to consider stale (M365 tenants)

    # Ticketing Options
    [bool]$CreateTickets = $true,                                ## Set $true to create tickets in ConnectWise
    [bool]$UpdateTickets = $true,                                ## Set $true to update existing tickets
    [bool]$CloseResolvedTickets = $true,                         ## Set $true to close resolved tickets
    [bool]$AutoCreateCompanies = $false,                          ## Set $true to auto-create missing companies in ConnectWise
    [bool]$UpdateCoveReferences = $false,                        ## Set $true to update Cove partner ExternalCode with CWM Company ID
    [bool]$UseDevicePartner = $false,                            ## Set $true to create tickets at device partner level (not End Customer) - PROTOTYPE
    [bool]$UseLocalTime = $true,                                 ## Display timestamps in local time instead of UTC

    # ConnectWise Ticket Settings
    [string]$TicketBoard = "Service Desk",                       ## ConnectWise Service Board name
    [string]$TicketType = "ServiceTicket",                       ## Ticket type
    [string]$TicketStatus = "New Support Issue",                 ## New ticket status
    [string]$TicketPriorityServer = "Priority 1 - Emergency Response",          ## Ticket priority for servers
    [string]$TicketPriorityWorkstation = "Priority 3 - Normal Response",        ## Ticket priority for workstations
    [string]$TicketPriorityM365 = "Priority 2 - Quick Response",                ## Ticket priority for M365 tenants
    [string]$TicketClosedStatus = "Closed",                      ## Closed ticket status
    [string]$TicketCompany = "",                                 ## Default company for tickets - PROTOTYPE
    [string]$ExportPath = "$PSScriptRoot",                       ## Export Path for CSV

    [bool]$ClearCDPCredentials = $false,                         ## Remove Stored Cove API Credentials
    [bool]$ClearCWMCredentials = $false,                         ## Remove Stored CWM API Credentials
    
    # Testing Commands
    [bool]$AllowInsecureSSL = $true,                             ## Bypass SSL certificate validation (for staging/dev)
    [bool]$CleanupCoveTickets = $false,                          ## Delete all Cove-created tickets from last 24 hours (WARNING: Cannot be undone!)
    [bool]$TestMode = $false,                                    ## Simulate 50% ticket closures for testing
    [bool]$DebugCDP = $false,                                    ## Enable Debug for Cove Data Protection
    [bool]$DebugCWM = $false                                      ## Enable Debug for ConnectWise Manage
)

#Requires -Version 7.0

# PowerShell 7 version check with helpful error message
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "`n❌ ERROR: This script requires PowerShell 7 or later" -ForegroundColor Red
    Write-Host "`nYour current version: PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "`nTo install PowerShell 7:" -ForegroundColor Cyan
    Write-Host "  • Download: https://aka.ms/powershell" -ForegroundColor White
    Write-Host "  • Or run: winget install Microsoft.PowerShell" -ForegroundColor White
    Write-Host "`nAfter installing, launch PowerShell 7 with: pwsh`n" -ForegroundColor Cyan
    exit 1
}

Clear-Host
$Script:ScriptVersion = "v10"
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray

#region ----- Variable Cleanup (Prevent Cross-Contamination from Previous Runs) ----
# Clear critical variables that could cause wrong customer selection if script is re-run in same session
Write-Verbose "Clearing script-scoped variables to prevent reuse from previous runs..."

# Partner/Customer Selection Variables
# CRITICAL v09: Explicitly set to $null to clear cached values from previous runs
$Script:SelectedPartners = $null
$Script:Partner = $null
$Script:Partnerslist = $null
$Script:EnumeratePartnersSession = $null
# CRITICAL v09: Initialize as NEW empty hashtable (not $null) to prevent cached hierarchy data
$Script:PartnerHierarchyCache = @{}

# API Authentication Variables
Remove-Variable -Name visa -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name websession -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name VisaTime -Scope Script -ErrorAction SilentlyContinue

# Device/Data Collection Variables
Remove-Variable -Name DeviceDetail -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name AllDevices -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name SelectedDevices -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name FilteredDevices -Scope Script -ErrorAction SilentlyContinue

# CWM Variables
Remove-Variable -Name CWMCompanies -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name CWMTickets -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name CWMBoards -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name CWMServer -Scope Script -ErrorAction SilentlyContinue
Remove-Variable -Name CWMCompany -Scope Script -ErrorAction SilentlyContinue

# Clear ConnectWiseManageAPI module cache (if module is loaded)
if (Get-Module -Name ConnectWiseManageAPI) {
    Write-Verbose "Clearing ConnectWiseManageAPI module cache..."
    # Disconnect any existing CWM session first to ensure clean state
    try {
        Disconnect-CWM -ErrorAction SilentlyContinue
    } catch {
        # Ignore errors if not connected
    }
    # The module may cache API responses - force it to clear by removing and re-importing
    Remove-Module -Name ConnectWiseManageAPI -Force -ErrorAction SilentlyContinue
}

Write-Verbose "Variable cleanup complete"
#endregion ----- Variable Cleanup ----

#region ----- Script Location Change ----
    # CRITICAL: Change to script directory for relative path operations
    # This ensures $PSScriptRoot-based paths work correctly regardless of where script is launched from
    if ($PSScriptRoot) {
        Set-Location $PSScriptRoot
    }
#endregion

if ($DebugCDP) { Write-Host "[DEBUG] Parameter received: PartnerName='$PartnerName'" -ForegroundColor Magenta }
if ($DebugCDP) { Write-Host "[DEBUG] After init: Script:PartnerName='$Script:PartnerName', Script:OriginalPartnerName='$Script:OriginalPartnerName'" -ForegroundColor Magenta }

# OPTIMIZATION v10: EndCustomer → CWM Company cache (eliminates redundant lookups for same customer)
$Script:EndCustomerToCWMCompanyCache = @{}

# OPTIMIZATION v10: EndCustomer cache performance tracking
$Script:EndCustomerCacheStats = @{
    TotalLookups = 0
    CacheHits = 0
    CacheMisses = 0
    UniqueCustomersCached = New-Object 'System.Collections.Generic.HashSet[string]'
}

# OPTIMIZATION v10: Failed ticket operation retry queue
$Script:FailedTicketOperations = @()

# SSL Certificate Error Debug Log
$sslDebugLogPath = Join-Path $PSScriptRoot "SSL_Certificate_Errors_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

#region ----- Environment, Variables, Names and Paths ----


    
    # Change to script directory
    if ($PSScriptRoot) {
        Set-Location -Path $PSScriptRoot
    }
    
    $scriptpath = $PSCommandPath
    
    Write-Output "  Cove Data Protection Monitoring with ConnectWise Manage Ticket Integration`n"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "  Script Parameter Syntax:"
    Write-Output "  $Syntax"
    Write-Output "  Executing Script: $scriptpath`n"

    # Display monitoring rules summary
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                  MONITORING RULES SUMMARY                     ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    Write-Host "  Lookback Period                  : $DaysBack days" -ForegroundColor White
    if ($MonitorSystems) {
        Write-Host "  Servers and Workstations         : " -ForegroundColor White -NoNewline
        Write-Host "ENABLED" -ForegroundColor Green
        Write-Host "    - Server Stale Threshold       : $StaleHoursServers hours" -ForegroundColor Gray
        Write-Host "    - Workstation Stale Threshold  : $StaleHoursWorkstations hours" -ForegroundColor Gray
    } else {
        Write-Host "  Servers and Workstations         : " -ForegroundColor White -NoNewline
        Write-Host "DISABLED" -ForegroundColor Yellow
    }
    if ($MonitorM365) {
        Write-Host "  M365 Tenants                     : " -ForegroundColor White -NoNewline
        Write-Host "ENABLED" -ForegroundColor Green -NoNewline
        Write-Host " - Stale Threshold: $StaleHoursM365 hours" -ForegroundColor Gray
    } else {
        Write-Host "  M365 Tenants                     : " -ForegroundColor White -NoNewline
        Write-Host "DISABLED" -ForegroundColor Yellow
    }
    if ($PartnerName) {
        Write-Host "  Target Partner                   : $PartnerName" -ForegroundColor White
    }
    
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                   TICKET ACTION RULES                         ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    Write-Host "  Create New Tickets               : " -ForegroundColor White -NoNewline
    Write-Host "$(if ($CreateTickets) { 'ENABLED' } else { 'DISABLED (Test Mode)' })" -ForegroundColor $(if ($CreateTickets) { 'Green' } else { 'Yellow' })
    
    Write-Host "  Update Existing Tickets          : " -ForegroundColor White -NoNewline
    Write-Host "$(if ($UpdateTickets) { 'ENABLED' } else { 'DISABLED' })" -ForegroundColor $(if ($UpdateTickets) { 'Green' } else { 'Yellow' })
    
    Write-Host "  Close Resolved Tickets           : " -ForegroundColor White -NoNewline
    Write-Host "$(if ($CloseResolvedTickets) { 'ENABLED' } else { 'DISABLED' })" -ForegroundColor $(if ($CloseResolvedTickets) { 'Green' } else { 'Yellow' })
    
    Write-Host "  Auto-Create Companies            : " -ForegroundColor White -NoNewline
    Write-Host "$(if ($AutoCreateCompanies) { 'ENABLED' } else { 'DISABLED' })" -ForegroundColor $(if ($AutoCreateCompanies) { 'Green' } else { 'Yellow' })
    
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                   CONNECTWISE SETTINGS                        ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    Write-Host "  Service Board                    : $TicketBoard" -ForegroundColor White
    Write-Host "  New Ticket Status                : $TicketStatus" -ForegroundColor White
    Write-Host "  Closed Ticket Status             : $TicketClosedStatus" -ForegroundColor White
    if ($AllowInsecureSSL) {
        Write-Host "  SSL Validation                   : " -ForegroundColor White -NoNewline
        Write-Host "BYPASSED (Staging Mode)" -ForegroundColor Yellow
    }
    if ($TestMode) {
        Write-Host "  Test Mode                        : " -ForegroundColor White -NoNewline
        Write-Host 'ENABLED (50% Ticket Close Simulation)' -ForegroundColor Magenta
    }
    
    # Check for required external script if AutoCreateCompanies is enabled
    if ($AutoCreateCompanies) {
        $syncScriptPath = Join-Path $PSScriptRoot "Cove2CWM-SyncCustomers.v10.ps1"
        
        if (-not (Test-Path $syncScriptPath)) {
            # Try to find any version of the script
            $allVersions = Get-ChildItem -Path $PSScriptRoot -Filter "Cove2CWM-SyncCustomers*.ps1" -File | Sort-Object Name -Descending
            
            if ($allVersions.Count -eq 1) {
                # Only one version found, use it automatically
                $syncScriptPath = $allVersions[0].FullName
                Write-Host "  Note: Using $($allVersions[0].Name) (v10 not found)" -ForegroundColor Gray
            }
            elseif ($allVersions.Count -gt 1) {
                # Multiple versions found, use latest (first after descending sort)
                $syncScriptPath = $allVersions[0].FullName
                Write-Host "  Note: Using $($allVersions[0].Name) (latest version found)" -ForegroundColor Gray
            }
        }
        
        if (-not (Test-Path $syncScriptPath)) {
            Write-Host ""
            Write-Host "  ERROR: Auto-Create Companies is ENABLED but required script is missing:" -ForegroundColor Red
            Write-Host "         Cove2CWM-SyncCustomers.ps1 (any version)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  This script is required to automatically create missing companies in ConnectWise." -ForegroundColor Yellow
            Write-Host "  Please either:" -ForegroundColor Yellow
            Write-Host "    1. Place 'Cove2CWM-SyncCustomers.v10.ps1' in the same folder as this script, OR" -ForegroundColor White
            Write-Host "    2. Run with -AutoCreateCompanies `$false to disable this feature" -ForegroundColor White
            Write-Host ""
            
            $response = Read-Host "  Continue anyway? Companies will NOT be auto-created. (Y/N)"
            if ($response -ne 'Y' -and $response -ne 'y') {
                Write-Host "  Script execution cancelled." -ForegroundColor Yellow
                exit 1
            }
            
            Write-Host "  Continuing with Auto-Create Companies DISABLED..." -ForegroundColor Yellow
            $AutoCreateCompanies = $false
        }
    }
    
    # Validate required ConnectWise ticket parameters
    $missingParams = @()
    if ([string]::IsNullOrWhiteSpace($TicketBoard)) { $missingParams += 'TicketBoard' }
    if ([string]::IsNullOrWhiteSpace($TicketPriorityServer)) { $missingParams += 'TicketPriorityServer' }
    if ([string]::IsNullOrWhiteSpace($TicketPriorityWorkstation)) { $missingParams += 'TicketPriorityWorkstation' }
    if ([string]::IsNullOrWhiteSpace($TicketPriorityM365)) { $missingParams += 'TicketPriorityM365' }
    if ([string]::IsNullOrWhiteSpace($TicketStatus)) { $missingParams += 'TicketStatus' }
    if ([string]::IsNullOrWhiteSpace($TicketClosedStatus)) { $missingParams += 'TicketClosedStatus' }
    
    if ($missingParams.Count -gt 0) {
        Write-Host ""
        Write-Host "  WARNING: Required ConnectWise ticket parameters are not defined:" -ForegroundColor Yellow
        foreach ($param in $missingParams) {
            Write-Host "           -$param" -ForegroundColor Red
        }
        Write-Host ""
        
        # Check for helper script - try specific version first, then find latest version
        $optionsScriptPath = Join-Path $PSScriptRoot "Cove2CWM-SetTicketsConfig.v10.ps1"
        
        if (-not (Test-Path $optionsScriptPath)) {
            # Try to find any version of the script
            $allVersions = Get-ChildItem -Path $PSScriptRoot -Filter "Cove2CWM-SetTicketsConfig*.ps1" -File | Sort-Object Name -Descending
            
            if ($allVersions.Count -eq 1) {
                # Only one version found, use it automatically
                $optionsScriptPath = $allVersions[0].FullName
                Write-Host "  Note: Using $($allVersions[0].Name) (v10 not found)" -ForegroundColor Gray
            }
            elseif ($allVersions.Count -gt 1) {
                # Multiple versions found, prompt user
                Write-Host "  Multiple versions of SetTicketsConfig found:" -ForegroundColor Yellow
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "    $($i+1). $($allVersions[$i].Name)" -ForegroundColor White
                }
                
                $selection = Read-Host "  Select version to use (1-$($allVersions.Count)) or press Enter to skip"
                if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $allVersions.Count) {
                    $optionsScriptPath = $allVersions[[int]$selection - 1].FullName
                }
                else {
                    $optionsScriptPath = $null
                }
            }
            else {
                $optionsScriptPath = $null
            }
        }
        
        if ($optionsScriptPath -and (Test-Path $optionsScriptPath)) {
            $scriptName = Split-Path $optionsScriptPath -Leaf
            Write-Host "  A helper script is available to enumerate valid options from your ConnectWise instance:" -ForegroundColor Cyan
            Write-Host "    .\$scriptName" -ForegroundColor White
            Write-Host ""
            Write-Host "  This script will show you:" -ForegroundColor Yellow
            Write-Host "    - Available Service Boards" -ForegroundColor White
            Write-Host "    - Valid Status names" -ForegroundColor White
            Write-Host "    - Valid Priority names" -ForegroundColor White
            Write-Host ""
            
            $response = Read-Host "  Run Cove2CWM-SetTicketsConfig.ps1 now to see valid options? (Y/N)"
            if ($response -eq 'Y' -or $response -eq 'y') {
                Write-Host ""
                Write-Host "  Launching Cove2CWM-SetTicketsConfig.ps1..." -ForegroundColor Cyan
                Write-Host ""
                
                # Pass the monitoring script path so the options script can update parameters
                $monitoringScriptPath = $PSCommandPath
                & $optionsScriptPath -MonitoringScriptPath $monitoringScriptPath
                
                Write-Host ""
                Write-Host "  Checking if parameters were updated in the script..." -ForegroundColor Cyan
                
                # Re-read the script to check if parameters were updated
                $updatedContent = Get-Content $monitoringScriptPath -Raw
                $boardMatch = if ($updatedContent -match '\$TicketBoard\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                $statusMatch = if ($updatedContent -match '\$TicketStatus\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                $priorityServerMatch = if ($updatedContent -match '\$TicketPriorityServer\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                $priorityWorkstationMatch = if ($updatedContent -match '\$TicketPriorityWorkstation\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                $priorityM365Match = if ($updatedContent -match '\$TicketPriorityM365\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                $closedMatch = if ($updatedContent -match '\$TicketClosedStatus\s*=\s*"(.*?)"') { $matches[1] } else { $null }
                
                $allUpdated = $boardMatch -and $statusMatch -and $priorityServerMatch -and $priorityWorkstationMatch -and $priorityM365Match -and $closedMatch
                
                if ($allUpdated) {
                    Write-Host "  Parameters were successfully updated in the script!" -ForegroundColor Green
                    Write-Host ""
                    Write-Host "  Updated values:" -ForegroundColor Cyan
                    Write-Host "    -TicketBoard               : $boardMatch" -ForegroundColor White
                    Write-Host "    -TicketStatus              : $statusMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityServer      : $priorityServerMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityWorkstation : $priorityWorkstationMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityM365        : $priorityM365Match" -ForegroundColor White
                    Write-Host "    -TicketClosedStatus        : $closedMatch" -ForegroundColor White
                    Write-Host ""
                    Write-Host "  Continuing with updated parameters..." -ForegroundColor Green
                    Write-Host ""
                    
                    # Update the current session variables with the new values
                    $TicketBoard = $boardMatch
                    $TicketStatus = $statusMatch
                    $TicketPriorityServer = $priorityServerMatch
                    $TicketPriorityWorkstation = $priorityWorkstationMatch
                    $TicketPriorityM365 = $priorityM365Match
                    $TicketClosedStatus = $closedMatch
                } else {
                    Write-Host "  Parameters were not updated. Exiting..." -ForegroundColor Yellow
                    Write-Host "  Please re-run this script after reviewing the options." -ForegroundColor Yellow
                    Write-Host ""
                    exit 1
                }
            }
        }
        
        if (-not $optionsScriptPath -or -not (Test-Path $optionsScriptPath)) {
            Write-Host "  Helper script not found: Cove2CWM-SetTicketsConfig.ps1" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  You need to set these parameters when running the script:" -ForegroundColor Yellow
            Write-Host "    -TicketBoard               : Name of your ConnectWise Service Board" -ForegroundColor White
            Write-Host "    -TicketStatus              : Status name for new tickets" -ForegroundColor White
            Write-Host "    -TicketPriorityServer      : Priority for server tickets" -ForegroundColor White
            Write-Host "    -TicketPriorityWorkstation : Priority for workstation tickets" -ForegroundColor White
            Write-Host "    -TicketPriorityM365        : Priority for M365 tickets" -ForegroundColor White
            Write-Host "    -TicketClosedStatus        : Status name for closed tickets" -ForegroundColor White
            Write-Host ""
            Write-Host "  To get valid values, you can:" -ForegroundColor Yellow
            Write-Host "    1. Log into ConnectWise Manage and check your Service Board settings, OR" -ForegroundColor White
            Write-Host "    2. Get the 'Cove2CWM-SetTicketsConfig.ps1' helper script from the repository" -ForegroundColor White
            Write-Host ""
        }
        
        $response = Read-Host "  Continue with default values anyway? (Y/N)"
        if ($response -ne 'Y' -and $response -ne 'y') {
            Write-Host "  Script execution cancelled." -ForegroundColor Yellow
            exit 1
        }
        Write-Host ""
    }
    
    Write-Host ""
    Write-Host ""

    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
    $Script:ScriptStartTime = Get-Date
    
    # PERFORMANCE TRACKING v07: Initialize comprehensive performance metrics
    $Script:PerformanceMetrics = @{
        DeviceQueryTime = 0
        HierarchyTime = 0
        CompanyLookupTime = 0
        CompanyLookupCount = 0  # Track actual API calls (not cache hits)
        TicketSearchTime = 0
        TicketSearchCount = 0   # Track actual API calls
        TicketCreateTime = 0
        TicketUpdateTime = 0
        TicketCloseTime = 0
        ParallelBatchCount = 0
        ParallelJobsRun = 0
    }
    
    # Reset tracking arrays to prevent accumulation across multiple script runs in same session
    $Script:AllIssues = @()
    $Script:TicketActions = @()
    $Script:CompanyLookupCache = @()  # Session cache to prevent duplicate CWM company lookups
    $Script:CompaniesCreatedViaHelper = @()  # Track actual new companies created by helper script
    $Script:ReferencesUpdated = @()
    
    $ticketsFolder = Join-Path $ExportPath "tickets"
    if (-not (Test-Path $ticketsFolder)) {
        New-Item -Path $ticketsFolder -ItemType Directory -Force | Out-Null
    }
    
    # Create timestamped subfolder for this script run's ticket exports
    $Script:TicketExportFolder = Join-Path $ticketsFolder $CurrentDate
    if (-not (Test-Path $Script:TicketExportFolder)) {
        New-Item -Path $Script:TicketExportFolder -ItemType Directory -Force | Out-Null
    }
    
    $Script:LogFile = "$ticketsFolder\$($CurrentDate)_CoveMonitoring_$($Script:ScriptVersion).csv"
    $Script:TicketLogFile = "$ticketsFolder\$($CurrentDate)_CWM_Tickets_$($Script:ScriptVersion).csv"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:True_path = "C:\ProgramData\MXB\"

    # Bypass SSL certificate validation if requested (for staging/dev environments)
    if ($AllowInsecureSSL) {
        Write-Warning "SSL certificate validation is disabled - use only for staging/dev environments!"
        
        # Use simpler approach - direct ServicePointManager manipulation
        try {
            $addTypeCode = @'
                using System.Net;
                using System.Security.Cryptography.X509Certificates;
                public class TrustAllCertsPolicy : ICertificatePolicy {
                    public bool CheckValidationResult(
                        ServicePoint srvPoint, X509Certificate certificate,
                        WebRequest request, int certificateProblem) {
                        return true;
                    }
                }
'@
            Add-Type -TypeDefinition $addTypeCode -ErrorAction SilentlyContinue
            [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        } catch {
            # Type already loaded, ignore error
        }
    }

    # Define issue severity levels (Priority determined dynamically based on device type)
    $Script:IssueSeverity = @{
        Critical = @{
            Color = "Red"
            Description = "Backup failed or critical errors"
        }
        Warning = @{
            Color = "Yellow"
            Description = "Backup completed with warnings"
        }
        Stale = @{
            Color = "Magenta"
            Description = "No recent successful backup"
        }
        Success = @{
            Color = "Green"
            Description = "Backup successful"
        }
    }
    
    # Track processed tickets to prevent duplicate operations in same session
    $Script:ProcessedTickets = @{
        Created = @()  # Track created tickets to prevent duplicates
        Updated = @()
        Closed = @()
    }

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Data Conversion Functions ----

Function Get-TicketPriorityForDevice {
    param(
        [Parameter(Mandatory=$true)][object]$Device
    )
    
    # Determine priority based on device type
    if ($Device.AccountType -eq 2) {
        # M365 Tenant
        return $Script:TicketPriorityM365
    } elseif ($Device.OSType -eq 2 -or $Device.DeviceType -like "*Server*") {
        # Server (OT == 2)
        return $Script:TicketPriorityServer
    } else {
        # Workstation (OT == 1 or default)
        return $Script:TicketPriorityWorkstation
    }
}

Function Get-SessionStatusText {
    param($StatusCode)
    
    # Convert to string and handle both string and int inputs
    $code = "$StatusCode"
    
    switch ($code) {
        "0" { return "Unknown" }
        "1" { return "InProcess" }
        "2" { return "Failed" }
        "3" { return "Completed with warnings" }
        "4" { return "Aborted" }
        "5" { return "Completed" }
        "6" { return "Interrupted" }
        "7" { return "NotStarted" }
        "8" { return "CompletedWithErrors" }
        "9" { return "Skipped" }
        "10" { return "Canceled" }
        "11" { return "SessionStopped" }
        default { return "Unknown" }
    }
}

Function Format-HoursAsRelativeTime {
    param(
        [double]$Hours,
        [switch]$NeverBackedUp,
        [switch]$IncludeParentheses = $true
    )
    
    # Build the base time string with abbreviated units (no spaces)
    $timeString = if ($Hours -lt 0.25) {
        "just now"
    } elseif ($Hours -lt 1) {
        $roundedMins = [math]::Floor([math]::Floor($Hours * 60) / 15) * 15
        "~$($roundedMins)m ago"
    } elseif ($Hours -lt 24) {
        $hrs = [math]::Floor($Hours)
        $roundedMins = [math]::Floor([math]::Floor(($Hours - $hrs) * 60) / 15) * 15
        if ($roundedMins -eq 0) {
            "~$($hrs)h ago"
        } else {
            "~$($hrs)h $($roundedMins)m ago"
        }
    } elseif ($Hours -lt 720) {  # Less than 30 days
        $days = [math]::Floor($Hours / 24)
        $remainingHours = [math]::Floor($Hours % 24)
        if ($remainingHours -eq 0) {
            "~$($days)d ago"
        } else {
            "~$($days)d $($remainingHours)h ago"
        }
    } elseif ($Hours -lt 8760) {  # Less than 1 year (365 days)
        # BUGFIX v10.3: Show months and weeks for more precision (e.g., "~1mo 3w ago")
        # Use 730.5 hours/month (365.25 days / 12 months)
        $months = [math]::Floor($Hours / 730.5)
        $remainingHours = $Hours - ($months * 730.5)
        $weeks = [math]::Floor($remainingHours / 168)  # 168 hours per week
        
        if ($weeks -eq 0) {
            "~$($months)mo ago"
        } else {
            "~$($months)mo $($weeks)w ago"
        }
    } else {
        $years = [math]::Floor($Hours / 8760)
        "~$($years)yr ago"
    }
    
    # Add context for never-backed-up devices
    if ($NeverBackedUp) {
        $timeString = "created $timeString - No backup"
    }
    
    # Add parentheses if requested (default for ticket display fields)
    if ($IncludeParentheses) {
        return "($timeString)"
    } else {
        return $timeString
    }
}

Function Format-TimezoneOffset {
    param(
        [string]$TzValue
    )
    
    # If empty or null, return empty string
    if ([string]::IsNullOrWhiteSpace($TzValue)) {
        return ""
    }
    
    # Try to parse as integer (API returns hours offset)
    try {
        $hours = [int]$TzValue
        
        # Format as UTC±HH:MM
        $sign = if ($hours -ge 0) { "+" } else { "-" }
        $absHours = [Math]::Abs($hours)
        return "UTC{0}{1:D2}:00" -f $sign, $absHours
    }
    catch {
        # If parsing fails, return original value
        return $TzValue
    }
}

# Helper function to format aligned labels with padding
Function Format-AlignedLabel {
    param(
        [string]$Label,
        [int]$MaxLength = 18  # Longest label is "Last Completed" (14) + buffer
    )
    $padding = $MaxLength - $Label.Length
    return "$Label$(' ' * $padding): "
}

# Helper function to clean JSON error messages (e.g., Graph API errors)
Function Format-CleanErrorMessage {
    param(
        [string]$ErrorMessage
    )
    
    if ([string]::IsNullOrWhiteSpace($ErrorMessage)) {
        return $ErrorMessage
    }
    
    # STEP 1: Clean up Graph API wrapper format
    # Pattern: "Graph API request failed: 404 (Not Found); Request: GET https://...; Request body size: 0; ReplyBodyError: { ... }"
    # This removes the HTTP status code, request URL, and body size - keeping only the actual error
    if ($ErrorMessage -match 'Graph API request failed.*ReplyBodyError:\s*(.+)$') {
        # Extract the ReplyBodyError JSON portion
        $ErrorMessage = $matches[1].Trim()
    }
    
    # STEP 2: Clean up embedded JSON format with unquoted property names
    # Pattern: { ErrorCode: "MailboxNotEnabledForRESTAPI", ErrorMessage: "The mailbox is either inactive..." }
    # Note: This is non-standard JSON (unquoted property names) so we use regex instead of JSON parser
    if ($ErrorMessage -match '\{\s*ErrorCode:\s*"[^"]*"\s*,\s*ErrorMessage:\s*"([^"]+)"\s*\}') {
        return $matches[1]
    }
    
    # STEP 3: Clean up additional Graph API error patterns
    # Pattern with error_description: {"error":"...", "error_description":"..."}
    if ($ErrorMessage -match '"error_description"\s*:\s*"([^"]+)"') {
        return $matches[1]
    }
    
    # Pattern with nested error object: {"error":{"code":"...","message":"..."}}
    if ($ErrorMessage -match '"error"\s*:\s*\{[^}]*"message"\s*:\s*"([^"]+)"') {
        return $matches[1]
    }
    
    # STEP 4: Try standard JSON parsing for properly formatted JSON
    # Only attempt if the message looks like valid JSON (starts with { and ends with })
    $trimmed = $ErrorMessage.Trim()
    if (-not ($trimmed.StartsWith('{') -and $trimmed.EndsWith('}'))) {
        # Not JSON format - return current message as-is
        return $ErrorMessage
    }
    
    # Looks like JSON - try to parse it
    try {
        $errorObj = $ErrorMessage | ConvertFrom-Json -ErrorAction Stop
        
        # Extract the meaningful error message from common JSON structures
        # Priority order: ErrorMessage -> errorMessage -> error_description -> message -> error.message
        if ($errorObj.PSObject.Properties['ErrorMessage'] -and $errorObj.ErrorMessage) {
            return $errorObj.ErrorMessage
        }
        elseif ($errorObj.PSObject.Properties['errorMessage'] -and $errorObj.errorMessage) {
            return $errorObj.errorMessage
        }
        elseif ($errorObj.PSObject.Properties['error_description'] -and $errorObj.error_description) {
            return $errorObj.error_description
        }
        elseif ($errorObj.PSObject.Properties['message'] -and $errorObj.message) {
            return $errorObj.message
        }
        elseif ($errorObj.error -and $errorObj.error.message) {
            return $errorObj.error.message
        }
        else {
            # JSON parsed but no recognizable message field - return original
            return $ErrorMessage
        }
    }
    catch {
        # JSON parsing failed (malformed JSON) - return original
        return $ErrorMessage
    }
}

Function Get-DatasourceStatus {
    param(
        [object]$DeviceSettings,
        [string]$DataSourceString
    )
    
    $statusList = @()
    
    # Files and Folders (D01 only - not D1 to avoid matching D19)
    if ($DataSourceString -match '\bD01\b') {
        $statusCode = $DeviceSettings.F0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "Files: $statusText"
    }
    
    # System State (D02 only - not D2 to avoid matching D20, D23)
    if ($DataSourceString -match '\bD02\b') {
        $statusCode = $DeviceSettings.S0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "SystemState: $statusText"
    }
    
    # Linux System State - DEPRECATED - Commented out 2025-12-26
    <#
    if ($DataSourceString -match 'LinuxSystemState') {
        $statusCode = $DeviceSettings.K0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "LinuxSystemState: $statusText"
    }
    #>
    
    # MS SQL (D10)
    if ($DataSourceString -match 'D10') {
        $statusCode = $DeviceSettings.Z0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "MSSQL: $statusText"
    }
    
    # M365 Exchange (D19)
    if ($DataSourceString -match 'D19') {
        $statusCode = $DeviceSettings.D19F0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "M365Exchange: $statusText"
    }
    
    # M365 OneDrive (D20)
    if ($DataSourceString -match 'D20') {
        $statusCode = $DeviceSettings.D20F0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "M365OneDrive: $statusText"
    }
    
    # M365 SharePoint (D05/D5)
    if ($DataSourceString -match 'D05|D5') {
        $statusCode = $DeviceSettings.D5F0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "M365SharePoint: $statusText"
    }
    
    # M365 Teams (D23)
    if ($DataSourceString -match 'D23') {
        $statusCode = $DeviceSettings.D23F0 -join ''
        $statusText = Get-SessionStatusText $statusCode
        $statusList += "M365Teams: $statusText"
    }
    
    if ($statusList.Count -gt 0) {
        return ($statusList -join ' | ')
    } else {
        return "No datasources"
    }
}

# Helper function to format data size with appropriate units
Function Format-DataSize {
    param([object]$BytesValue)
    
    if (-not $BytesValue) { return 'N/A' }
    $bytes = [decimal]($BytesValue -join '')
    
    if ($bytes -eq 0) { return '0 B' }
    
    if ($bytes -lt 1KB) {
        return "$([Math]::Round($bytes, 2)) B"
    } elseif ($bytes -lt 1MB) {
        return "$([Math]::Round($bytes / 1KB, 2)) KB"
    } elseif ($bytes -lt 1GB) {
        return "$([Math]::Round($bytes / 1MB, 2)) MB"
    } elseif ($bytes -lt 1TB) {
        return "$([Math]::Round($bytes / 1GB, 2)) GB"
    } else {
        return "$([Math]::Round($bytes / 1TB, 2)) TB"
    }
}

Function Get-DatasourceDetails {
    param(
        [object]$DeviceSettings,
        [string]$DataSourceString,
        [hashtable]$ErrorMessages = @{},
        [int]$AccountType = 1  # 1=Server/Workstation, 2=M365
    )
    
    # Helper function to safely convert timestamp with timezone awareness
    Function Get-SafeTimestamp {
        param(
            [object]$TimestampValue,
            [string]$TimezoneOffset = $null,
            [switch]$WithLabel
        )
        if (-not $TimestampValue) { return 'N/A' }
        $joined = $TimestampValue -join ''
        if ([string]::IsNullOrWhiteSpace($joined) -or $joined -eq '0') { return 'N/A' }
        
        $converted = Convert-UnixTimeToDateTime $joined
        if ($converted -eq 'N/A' -or [string]::IsNullOrWhiteSpace($converted)) { return 'N/A' }
        
        # The API returns timestamps in UTC - optionally convert to local time
        if ($WithLabel) {
            if ($UseLocalTime) {
                # Convert from UTC to local time
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($converted, [System.TimeZoneInfo]::Local)
                # Create abbreviation from timezone name (e.g., "Eastern Standard Time" -> "EST")
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                return "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else {
                # Keep UTC
                return "$($converted.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            }
        } else {
            if ($UseLocalTime) {
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($converted, [System.TimeZoneInfo]::Local)
                return $localTime.ToString('yyyy-MM-dd HH:mm:ss')
            } else {
                return $converted.ToString('yyyy-MM-dd HH:mm:ss')
            }
        }
    }
    
    # Wrapper function to convert Unix timestamp to relative time display
    # Converts timestamp to hours, then calls unified Format-HoursAsRelativeTime
    Function Get-TimeAgo {
        param(
            [object]$TimestampValue,
            [string]$TimezoneOffset = $null
        )
        if (-not $TimestampValue) { return '' }
        $joined = $TimestampValue -join ''
        if ([string]::IsNullOrWhiteSpace($joined) -or $joined -eq '0') { return '' }
        
        $converted = Convert-UnixTimeToDateTime $joined
        if ($converted -eq 'N/A' -or [string]::IsNullOrWhiteSpace($converted)) { return '' }
        
        # Calculate hours difference
        $now = (Get-Date).ToUniversalTime()
        $timespan = $now - $converted
        $hours = $timespan.TotalHours
        
        # Use unified function for formatting
        return Format-HoursAsRelativeTime -Hours $hours -IncludeParentheses:$true
    }
    
    # NOTE: Old Format-HoursAsRelativeTime function removed
    # Use unified Format-HoursAsRelativeTime function at line 524 instead
    
    # NOTE: Format-DataSize function moved to global scope (before Get-DatasourceDetails function)
    # This allows it to be called from both inside and outside Get-DatasourceDetails
    
    $datasourceLines = @()
    
    # Files and Folders (D01/D1) - Skip for M365 devices
    if ($AccountType -ne 2 -and $DataSourceString -match 'D01|D1') {
        $statusCode = $DeviceSettings.F0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.FJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.FL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.FL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.FO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.FO -TimezoneOffset $tz
        
        # Validate Last Success accuracy using FQ (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.FQ -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.FL -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.FA -and ($DeviceSettings.FA -join '')) { 
            $seconds = [int]($DeviceSettings.FA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.FO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.F3
        $processed = Format-DataSize $DeviceSettings.F4
        $sent = Format-DataSize $DeviceSettings.F5
        $errors = if ($DeviceSettings.F7) { $DeviceSettings.F7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('FileSystem'))) {
            $sessionId = $ErrorMessages['FileSystem'].SessionId
            $timestampFormatted = $ErrorMessages['FileSystem'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['FileSystem'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['FileSystem'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | FileSystem | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

File & Folders Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # System State (D02/D2) - Skip for M365 devices
    if ($AccountType -ne 2 -and $DataSourceString -match 'D02|D2') {
        $statusCode = $DeviceSettings.S0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.SJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.SL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.SL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.SO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.SO -TimezoneOffset $tz
        
        # Validate Last Success accuracy using SQ (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.SQ -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.SL -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.SA -and ($DeviceSettings.SA -join '')) { 
            $seconds = [int]($DeviceSettings.SA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.SO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.S3
        $processed = Format-DataSize $DeviceSettings.S4
        $sent = Format-DataSize $DeviceSettings.S5
        $errors = if ($DeviceSettings.S7) { $DeviceSettings.S7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('SystemState'))) {
            $sessionId = $ErrorMessages['SystemState'].SessionId
            $timestampFormatted = $ErrorMessages['SystemState'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['SystemState'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['SystemState'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | SystemState | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

System State Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # Linux System State - DEPRECATED - Commented out 2025-12-26
    <#
    if ($DataSourceString -match 'LinuxSystemState') {
        $statusCode = $DeviceSettings.K0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.KJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.KL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.KL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.KO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.KO -TimezoneOffset $tz
        
        # Validate last success status
        $lastSuccessStatus = $DeviceSettings.KQ -join ''
        
        # Show warning if last successful session wasn't actually successful
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.KL -join '') -gt 0) {
            $lastSuccessStatusText = Get-SessionStatusText $lastSuccessStatus
            $lastSuccess = "$lastSuccess (WARNING: Status was '$lastSuccessStatusText' not 'Completed')"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.KA -and ($DeviceSettings.KA -join '')) { 
            $seconds = [int]($DeviceSettings.KA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.KO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.K3
        $processed = Format-DataSize $DeviceSettings.K4
        $sent = Format-DataSize $DeviceSettings.K5
        $errors = if ($DeviceSettings.K7) { $DeviceSettings.K7 -join '' } else { '0' }
        
        $datasourceLines += @"

Linux System State Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors
"@
    }
    #>
    
    # MS SQL (D10)
    if ($DataSourceString -match 'D10') {
        $statusCode = $DeviceSettings.Z0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.ZJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.ZL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.ZL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.ZO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.ZO -TimezoneOffset $tz
        
        # Validate Last Success accuracy using ZQ (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.ZQ -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.ZL -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.ZA -and ($DeviceSettings.ZA -join '')) { 
            $seconds = [int]($DeviceSettings.ZA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.ZO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.Z3 -and ($DeviceSettings.Z3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Z3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.Z4 -and ($DeviceSettings.Z4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Z4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.Z5 -and ($DeviceSettings.Z5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Z5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.Z7) { $DeviceSettings.Z7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('MSSQL'))) {
            $sessionId = $ErrorMessages['MSSQL'].SessionId
            $timestampFormatted = $ErrorMessages['MSSQL'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['MSSQL'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['MSSQL'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | MSSQL | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

MS SQL Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # M365 Exchange (D19)
    if ($DataSourceString -match 'D19') {
        $statusCode = $DeviceSettings.D19F0 -join ''
        
        # Check if datasource is configured (has session start time)
        if ([string]::IsNullOrWhiteSpace($statusCode) -or $statusCode -eq '0') {
            $status = "Not Configured"
        } else {
            $status = Get-SessionStatusText $statusCode
            
            # If current status is InProcess, show last completed status -> InProcess
            if ($statusCode -eq "1") {  # InProcess
                $lastCompletedStatusCode = $DeviceSettings.GJ -join ''
                $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
                $status = "$lastCompletedStatus -> $status"
            }
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.GL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.GL
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.GO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.GO
        
        # Validate Last Success accuracy using GQ (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.GQ -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.GL -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.GA -and ($DeviceSettings.GA -join '')) { 
            $seconds = [int]($DeviceSettings.GA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.GO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.G3
        $processed = Format-DataSize $DeviceSettings.G4
        $sent = Format-DataSize $DeviceSettings.G5
        $errors = if ($DeviceSettings.G7) { $DeviceSettings.G7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('Exchange'))) {
            $sessionId = $ErrorMessages['Exchange'].SessionId
            $timestampFormatted = $ErrorMessages['Exchange'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['Exchange'].Status
            $cleanErrorMsg = $ErrorMessages['Exchange'].Message  # Already cleaned by Format-CleanErrorMessage
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | Exchange | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        # Check AutoAdd and Archive configuration
        $autoAdd = if ($DeviceSettings.AA3147 -and ($DeviceSettings.AA3147 -join '') -eq 'true') { 'Enabled' } else { 'Disabled' }
        $archive = if ($DeviceSettings.AA3347 -and ($DeviceSettings.AA3347 -join '') -eq 'true') { 'Enabled' } else { 'Disabled' }
        
        # Extract protected resource counts (mailboxes)
        $protectedUserMbx = if ($DeviceSettings.GM) { $DeviceSettings.GM -join '' } else { '0' }
        $protectedSharedMbx = if ($DeviceSettings.'G@') { $DeviceSettings.'G@' -join '' } else { '0' }
        $totalProtected = [int]$protectedUserMbx + [int]$protectedSharedMbx
        $resourceLine = "$totalProtected mailboxes protected ($protectedUserMbx users, $protectedSharedMbx shared)"
        
        $datasourceLines += @"

M365 Exchange Status: $status
  $(Format-AlignedLabel 'Auto Add Users')$autoAdd
  $(Format-AlignedLabel 'In-Place Archive')$archive
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Resources')$resourceLine
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # M365 OneDrive (D20)
    if ($DataSourceString -match 'D20') {
        $statusCode = $DeviceSettings.D20F0 -join ''
        
        # Check if datasource is configured (has session start time)
        if ([string]::IsNullOrWhiteSpace($statusCode) -or $statusCode -eq '0') {
            $status = "Not Configured"
        } else {
            $status = Get-SessionStatusText $statusCode
            
            # If current status is InProcess, show last completed status -> InProcess
            if ($statusCode -eq "1") {  # InProcess
                $lastCompletedStatusCode = $DeviceSettings.JJ -join ''
                $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
                $status = "$lastCompletedStatus -> $status"
            }
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.JL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.JL
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.JO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.JO
        
        # Validate Last Success accuracy using JQ (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.JQ -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.JL -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.JA -and ($DeviceSettings.JA -join '')) { 
            $seconds = [int]($DeviceSettings.JA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.JO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.J3
        $processed = Format-DataSize $DeviceSettings.J4
        $sent = Format-DataSize $DeviceSettings.J5
        $errors = if ($DeviceSettings.J7) { $DeviceSettings.J7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('OneDrive'))) {
            $sessionId = $ErrorMessages['OneDrive'].SessionId
            $timestampFormatted = $ErrorMessages['OneDrive'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['OneDrive'].Status
            $cleanErrorMsg = $ErrorMessages['OneDrive'].Message  # Already cleaned by Format-CleanErrorMessage
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | OneDrive | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        # Check AutoAdd configuration
        $autoAdd = if ($DeviceSettings.AA3148 -and ($DeviceSettings.AA3148 -join '') -eq 'true') { 'Enabled' } else { 'Disabled' }
        
        # Extract protected resource counts (user accounts)
        $protectedUsers = if ($DeviceSettings.JM) { $DeviceSettings.JM -join '' } else { '0' }
        $resourceLine = "$protectedUsers user accounts protected"
        
        $datasourceLines += @"

M365 OneDrive Status: $status
  $(Format-AlignedLabel 'Auto Add Users')$autoAdd
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Resources')$resourceLine
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # M365 SharePoint (D05/D5)
    if ($DataSourceString -match 'D05|D5') {
        $statusCode = $DeviceSettings.D5F0 -join ''
        
        # Check if datasource is configured (has session start time)
        if ([string]::IsNullOrWhiteSpace($statusCode) -or $statusCode -eq '0') {
            $status = "Not Configured"
        } else {
            $status = Get-SessionStatusText $statusCode
            
            # If current status is InProcess, show last completed status -> InProcess
            if ($statusCode -eq "1") {  # InProcess
                $lastCompletedStatusCode = $DeviceSettings.D5F17 -join ''
                $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
                $status = "$lastCompletedStatus -> $status"
            }
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.D5F9 -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.D5F9
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.D5F18 -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.D5F18
        
        # Validate Last Success accuracy using D5F16 (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.D5F16 -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.D5F9 -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.D5F12 -and ($DeviceSettings.D5F12 -join '')) { 
            $seconds = [int]($DeviceSettings.D5F12 -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.D5F18 -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.D5F3
        $processed = Format-DataSize $DeviceSettings.D5F4
        $sent = Format-DataSize $DeviceSettings.D5F5
        $errors = if ($DeviceSettings.D5F6) { $DeviceSettings.D5F6 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('SharePoint'))) {
            $sessionId = $ErrorMessages['SharePoint'].SessionId
            $timestampFormatted = $ErrorMessages['SharePoint'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['SharePoint'].Status
            $cleanErrorMsg = $ErrorMessages['SharePoint'].Message  # Already cleaned by Format-CleanErrorMessage
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | SharePoint | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        # Check AutoAdd configuration
        $autoAdd = if ($DeviceSettings.AA3149 -and ($DeviceSettings.AA3149 -join '') -eq 'true') { 'Enabled' } else { 'Disabled' }
        
        # Extract protected resource counts (sites)
        $protectedSites = if ($DeviceSettings.D5F22) { $DeviceSettings.D5F22 -join '' } else { '0' }
        $resourceLine = "$protectedSites sites protected"
        
        $datasourceLines += @"

M365 SharePoint Status: $status
  $(Format-AlignedLabel 'Auto Add Sites')$autoAdd
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Resources')$resourceLine
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # M365 Teams (D23)
    if ($DataSourceString -match 'D23') {
        $statusCode = $DeviceSettings.D23F0 -join ''
        
        # Check if datasource is configured (has session start time)
        if ([string]::IsNullOrWhiteSpace($statusCode) -or $statusCode -eq '0') {
            $status = "Not Configured"
        } else {
            $status = Get-SessionStatusText $statusCode
            
            # If current status is InProcess, show last completed status -> InProcess
            if ($statusCode -eq "1") {  # InProcess
                $lastCompletedStatusCode = $DeviceSettings.D23F17 -join ''
                $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
                $status = "$lastCompletedStatus -> $status"
            }
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.D23F9 -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.D23F9
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.D23F18 -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.D23F18
        
        # Validate Last Success accuracy using D23F16 (Last Success Status Code)
        $lastSuccessStatus = $DeviceSettings.D23F16 -join ''
        $lastSuccessValid = $true
        $lastSuccessWarning = ""
        if ($lastSuccessStatus -and $lastSuccessStatus -ne '5' -and ($DeviceSettings.D23F9 -join '') -gt 0) {
            $lastSuccessValid = $false
            $lastSuccessWarning = " ⚠ NOT TRULY SUCCESSFUL (Status=$(Get-SessionStatusText $lastSuccessStatus))"
        }
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.D23F12 -and ($DeviceSettings.D23F12 -join '')) { 
            $seconds = [int]($DeviceSettings.D23F12 -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.D23F18 -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = Format-DataSize $DeviceSettings.D23F3
        $processed = Format-DataSize $DeviceSettings.D23F4
        $sent = Format-DataSize $DeviceSettings.D23F5
        $errors = if ($DeviceSettings.D23F6) { $DeviceSettings.D23F6 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('Teams'))) {
            $sessionId = $ErrorMessages['Teams'].SessionId
            $timestampFormatted = $ErrorMessages['Teams'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['Teams'].Status
            $cleanErrorMsg = $ErrorMessages['Teams'].Message  # Already cleaned by Format-CleanErrorMessage
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | Teams | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        # Check AutoAdd configuration
        $autoAdd = if ($DeviceSettings.AA3150 -and ($DeviceSettings.AA3150 -join '') -eq 'true') { 'Enabled' } else { 'Disabled' }
        
        # Extract protected resource counts (teams and channels)
        $protectedTeams = if ($DeviceSettings.D23F23) { $DeviceSettings.D23F23 -join '' } else { '0' }
        $protectedChannels = if ($DeviceSettings.D23F24) { $DeviceSettings.D23F24 -join '' } else { '0' }
        $resourceLine = "$protectedTeams teams protected ($protectedChannels channels)"
        
        $datasourceLines += @"

M365 Teams Status: $status
  $(Format-AlignedLabel 'Auto Add Teams')$autoAdd
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo$lastSuccessWarning
  $(Format-AlignedLabel 'Resources')$resourceLine
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # Exchange On-Premises (D04/D4) - uses legacy code X
    if ($DataSourceString -match 'D04|D4') {
        $statusCode = $DeviceSettings.X0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.XJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.XL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.XL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.XO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.XO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.XA -and ($DeviceSettings.XA -join '')) { 
            $seconds = [int]($DeviceSettings.XA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.XO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.X3 -and ($DeviceSettings.X3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.X3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.X4 -and ($DeviceSettings.X4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.X4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.X5 -and ($DeviceSettings.X5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.X5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.X7) { $DeviceSettings.X7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('ExchangeVSS'))) {
            $sessionId = $ErrorMessages['ExchangeVSS'].SessionId
            $timestampFormatted = $ErrorMessages['ExchangeVSS'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['ExchangeVSS'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['ExchangeVSS'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | ExchangeVSS | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

Exchange (On-Premises) Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # Network Shares (D06/D6) - uses legacy code N
    if ($DataSourceString -match 'D06|D6') {
        $statusCode = $DeviceSettings.N0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.NJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.NL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.NL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.NO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.NO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.NA -and ($DeviceSettings.NA -join '')) { 
            $seconds = [int]($DeviceSettings.NA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.NO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.N3 -and ($DeviceSettings.N3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.N3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.N4 -and ($DeviceSettings.N4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.N4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.N5 -and ($DeviceSettings.N5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.N5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.N7) { $DeviceSettings.N7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('NetworkShares'))) {
            $sessionId = $ErrorMessages['NetworkShares'].SessionId
            $timestampFormatted = $ErrorMessages['NetworkShares'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['NetworkShares'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['NetworkShares'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | NetworkShares | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

Network Shares Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # VMware (D08/D8) - uses legacy code W
    if ($DataSourceString -match 'D08|D8') {
        $statusCode = $DeviceSettings.W0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.WJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.WL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.WL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.WO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.WO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.WA -and ($DeviceSettings.WA -join '')) { 
            $seconds = [int]($DeviceSettings.WA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.WO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.W3 -and ($DeviceSettings.W3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.W3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.W4 -and ($DeviceSettings.W4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.W4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.W5 -and ($DeviceSettings.W5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.W5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.W7) { $DeviceSettings.W7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('VMware'))) {
            $sessionId = $ErrorMessages['VMware'].SessionId
            $timestampFormatted = $ErrorMessages['VMware'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['VMware'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['VMware'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | VMware | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

VMware Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # Oracle (D12) - DEPRECATED - Commented out 2025-12-26
    <#
    if ($DataSourceString -match 'D12') {
        $statusCode = $DeviceSettings.Y0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.YJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.YL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.YL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.YO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.YO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.YA -and ($DeviceSettings.YA -join '')) { 
            $seconds = [int]($DeviceSettings.YA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.YO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.Y3 -and ($DeviceSettings.Y3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Y3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.Y4 -and ($DeviceSettings.Y4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Y4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.Y5 -and ($DeviceSettings.Y5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.Y5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.Y7) { $DeviceSettings.Y7 -join '' } else { '0' }
        
        $datasourceLines += @"

Oracle Status
  $(Format-AlignedLabel 'Status')$status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors
"@
    }
    #>
    
    # Hyper-V (D14) - uses legacy code H
    if ($DataSourceString -match 'D14') {
        $statusCode = $DeviceSettings.H0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.HJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.HL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.HL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.HO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.HO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.HA -and ($DeviceSettings.HA -join '')) { 
            $seconds = [int]($DeviceSettings.HA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.HO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.H3 -and ($DeviceSettings.H3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.H3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.H4 -and ($DeviceSettings.H4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.H4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.H5 -and ($DeviceSettings.H5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.H5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.H7) { $DeviceSettings.H7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('HyperV'))) {
            $sessionId = $ErrorMessages['HyperV'].SessionId
            $timestampFormatted = $ErrorMessages['HyperV'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['HyperV'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['HyperV'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | HyperV | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

Hyper-V Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    # MySQL (D15) - uses legacy code L
    if ($DataSourceString -match 'D15') {
        $statusCode = $DeviceSettings.L0 -join ''
        $status = Get-SessionStatusText $statusCode
        
        # If current status is InProcess, show last completed status -> InProcess
        if ($statusCode -eq "1") {  # InProcess
            $lastCompletedStatusCode = $DeviceSettings.LJ -join ''
            $lastCompletedStatus = Get-SessionStatusText $lastCompletedStatusCode
            $status = "$lastCompletedStatus -> $status"
        }
        
        $tz = $DeviceSettings.TZ -join ''
        $lastSession = Get-SafeTimestamp $DeviceSettings.TL -TimezoneOffset $tz -WithLabel
        $lastSessionAgo = Get-TimeAgo $DeviceSettings.TL -TimezoneOffset $tz
        $lastSuccess = Get-SafeTimestamp $DeviceSettings.LL -TimezoneOffset $tz -WithLabel
        $lastSuccessAgo = Get-TimeAgo $DeviceSettings.LL -TimezoneOffset $tz
        $lastCompleted = Get-SafeTimestamp $DeviceSettings.LO -TimezoneOffset $tz -WithLabel
        $lastCompletedAgo = Get-TimeAgo $DeviceSettings.LO -TimezoneOffset $tz
        
        # Calculate session start time (completion time - duration)
        $sessionStart = 'N/A'
        $duration = if ($DeviceSettings.LA -and ($DeviceSettings.LA -join '')) { 
            $seconds = [int]($DeviceSettings.LA -join '')
            $ts = New-TimeSpan -Seconds $seconds
            
            # Calculate start time
            $completedUnix = $DeviceSettings.LO -join ''
            if ($completedUnix -and $completedUnix -gt 0) {
                $startUnix = [int64]$completedUnix - $seconds
                $sessionStart = Get-SafeTimestamp $startUnix -TimezoneOffset $tz -WithLabel
            }
            
            $ts.ToString('hh\:mm\:ss')
        } else { 'N/A' }
        $selected = if ($DeviceSettings.L3 -and ($DeviceSettings.L3 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.L3 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $processed = if ($DeviceSettings.L4 -and ($DeviceSettings.L4 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.L4 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $sent = if ($DeviceSettings.L5 -and ($DeviceSettings.L5 -join '')) { 
            [Math]::Round([Decimal](($DeviceSettings.L5 -join '') / 1GB), 2).ToString('0.00') + ' GB'
        } else { 'N/A' }
        $errors = if ($DeviceSettings.L7) { $DeviceSettings.L7 -join '' } else { '0' }
        
        # Get error message if available AND error count > 0
        $errorDetail = ""
        if (([int]$errors -gt 0) -and ($ErrorMessages.ContainsKey('MySQL'))) {
            $sessionId = $ErrorMessages['MySQL'].SessionId
            $timestampFormatted = $ErrorMessages['MySQL'].Timestamp  # Already formatted with relative time
            $errorStatus = Get-SessionStatusText $ErrorMessages['MySQL'].Status
            $cleanErrorMsg = Format-CleanErrorMessage $ErrorMessages['MySQL'].Message
            $errorDetail = "`n  $(Format-AlignedLabel 'Error Session')(ID: $sessionId) | MySQL | $errorStatus | $timestampFormatted`n  $(Format-AlignedLabel 'Error Message')$cleanErrorMsg"
        }
        
        $datasourceLines += @"

MySQL Status: $status
  $(Format-AlignedLabel 'Session Start')$sessionStart
  $(Format-AlignedLabel 'Last Completed')$lastCompleted $lastCompletedAgo
  $(Format-AlignedLabel 'Last Success')$lastSuccess $lastSuccessAgo
  $(Format-AlignedLabel 'Duration')$duration | Selected: $selected | Processed: $processed | Sent: $sent | Errors: $errors$errorDetail
─────────────────────────────────────────────────────────────────
"@
    }
    
    if ($datasourceLines.Count -gt 0) {
        return $datasourceLines -join "`n"
    } else {
        return "  No datasource details available"
    }
}

Function Convert-DataSourceCodes {
    param(
        [string]$DataSourceString
    )
    
    # Data source code mapping from Cove API column I78 (new D## format)
    # Reference: https://developer.n-able.com/n-able-cove/docs/column-codes
    # Note: API sometimes returns D1 instead of D01, D2 instead of D02, etc.
    $dataSourceMap = @{
        'D01' = 'Files and Folders'
        'D1'  = 'Files and Folders'
        'D02' = 'System State'
        'D2'  = 'System State'
        'D04' = 'MS Exchange'
        'D4'  = 'MS Exchange'
        'D05' = 'MS SharePoint Online'
        'D5'  = 'MS SharePoint Online'
        'D06' = 'Network Shares'
        'D6'  = 'Network Shares'
        'D08' = 'VMware'
        'D8'  = 'VMware'
        'D10' = 'MS SQL'
        # 'D12' = 'Oracle'  # DEPRECATED 2025-12-26
        'D14' = 'Hyper-V'
        'D15' = 'MySQL'
        'D19' = 'MS Exchange Online'
        'D20' = 'MS OneDrive for Business'
        'D23' = 'MS Teams'
    }
    
    # Split by D and process each code (handles both D01 and D1 formats)
    $dataSources = @()
    $codes = $DataSourceString -split '(?=D\d+)' | Where-Object { $_ -match 'D\d+' }
    
    foreach ($code in $codes) {
        if ($dataSourceMap.ContainsKey($code)) {
            $dataSources += $dataSourceMap[$code]
        } else {
            $dataSources += $code  # Keep unknown codes as-is
        }
    }
    
    if ($dataSources.Count -eq 0 -and $DataSourceString) {
        return $DataSourceString  # Return original if no matches found
    }
    
    return ($dataSources -join ', ')
}  ## Convert I78 column data source codes (D## or D#) to full names

Function Get-DatasourceCodes {
    <#
    .SYNOPSIS
    Convert datasource codes to abbreviations - for failed OR stale datasources
    .DESCRIPTION
    Checks status codes for each configured datasource and only includes those that have failed (status != 5)
    OR if StaleHours is provided, checks last success timestamps for stale datasources
    #>
    param(
        [string]$DataSourceString,
        [object]$DeviceSettings,  # Optional: Check status codes to only include failed datasources
        [int]$StaleHours = 0       # Optional: If provided, check for stale datasources instead of failed
    )
    
    if (-not $DataSourceString) { return "" }
    
    $codes = @()
    
    # If DeviceSettings provided, check either status codes (failed) or timestamps (stale)
    if ($DeviceSettings) {
        if ($StaleHours -gt 0) {
            # STALE MODE: Check last success timestamps against threshold
            # File System (D01) - Check FL (Last Success)
            if ($DataSourceString -match 'D01') {
                $lastSuccess = $DeviceSettings.FL -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'FS' }
                        }
                    } catch { }
                }
            }
            # System State (D02) - Check SL (Last Success)
            if ($DataSourceString -match 'D02') {
                $lastSuccess = $DeviceSettings.SL -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'SS' }
                        }
                    } catch { }
                }
            }
            # M365 Exchange (D19) - Check GL (Last Success)
            if ($DataSourceString -match 'D19') {
                $lastSuccess = $DeviceSettings.GL -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'EX' }
                        }
                    } catch { }
                }
            }
            # M365 OneDrive (D20) - Check JL (Last Success)
            if ($DataSourceString -match 'D20') {
                $lastSuccess = $DeviceSettings.JL -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'OD' }
                        }
                    } catch { }
                }
            }
            # M365 SharePoint (D5) - Check D5F9 (Last Success)
            if ($DataSourceString -match 'D5') {
                $lastSuccess = $DeviceSettings.D5F9 -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'SP' }
                        }
                    } catch { }
                }
            }
            # M365 Teams (D23) - Check D23F9 (Last Success)
            if ($DataSourceString -match 'D23') {
                $lastSuccess = $DeviceSettings.D23F9 -join ''
                if ($lastSuccess -and $lastSuccess -ne '' -and $lastSuccess -ne 'N/A') {
                    try {
                        $lastSuccessDate = Convert-UnixTimeToDateTime ([int]$lastSuccess)
                        if ($lastSuccessDate -and $lastSuccessDate -is [DateTime]) {
                            $hoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $lastSuccessDate).TotalHours, 1)
                            if ($hoursSinceSuccess -ge $StaleHours) { $codes += 'TM' }
                        }
                    } catch { }
                }
            }
        } else {
            # FAILED MODE: Check status codes for failures (status != 5)
            # File System (D01) - Check F0 status
            if ($DataSourceString -match 'D01') {
                $status = $DeviceSettings.F0 -join ''
                if ($status -and $status -ne '5') { $codes += 'FS' }
            }
            # System State (D02) - Check S0 status
            if ($DataSourceString -match 'D02') {
                $status = $DeviceSettings.S0 -join ''
                if ($status -and $status -ne '5') { $codes += 'SS' }
            }
            # M365 Exchange (D19) - Check D19F0 status
            if ($DataSourceString -match 'D19') {
                $status = $DeviceSettings.D19F0 -join ''
                if ($status -and $status -ne '5') { $codes += 'EX' }
            }
            # M365 OneDrive (D20) - Check D20F0 status
            if ($DataSourceString -match 'D20') {
                $status = $DeviceSettings.D20F0 -join ''
                if ($status -and $status -ne '5') { $codes += 'OD' }
            }
            # M365 SharePoint (D5) - Check D5F0 status
            if ($DataSourceString -match 'D5') {
                $status = $DeviceSettings.D5F0 -join ''
                if ($status -and $status -ne '5') { $codes += 'SP' }
            }
            # M365 Teams (D23) - Check D23F0 status
            if ($DataSourceString -match 'D23') {
                $status = $DeviceSettings.D23F0 -join ''
                if ($status -and $status -ne '5') { $codes += 'TM' }
            }
        }
    } else {
        # Fallback: Include all configured datasources if no settings provided
        if ($DataSourceString -match 'D01') { $codes += 'FS' }  # File System
        if ($DataSourceString -match 'D02') { $codes += 'SS' }  # System State
        if ($DataSourceString -match 'D19') { $codes += 'EX' }  # M365 Exchange
        if ($DataSourceString -match 'D20') { $codes += 'OD' }  # M365 OneDrive
        if ($DataSourceString -match 'D5') { $codes += 'SP' }   # M365 SharePoint
        if ($DataSourceString -match 'D23') { $codes += 'TM' }  # M365 Teams
    }
    
    if ($codes.Count -eq 0) { return "" }
    if ($codes.Count -eq 1) { return $codes[0] }
    if ($codes.Count -eq 2) { return ($codes -join '/') }
    return "MULTI($($codes.Count))"
}  ## Convert datasource codes to abbreviations - ONLY failed datasources

#endregion ----- Data Conversion Functions ----

#region ----- Error Retrieval Functions ----

Function Get-AccountInfoById {
    <#
    .SYNOPSIS
        Retrieves device-specific repserv connection details for error querying
    .DESCRIPTION
        Calls GetAccountInfoById API to get device token and repserv URL needed for QueryErrors
        OPTIMIZED v09: Added retry logic with exponential backoff for timeout errors
    #>
    Param (
        [Parameter(Mandatory=$True)] [string]$accountid
    )
    
    $maxRetries = 3
    $retryCount = 0
    $baseDelay = 2  # Start with 2 seconds
    
    while ($retryCount -le $maxRetries) {
        try {
            $url = "https://api.backup.management/jsonapi"

            # GetAccountInfoById request
            $dataAccountInfo = @{
                method  = "GetAccountInfoById"
                params  = @{
                    accountId = [int]$accountid
                }
                jsonrpc = "2.0"
                visa    = $Script:visa
                id      = "jsonrpc"
            }

            $responseAccountInfo = Invoke-RestMethod -Uri $url -Method POST -ContentType 'application/json' -Body ($dataAccountInfo | ConvertTo-Json -Depth 5) -ErrorAction Stop

            # Build repserv object
            $repsvr = @{}
            $repsvr.repurl = ("https://" + ($responseAccountInfo.result.homeNodeInfo.CommonInfo.Host -split ':')[0] + "/repserv_json").ToLower()
            $repsvr.AccountId = $responseAccountInfo.result.result.Id
            $repsvr.Token = $responseAccountInfo.result.result.Token
            $repsvr.Name = $responseAccountInfo.result.result.Name

            return $repsvr
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            # Check if this is a retryable error (timeout, connection failure)
            $isRetryable = $errorMsg -match "timeout|timed out|connection.*failed|failed to respond|Unable to connect"
            
            if ($isRetryable -and $retryCount -lt $maxRetries) {
                # Exponential backoff: 2s, 4s, 8s
                $delay = $baseDelay * [Math]::Pow(2, $retryCount)
                $retryCount++
                Write-Verbose "  Timeout on AccountId $accountid (attempt $retryCount/$maxRetries) - retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
                continue
            }
            
            # Check for SSL/TLS error pattern
            if ($errorMsg -like "*SSL connection could not be established*" -or $errorMsg -like "*secure channel*") {
                # CRITICAL: SSL error on GetAccountInfoById means we failed to get device info from MAIN API
                # This is NOT a software-only infrastructure issue - it's a network/timeout problem with PRIMARY Cove API
                # RepURL will be empty because the API call failed, so we CAN'T use it for diagnostics
                
                Write-Warning "SSL/Network Error calling GetAccountInfoById for AccountId $accountid - Main Cove API timeout/SSL issue (RETRYABLE)"
                
                # CAPTURE FULL DEBUG OUTPUT FOR SSL CERTIFICATE ERRORS
                $debugOutput = @()
                $debugOutput += "="*100
                $debugOutput += "SSL/NETWORK ERROR - MAIN COVE API"
                $debugOutput += "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                $debugOutput += "="*100
                $debugOutput += ""
                $debugOutput += "DEVICE IDENTIFICATION:"
                $debugOutput += "  Account ID: $accountid"
                $debugOutput += ""
                $debugOutput += "API CALL DETAILS:"
                $debugOutput += "  Method: GetAccountInfoById"
                $debugOutput += "  Endpoint: https://api.backup.management/jsonapi (PRIMARY COVE API)"
                $debugOutput += "  Retry Attempt: $retryCount of $maxRetries"
                $debugOutput += ""
                $debugOutput += "ERROR DETAILS:"
                $debugOutput += "  Error Message: $errorMsg"
                $debugOutput += "  Error Type: $($_.Exception.GetType().FullName)"
                $debugOutput += "  Error Category: $($_.CategoryInfo.Category)"
                if ($_.Exception.InnerException) {
                    $debugOutput += "  Inner Exception: $($_.Exception.InnerException.Message)"
                }
                $debugOutput += ""
                $debugOutput += "FULL EXCEPTION STACK TRACE:"
                $debugOutput += $_.Exception.ToString()
                $debugOutput += ""
                $debugOutput += "DIAGNOSTIC NOTES:"
                $debugOutput += "  ⚠ SSL/TLS error calling PRIMARY Cove API (backup.management)"
                $debugOutput += "  ⚠ This is NOT a software-only infrastructure issue"
                $debugOutput += "  ⚠ This is a NETWORK/TIMEOUT problem with the main API"
                $debugOutput += "  Error is RETRYABLE - should be added to deferred retry queue"
                $debugOutput += "  Common causes: Network congestion, API throttling, transient SSL handshake failures"
                $debugOutput += "  RepURL is EMPTY because we never got a response from the API"
                $debugOutput += "  Cannot determine infrastructure type - error occurred before that data was retrieved"
                $debugOutput += ""
                $debugOutput += "="*100
                $debugOutput += ""
                
                # Append to debug log file
                $debugOutput | Out-File -FilePath $sslDebugLogPath -Append -Encoding UTF8
                if ($DebugCDP) { Write-Host "  [DEBUG] SSL error details saved to: $sslDebugLogPath" -ForegroundColor Gray }
                
                # ADD TO DEFERRED RETRY QUEUE - This is a retryable network/API error
                # Note: Caller should check return value of $null and add to deferred retry queue
            } else {
                # Real API error (after retries exhausted)
                if ($retryCount -gt 0) {
                    Write-Warning "Get-AccountInfoById API Error for AccountId $accountid (after $retryCount retries): $errorMsg"
                } else {
                    Write-Warning "Get-AccountInfoById API Error for AccountId $accountid : $errorMsg"
                }
            }
            
            return $null
        }
    }
}

Function Get-M365SessionErrors {
    <#
    .SYNOPSIS
        Retrieves error messages from M365 backup sessions using reporting_api endpoint
    .DESCRIPTION
        M365 devices use the reporting_api EnumerateSessionErrors method instead of repserv QueryErrors.
        This function:
        1. Calls EnumerateSessionErrors with accountToken authentication
        2. Filters by SessionId and SessionType (Backup/Restore)
        3. Parses the Description field which contains JSON error details
        4. Extracts the "description" property from the JSON
        5. Returns array of error messages ready for cleaning
        
        ARCHITECTURE NOTE:
        - Systems devices (AccountType=1): Use repserv endpoint with QueryErrors
        - M365 devices (AccountType=2): Use reporting_api endpoint with EnumerateSessionErrors
        
    .PARAMETER AccountId
        The M365 device AccountId
    .PARAMETER AccountToken
        The account token from GetAccountInfoById (result.result.Token)
    .PARAMETER SessionId
        The specific session ID to query errors for
    .PARAMETER SessionType
        "Backup" or "Restore" - filters error type
    .PARAMETER MaxErrors
        Maximum number of errors to retrieve (default: 100)
        
    .OUTPUTS
        Array of error message strings, or $null if no errors or API failure
        
    .EXAMPLE
        $errors = Get-M365SessionErrors -AccountId 12345 -AccountToken "abc-def-ghi" -SessionId 67890 -SessionType "Backup"
    #>
    
    Param(
        [Parameter(Mandatory=$True)] [int]$AccountId,
        [Parameter(Mandatory=$True)] [string]$AccountToken,
        [Parameter(Mandatory=$True)] [int]$SessionId,
        [Parameter(Mandatory=$True)] [string]$DataSourceType,
        [Parameter(Mandatory=$False)] [ValidateSet("Backup","Restore")] [string]$SessionType = "Backup",
        [Parameter(Mandatory=$False)] [int]$MaxErrors = 100
    )
    
    try {
        if ($Script:DebugCDP) {
            Write-Host "`n  [DEBUG] Get-M365SessionErrors called:" -ForegroundColor Magenta
            Write-Host "    AccountId      : $AccountId" -ForegroundColor Gray
            Write-Host "    SessionId      : $SessionId" -ForegroundColor Gray
            Write-Host "    DataSourceType : $DataSourceType" -ForegroundColor Gray
            Write-Host "    SessionType    : $SessionType" -ForegroundColor Gray
            Write-Host "    MaxErrors      : $MaxErrors" -ForegroundColor Gray
        }
        
        # Build EnumerateSessionErrors request for M365 reporting_api
        # CRITICAL: Use /reporting_api endpoint (NOT /jsonapi)
        # CRITICAL: Do NOT include visa in body - Invoke-RestMethod adds Authorization header automatically
        $url = "https://api.backup.management/reporting_api"
        
        $dataSessionErrors = @{
            jsonrpc = "2.0"
            id      = "jsonrpc"
            method  = "EnumerateSessionErrors"
            params  = @{
                accountToken = $AccountToken
                dataSourceType = $DataSourceType
                filter = @{
                    SessionId   = $SessionId
                    SessionType = $SessionType
                }
                range = @{
                    Offset = 0
                    Size   = $MaxErrors
                }
            }
        }
        
        $jsonBody = $dataSessionErrors | ConvertTo-Json -Depth 5
        
        if ($DebugCDP) {
            Write-Host "    API Endpoint : $url" -ForegroundColor Gray
            Write-Host "    Method       : EnumerateSessionErrors (reporting_api)" -ForegroundColor Gray
        }
        
        # Call reporting_api EnumerateSessionErrors with Bearer token authentication
        $headers = @{
            Authorization = "Bearer $Script:visa"
        }
        $response = Invoke-RestMethod -Uri $url -Method POST -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonBody)) -ErrorAction Stop
        
        # Check for API-level errors
        if ($response.error) {
            Write-Warning "  M365 EnumerateSessionErrors API Error (AccountId $AccountId, SessionId $SessionId): $($response.error.message)"
            if ($DebugCDP) {
                Write-Host "    Error Code   : $($response.error.code)" -ForegroundColor Red
                Write-Host "    Error Message: $($response.error.message)" -ForegroundColor Red
            }
            return $null
        }
        
        # Extract error array from response
        $sessionErrors = $response.result.result
        
        if (-not $sessionErrors -or $sessionErrors.Count -eq 0) {
            if ($DebugCDP) {
                Write-Host "    No errors found for SessionId $SessionId" -ForegroundColor Gray
            }
            return $null
        }
        
        if ($DebugCDP) {
            Write-Host "    Errors Found : $($sessionErrors.Count)" -ForegroundColor Cyan
        }
        
        # Parse error messages from Description JSON field
        # Return array of error objects with Text and Time properties (matching Systems format)
        $errorObjects = @()
        
        foreach ($error in $sessionErrors) {
            try {
                # M365 errors come in Description field as JSON string
                # Format: {"description":"Graph API error message","entity":"user@domain.com","path":"/Inbox/Folder"}
                
                $errorText = $null
                
                if ($error.Description) {
                    # Parse the JSON description
                    $descriptionJson = $error.Description | ConvertFrom-Json -ErrorAction Stop
                    
                    # Extract the "description" property (actual error message)
                    if ($descriptionJson.description) {
                        $errorText = $descriptionJson.description
                        
                        if ($DebugCDP) {
                            Write-Host "      Error: $($descriptionJson.description.Substring(0, [Math]::Min(80, $descriptionJson.description.Length)))..." -ForegroundColor Gray
                            if ($descriptionJson.entity) {
                                Write-Host "      Entity: $($descriptionJson.entity)" -ForegroundColor DarkGray
                            }
                            if ($descriptionJson.path) {
                                Write-Host "      Path: $($descriptionJson.path)" -ForegroundColor DarkGray
                            }
                            if ($error.Timestamp) {
                                $ts = Convert-UnixTimeToDateTime $error.Timestamp
                                Write-Host "      Timestamp: $($ts.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor DarkGray
                            }
                        }
                    }
                    else {
                        # Description JSON doesn't have "description" property - use raw Description
                        Write-Warning "  M365 error missing 'description' property in JSON - using raw Description"
                        $errorText = $error.Description
                    }
                }
                else {
                    # No Description field at all
                    Write-Warning "  M365 error missing Description field (SessionId $SessionId, ErrorId $($error.Id))"
                }
                
                # Build error object matching Systems format (has Text and Time properties)
                if ($errorText) {
                    $errorObjects += [PSCustomObject]@{
                        Text = $errorText
                        Time = $error.Timestamp  # Unix timestamp from M365 API
                        Id = if ($error.Id) { $error.Id } else { 0 }
                    }
                }
            }
            catch {
                # JSON parsing failed - use raw Description as fallback
                Write-Warning "  Failed to parse M365 error Description JSON (SessionId $SessionId): $($_.Exception.Message)"
                if ($error.Description) {
                    $errorObjects += [PSCustomObject]@{
                        Text = $error.Description
                        Time = $error.Timestamp
                        Id = if ($error.Id) { $error.Id } else { 0 }
                    }
                }
            }
        }
        
        if ($DebugCDP) {
            Write-Host "    Parsed Errors: $($errorObjects.Count) error objects with timestamps" -ForegroundColor Green
        }
        
        return $errorObjects
    }
    catch {
        Write-Warning "  Get-M365SessionErrors Exception (AccountId $AccountId, SessionId $SessionId): $($_.Exception.Message)"
        if ($DebugCDP) {
            Write-Host "    Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            Write-Host "    Stack Trace:`n$($_.ScriptStackTrace)" -ForegroundColor DarkGray
        }
        return $null
    }
}

Function Get-LatestSessionError {
    <#
    .SYNOPSIS
        Retrieves error messages from recent sessions grouped by datasource
    .DESCRIPTION
        OPTIMIZED VERSION: Uses server-side filtering to query only sessions with errors
        1. For each datasource, queries sessions with errors (ErrorsCount >= 1)
        2. Gets most recent error session per datasource
        3. Queries detailed error messages for those sessions
        4. Compares against last successful session to determine if errors are stale
    .PARAMETER DeviceId
        The AccountId of the device to query
    .PARAMETER Days
        Filter errors to those within the last X days (default=7)
    .PARAMETER AccountType
        Device account type: 1=Server/Workstation, 2=M365. For M365, only queries Exchange, OneDrive, SharePoint, Teams (default=1)
    .PARAMETER DataSources
        I78 datasource string (e.g., "D01D02") - only query datasources that are actually configured (OPTIMIZATION v08)
    .OUTPUTS
        Returns hashtable with datasource error messages: @{ "Files" = "Error message (SessionId: X, Time: YYYY-MM-DD HH:MM:SS)"; "SystemState" = "..." }
    #>
    Param (
        [Parameter(Mandatory=$True)] [string]$DeviceId,
        [Parameter(Mandatory=$False)] [int]$Days = 7,
        [Parameter(Mandatory=$False)] [int]$AccountType = 1,  # 1=Server/Workstation, 2=M365
        [Parameter(Mandatory=$False)] [string]$DataSources = "",  # OPTIMIZATION v08: I78 value to filter datasources
        [Parameter(Mandatory=$False)] [int]$DeviceNum = 0,  # Device counter for progress display
        [Parameter(Mandatory=$False)] [int]$TotalDevices = 0  # Total device count for progress display
    )

    try {
        # Get device account info and repserv details
        $repsvr = Get-AccountInfoById -accountid $DeviceId
        
        if (-not $repsvr) {
            Write-Verbose "DeviceId# $DeviceId - Unable to retrieve account info for error querying"
            return $null
        }

        Write-Verbose "DeviceId# $DeviceId - Querying sessions from $($repsvr.repurl)"

        # Map Plugin numeric IDs to datasource names (based on Cove API QuerySessions response)
        # NOTE: Correct Plugin IDs verified from API QuerySessions calls
        $pluginIdToName = @{
            1  = "FileSystem"         # FileSystem
            7  = "SystemState"        # VssSystemState (Plugin 7)
            18 = "SystemState"        # VssSystemState (Plugin 18 - alternate ID)
            4  = "Exchange"           # VssExchange (on-premises)
            5  = "SharePoint"         # SharePointOnline
            6  = "NetworkShares"      # NetworkShares
            8  = "VMware"             # VmwareVirtualMachines
            10 = "MSSQL"              # MsSql
            # 12 = "Oracle"             # Oracle - DEPRECATED 2025-12-26
            14 = "HyperV"             # VssHyperV
            15 = "MySQL"              # MySql
            19 = "Exchange"           # ExchangeOnline (M365) - maps to same name as on-prem
            20 = "OneDrive"           # OneDriveForBusiness
            23 = "Teams"              # TeamsOnline
        }
        
        # Build inverse map: datasource name to Plugin IDs (handles duplicates like Exchange)
        $nameToPluginIds = @{}
        foreach ($pluginId in $pluginIdToName.Keys) {
            $name = $pluginIdToName[$pluginId]
            if (-not $nameToPluginIds.ContainsKey($name)) {
                $nameToPluginIds[$name] = @()
            }
            $nameToPluginIds[$name] += $pluginId
        }

        # OPTIMIZATION v08: Filter to only configured datasources from I78 column
        # Map datasource codes to names: D01=FileSystem, D02=SystemState, D03=MSSQL, D04=Exchange, etc.
        if ($DataSources -and $DataSources -ne "") {
            Write-Verbose "DeviceId# $DeviceId - Filtering to configured datasources only (I78: $DataSources)"
            $dsCodeToName = @{
                "D01" = "FileSystem"
                "D02" = "SystemState"
                "D03" = "MSSQL"
                "D04" = "Exchange"
                "D05" = "SharePoint"  # Legacy on-prem (rare)
                "D5"  = "SharePoint"  # M365 SharePoint
                "D06" = "NetworkShares"
                "D08" = "VMware"
                "D14" = "HyperV"
                "D15" = "MySQL"
                "D19" = "Exchange"  # M365 Exchange
                "D20" = "OneDrive"
                "D23" = "Teams"
            }
            
            $configuredDatasources = @{}
            foreach ($code in $dsCodeToName.Keys) {
                if ($DataSources -like "*$code*") {
                    $dsName = $dsCodeToName[$code]
                    if ($nameToPluginIds.ContainsKey($dsName)) {
                        $configuredDatasources[$dsName] = $nameToPluginIds[$dsName]
                        Write-Verbose "  Configured datasource: $dsName (Code: $code, Plugin IDs: $($nameToPluginIds[$dsName] -join ', '))"
                    }
                }
            }
            $nameToPluginIds = $configuredDatasources
        }
        # Filter datasources for M365 devices - only query Exchange, OneDrive, SharePoint, Teams
        elseif ($AccountType -eq 2) {
            Write-Verbose "DeviceId# $DeviceId - AccountType is M365 (2) - filtering to M365-specific datasources only"
            $m365Datasources = @("Exchange", "OneDrive", "SharePoint", "Teams")
            $filteredDatasources = @{}
            foreach ($dsName in $m365Datasources) {
                if ($nameToPluginIds.ContainsKey($dsName)) {
                    $filteredDatasources[$dsName] = $nameToPluginIds[$dsName]
                    Write-Verbose "  Including M365 datasource: $dsName (Plugin IDs: $($nameToPluginIds[$dsName] -join ', '))"
                }
            }
            $nameToPluginIds = $filteredDatasources
        } else {
            Write-Verbose "DeviceId# $DeviceId - AccountType is Server/Workstation ($AccountType) - querying all datasources (no I78 filter provided)"
        }

        $datasourceErrors = @{}
        $datasourceSessionTimes = @{}
        $filterDate = (Get-Date).AddDays(-$Days)
        
        # Start timing for session queries
        $sessionQueryStartTime = Get-Date
        $datasourceCount = $nameToPluginIds.Keys.Count
        $currentDatasource = 0
        
        # Display query progress with device counter
        $counterDisplay = if ($DeviceNum -gt 0 -and $TotalDevices -gt 0) { "[$DeviceNum/$TotalDevices] " } else { "" }
        Write-Host "    $counterDisplay[$($repsvr.Name)] Querying $datasourceCount datasource(s)..." -ForegroundColor Cyan
        Write-Verbose "DeviceId# $DeviceId - Querying sessions with errors for each datasource (optimized server-side filtering)"
        
        # OPTIMIZATION: Query each datasource separately with error filter
        # This eliminates need for PowerShell grouping and reduces payload size
        foreach ($datasourceName in $nameToPluginIds.Keys) {
            $currentDatasource++
            $dsQueryStartTime = Get-Date
            $pluginIds = $nameToPluginIds[$datasourceName]
            
            # Build query for this datasource (handle multiple Plugin IDs for same datasource)
            if ($pluginIds.Count -eq 1) {
                $pluginQuery = "Plugin == $($pluginIds[0])"
            } else {
                # Multiple IDs (e.g., Exchange = 4 or 19)
                $pluginQuery = "(" + (($pluginIds | ForEach-Object { "Plugin == $_" }) -join " or ") + ")"
            }
            
            # Query recent sessions for this datasource - ALL statuses except Completed (EXPANDED v10.1)
            # Excluded: Status 5 (Completed) - only successful sessions with no errors
            # Included: 1=InProcess, 2=Failed, 3=Aborted, 4=Unknown, 6=Interrupted, 7=NotStarted, 
            #           8=CompletedWithErrors, 9=InProgressWithFaults, 10=OverQuota, 11=NoSelection, 12=Restarted
            # Rationale: Check for errors in ALL non-successful sessions (running, failed, aborted, etc.)
            $query = "0 != 1 and $pluginQuery and (Status == 1 or Status == 2 or Status == 3 or Status == 4 or Status == 6 or Status == 7 or Status == 8 or Status == 9 or Status == 10 or Status == 11 or Status == 12)"
            
            if ($DebugCDP) { Write-Host "      [$currentDatasource/$datasourceCount] $datasourceName" -ForegroundColor Gray -NoNewline }
            Write-Verbose "  Datasource: $datasourceName | Query: $query"
            
            # ARCHITECTURE SPLIT: M365 vs Systems session queries
            # - M365 (AccountType=2): Use /reporting_api EnumerateSessions
            # - Systems (AccountType=1): Use /repserv_json QuerySessions
            
            if ($AccountType -eq 2) {
                # M365 SESSION QUERY PATH - Use /reporting_api EnumerateSessions
                $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
                $daysAgo = $now - ($Days * 86400)  # Convert days to seconds
                
                # Map datasource names to reporting_api DataSources values
                $dataSourceMap = @{
                    'Exchange' = 'Exchange'
                    'OneDrive' = 'OneDrive'
                    'SharePoint' = 'SharePoint'
                    'Teams' = 'Teams'
                }
                
                $dataSourceFilter = $dataSourceMap[$datasourceName]
                if (-not $dataSourceFilter) {
                    Write-Verbose "    Datasource $datasourceName not mapped for reporting_api, skipping"
                    if ($DebugCDP) { Write-Host " - No errors in date range (0ms)" -ForegroundColor Cyan }
                    continue
                }
                
                $requestBody = @{
                    method = "EnumerateSessions"
                    params = @{
                        accountToken = $repsvr.Token
                        range = @{
                            Offset = 0
                            Size = 200  # Get up to 200 backup sessions to find errors
                        }
                        filter = @{
                            SessionTypes = @("Backup")
                            DataSources = @($dataSourceFilter)
                            CreatedAfter = $daysAgo
                            CreatedBefore = $now
                            # Note: reporting_api doesn't support Status filter, will filter client-side
                        }
                    }
                    jsonrpc = "2.0"
                    id = "jsonrpc"
                }
                
                try {
                    $responseErrors = Invoke-RestMethod `
                        -Uri "https://api.backup.management/reporting_api" `
                        -Method POST `
                        -ContentType 'application/json; charset=utf-8' `
                        -Headers @{ Authorization = "Bearer $Script:visa" } `
                        -Body ([System.Text.Encoding]::UTF8.GetBytes(($requestBody | ConvertTo-Json -Depth 10))) `
                        -ErrorAction Stop
                    
                    # CRITICAL: reporting_api returns sessions in result.result (nested), not result
                    $allSessions = @($responseErrors.result.result)
                    
                    # Debug: Show what we received before filtering
                    if ($DebugCDP -and $allSessions.Count -gt 0) {
                        $statusCounts = $allSessions | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }
                        Write-Verbose "    API returned $($allSessions.Count) session(s) with statuses: $($statusCounts -join ', ')"
                    }
                    
                    # Filter sessions by Status string (API returns text status, not numeric codes)
                    # Include all statuses except 'Completed': InProcess, Failed, Aborted, Unknown, Interrupted, 
                    # NotStarted, CompletedWithErrors, InProgressWithFaults, OverQuota, NoSelection, Restarted
                    $sessions = @($allSessions | Where-Object { 
                        $_.Status -in @('InProcess', 'Failed', 'Aborted', 'Unknown', 'Interrupted', 'NotStarted', 'CompletedWithErrors', 'InProgressWithFaults', 'OverQuota', 'NoSelection', 'Restarted')
                    })
                    
                    if ($DebugCDP) {
                        Write-Verbose "    Filtered to $($sessions.Count) session(s) with errors (Status='CompletedWithErrors'/'Failed' or ErrorsCount>0)"
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    if ($DebugCDP) {
                        Write-Host " - Error (" -ForegroundColor Yellow -NoNewline
                        Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                        Write-Host ")" -ForegroundColor Yellow
                    }
                    Write-Verbose "    M365 reporting_api EnumerateSessions returned error: $errorMsg"
                    continue
                }
            }
            else {
                # SYSTEMS SESSION QUERY PATH - Use /repserv_json QuerySessions
                $dataSessionsWithErrors = @{
                    id      = 1
                    jsonrpc = "2.0"
                    method  = "QuerySessions"
                    params  = @{
                        accountId = [int]$repsvr.AccountId
                        query     = $query
                        orderBy   = "BackupStartTime DESC"
                        range     = @{
                            Offset = 0
                            Size   = 5  # Reduced from 30 - status filter means we only get failures
                        }
                        account   = $repsvr.Name
                        token     = $repsvr.Token
                    }
                }
                
                # Try to query sessions, handle SSL errors for software-only clients
                try {
                    $responseErrors = Invoke-RestMethod -Uri $repsvr.repurl -Method POST -ContentType 'application/json' -Body ($dataSessionsWithErrors | ConvertTo-Json -Depth 10) -ErrorAction Stop
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $isSSLError = $errorMsg -like "*SSL*" -or $errorMsg -like "*secure channel*" -or $errorMsg -like "*TLS*"
                    
                    if ($isSSLError) {
                        # Extract RepURL to determine infrastructure type
                        $repurl = ($repsvr.repurl -replace 'https://', '' -replace '/repserv_json', '')
                        $isSoftwareOnly = $repurl -notmatch 'cloudbackup\.management'
                    
                    if ($isSoftwareOnly) {
                        # Software-only infrastructure - SSL errors are expected, skip gracefully
                        if ($DebugCDP) {
                            Write-Host " - Software-only (SSL expected)" -ForegroundColor Cyan
                        }
                        Write-Verbose "    Software-only infrastructure detected ($repurl) - SSL error expected, skipping datasource"
                        
                        # Log to SSL debug file
                        $debugOutput = @()
                        $debugOutput += "="*100
                        $debugOutput += "SSL ERROR - REPSERV (SOFTWARE-ONLY INFRASTRUCTURE) - EXPECTED"
                        $debugOutput += "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                        $debugOutput += "="*100
                        $debugOutput += "DEVICE: $($repsvr.Name) (AccountId: $($repsvr.AccountId))"
                        $debugOutput += "REPSERV URL: $repurl"
                        $debugOutput += "DATASOURCE: $datasourceName"
                        $debugOutput += "METHOD: QuerySessions"
                        $debugOutput += "ERROR: $errorMsg"
                        $debugOutput += "DIAGNOSTIC: Software-only infrastructure - SSL errors are EXPECTED"
                        $debugOutput += "="*100
                        $debugOutput += ""
                        $debugOutput | Out-File -FilePath $sslDebugLogPath -Append -Encoding UTF8
                        
                        continue  # Skip to next datasource
                    }
                    else {
                        # CloudBackup infrastructure - SSL error is unexpected
                        Write-Warning "SSL error querying CloudBackup repserv ($repurl) for $datasourceName - unexpected SSL issue"
                        
                        # Log to SSL debug file
                        $debugOutput = @()
                        $debugOutput += "="*100
                        $debugOutput += "SSL ERROR - REPSERV (CLOUDBACKUP INFRASTRUCTURE) - UNEXPECTED"
                        $debugOutput += "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                        $debugOutput += "="*100
                        $debugOutput += "DEVICE: $($repsvr.Name) (AccountId: $($repsvr.AccountId))"
                        $debugOutput += "REPSERV URL: $repurl"
                        $debugOutput += "DATASOURCE: $datasourceName"
                        $debugOutput += "METHOD: QuerySessions"
                        $debugOutput += "ERROR: $errorMsg"
                        $debugOutput += "FULL EXCEPTION:"
                        $debugOutput += $_.Exception.ToString()
                        $debugOutput += "DIAGNOSTIC: ⚠ CloudBackup infrastructure - SSL error is UNEXPECTED"
                        $debugOutput += "="*100
                        $debugOutput += ""
                        $debugOutput | Out-File -FilePath $sslDebugLogPath -Append -Encoding UTF8
                        
                        continue  # Skip to next datasource but log as unexpected
                    }
                }
                else {
                    # Non-SSL error - rethrow
                    throw
                }
                }  # End Systems catch block
            }  # End Systems else block
            
            $dsQueryElapsed = ((Get-Date) - $dsQueryStartTime).TotalMilliseconds
            
            if ($responseErrors.error) {
                if ($DebugCDP) {
                    Write-Host " - Error (" -ForegroundColor Yellow -NoNewline
                    Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                    Write-Host ")" -ForegroundColor Yellow
                }
                Write-Verbose "    API returned error: $($responseErrors.error.message)"
                continue
            }
            
            # Handle response format differences between M365 and Systems
            # CRITICAL: M365 sessions already extracted and filtered above in $allSessions
            # Don't overwrite with wrong path!
            $sessions = if ($AccountType -eq 2) {
                # M365: Use already-extracted and filtered $sessions from M365 path above
                $sessions
            } else {
                # Systems: repserv QuerySessions returns sessions in result.result (nested)
                @($responseErrors.result.result)
            }
            
            if (-not $sessions -or $sessions.Count -eq 0) {
                if ($DebugCDP) {
                    Write-Host " - No failed sessions (" -ForegroundColor Green -NoNewline
                    Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                    Write-Host ")" -ForegroundColor Green
                }
                Write-Verbose "    No sessions found for $datasourceName"
                continue
            }
            
            Write-Verbose "    Found $($sessions.Count) session(s) with errors, checking most recent..."
            
            # OPTIMIZATION v10: Only check MOST RECENT session for each datasource
            # Sort by EndTime descending (newest first), check only first session
            # If most recent session has no errors, datasource is healthy - skip older sessions
            $sortedSessions = $sessions | Sort-Object { if ($_.BackupStartTime) { $_.BackupStartTime } else { $_.EndTime } } -Descending
            
            $errorSession = $null
            $sessionErrors = $null
            $actualErrorCount = 0
            
            # Check only the FIRST (most recent) session
            $session = $sortedSessions[0]
            
            # Both M365 and Systems use 'Id' field for SessionId (confirmed from yesterday's working code)
            $sessionId = [int]$session.Id
            $sessionErrorCount = if ($session.ErrorsCount) { [int]$session.ErrorsCount } else { 0 }
            
            # Check if session is within date filter
            $startTimeValue = if ($session.BackupStartTime) { $session.BackupStartTime } else { $session.StartTime }
            $sessionStartTime = Convert-UnixTimeToDateTime $startTimeValue
            
            if ($sessionStartTime -lt $filterDate) {
                Write-Verbose "      Most recent session $sessionId - Outside date range, skipping datasource"
                if ($DebugCDP) {
                    Write-Host " - No errors in date range (" -ForegroundColor Cyan -NoNewline
                    Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                    Write-Host ")" -ForegroundColor Cyan
                }
                continue
            }
            
            Write-Verbose "      Checking most recent session: $sessionId ($(Get-SessionStatusText $session.Status))"
                
                # ARCHITECTURE SPLIT: M365 vs Systems error retrieval
                # - M365 (AccountType=2): Use reporting_api EnumerateSessionErrors (new Get-M365SessionErrors function)
                # - Systems (AccountType=1): Use repserv QueryErrors (existing logic)
                
                if ($AccountType -eq 2) {
                    # M365 ERROR RETRIEVAL PATH (NEW)
                    # Use reporting_api EnumerateSessionErrors via Get-M365SessionErrors function
                    Write-Host "      [DEBUG-M365] Session $sessionId - M365 device, using reporting_api" -ForegroundColor Magenta
                    Write-Verbose "      Session $sessionId - M365 device, using reporting_api EnumerateSessionErrors"
                    
                    # Get accountToken from repserv object (already populated by Get-AccountInfoById)
                    $accountToken = $repsvr.Token
                    
                    if (-not $accountToken) {
                        Write-Warning "      Session $sessionId - Missing accountToken for M365 error retrieval, skipping"
                        continue
                    }
                    
                    # Extract DataSourceType from session (Exchange, OneDrive, SharePoint, Teams)
                    # Datasource name comes from earlier loop: $datasourceName variable
                    $dataSourceType = $datasourceName
                    
                    # Call new M365 error retrieval function
                    $m365Errors = Get-M365SessionErrors -AccountId $repsvr.AccountId -AccountToken $accountToken -SessionId $sessionId -DataSourceType $dataSourceType -SessionType "Backup"
                    
                    if ($m365Errors -and $m365Errors.Count -gt 0) {
                        # Found M365 errors - use directly (already in correct format with Text, Time, Id)
                        $actualErrorCount = $m365Errors.Count
                        $sessionErrors = $m365Errors  # Already have Text, Time, Id properties from Get-M365SessionErrors
                        
                        Write-Verbose "      Session $sessionId ($datasourceName) - Retrieved $actualErrorCount M365 error(s) via reporting_api"
                        $errorSession = $session
                        # Found error session, use this data
                    } else {
                        Write-Verbose "      Session $sessionId - Most recent session has no errors, datasource is healthy"
                        # Most recent session has no errors - datasource is healthy, skip this datasource
                    }
                    
                } else {
                    # SYSTEMS ERROR RETRIEVAL PATH (EXISTING)
                    # Use repserv QueryErrors for AccountType=1 (Systems: Servers/Workstations)
                    Write-Verbose "      Session $sessionId - Systems device, using repserv QueryErrors"
                    
                    # Query error details ONCE and save results (OPTIMIZED - no duplicate call)
                    $dataErrorQuery = @{
                        id      = 1
                        jsonrpc = "2.0"
                        method  = "QueryErrors"
                        params  = @{
                            accountId = [int]$repsvr.AccountId
                            sessionId = $sessionId
                            query     = "SessionId == $sessionId"
                            orderBy   = "Time DESC"
                            groupId   = 0
                            account   = $repsvr.Name
                            token     = $repsvr.Token
                        }
                    }
                    
                    $responseErrorQuery = Invoke-RestMethod -Uri $repsvr.repurl -Method POST -ContentType 'application/json' -Body ($dataErrorQuery | ConvertTo-Json -Depth 10) -ErrorAction Stop
                    
                    if ($responseErrorQuery.result.result -and $responseErrorQuery.result.result.Count -gt 0) {
                        # Found a session with actual errors - save results for later use
                        $actualErrorCount = $responseErrorQuery.result.result.Count
                        $sessionErrors = $responseErrorQuery.result.result  # SAVE for later use
                        Write-Verbose "      Session $sessionId ($datasourceName) - Has $actualErrorCount error(s) (ErrorsCount field was $sessionErrorCount)"
                        $errorSession = $session
                        # Found error session, use this data
                    } else {
                        Write-Verbose "      Session $sessionId - Most recent session has no errors, datasource is healthy"
                        # Most recent session has no errors - datasource is healthy, skip this datasource
                    }
                }
            
            if (-not $errorSession) {
                if ($DebugCDP) {
                    Write-Host " - No errors in date range (" -ForegroundColor Cyan -NoNewline
                    Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                    Write-Host ")" -ForegroundColor Cyan
                }
                Write-Verbose "    No error sessions found for $datasourceName within date range"
                continue
            }
            
            # Use saved error details (OPTIMIZED - no duplicate QueryErrors call)
            # BUGFIX v10.2: Both M365 and Systems use 'Id' field (not SessionId) - QuerySessions response has "Id": 826
            $sessionId = [int]$errorSession.Id
            $startTimeValue = if ($errorSession.BackupStartTime) { $errorSession.BackupStartTime } else { $errorSession.StartTime }
            $sessionStartTime = Convert-UnixTimeToDateTime $startTimeValue
            
            # Error validation already done above when we saved $sessionErrors
            # No need to check again - we know we have valid error data
            
            Write-Verbose "      Retrieved $($sessionErrors.Count) error detail(s)"
            
            # Use the first error from this session
            $mostRecentError = $sessionErrors[0]
            
            # DEBUG: Log error object properties to understand structure
            if ($DebugCDP) {
                Write-Host "      [DEBUG] Error object properties:" -ForegroundColor Magenta
                $mostRecentError.PSObject.Properties | ForEach-Object {
                    Write-Host "        $($_.Name) = $($_.Value)" -ForegroundColor Gray
                }
            }
            
            # BUGFIX v10.3: Multi-level fallback for error timestamp (NotStarted sessions may have error.Time but not session times)
            # Priority: error.Time → session.EndTime → session.StartTime
            $errorTime = $null
            
            # Debug: Show what we're checking
            if ($DebugCDP) {
                Write-Host "      [DEBUG] Checking error.Time: $($mostRecentError.Time) (type: $($mostRecentError.Time.GetType().Name))" -ForegroundColor Magenta
            }
            
            # Try error.Time first (most accurate - when error actually occurred)
            if ($mostRecentError.Time -and $mostRecentError.Time -gt 0) {
                $errorTime = Convert-UnixTimeToDateTime $mostRecentError.Time
                Write-Verbose "      Using error.Time: $errorTime"
                if ($DebugCDP) { Write-Host "      [DEBUG] Using error.Time: $errorTime" -ForegroundColor Green }
            } else {
                if ($DebugCDP) { Write-Host "      [DEBUG] error.Time not available or <= 0" -ForegroundColor Red }
            }
            
            # Fallback to session EndTime
            if (-not $errorTime -or $errorTime -eq 'N/A') {
                $endTimeValue = if ($errorSession.BackupEndTime) { $errorSession.BackupEndTime } else { $errorSession.EndTime }
                if ($endTimeValue -and $endTimeValue -gt 0) {
                    $errorTime = Convert-UnixTimeToDateTime $endTimeValue
                    Write-Verbose "      No error.Time found, using session EndTime: $errorTime"
                    if ($DebugCDP) { Write-Host "      [DEBUG] Using session EndTime: $errorTime" -ForegroundColor Yellow }
                }
            }
            
            # Final fallback to session StartTime
            if (-not $errorTime -or $errorTime -eq 'N/A') {
                if ($startTimeValue -and $startTimeValue -gt 0) {
                    $errorTime = Convert-UnixTimeToDateTime $startTimeValue
                    Write-Verbose "      No EndTime found, using session StartTime: $errorTime"
                }
            }
            
            # Format error time with timezone label AND relative time
            if ($errorTime -and $errorTime -ne 'N/A') {
                if ($UseLocalTime) {
                    $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($errorTime, [System.TimeZoneInfo]::Local)
                    # Create abbreviation from timezone name (e.g., "Eastern Standard Time" -> "EST")
                    $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                        [System.TimeZoneInfo]::Local.DaylightName
                    } else {
                        [System.TimeZoneInfo]::Local.StandardName
                    }
                    $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                    # Add relative time
                    $now = (Get-Date).ToUniversalTime()
                    $timespan = $now - $errorTime
                    $hours = $timespan.TotalHours
                    $relativeTime = Format-HoursAsRelativeTime -Hours $hours -IncludeParentheses:$false
                    $errorTimeFormatted = "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr) ($relativeTime)"
                } else {
                    # Add relative time for UTC
                    $now = (Get-Date).ToUniversalTime()
                    $timespan = $now - $errorTime
                    $hours = $timespan.TotalHours
                    $relativeTime = Format-HoursAsRelativeTime -Hours $hours -IncludeParentheses:$false
                    $errorTimeFormatted = "$($errorTime.ToString('yyyy-MM-dd HH:mm:ss')) (UTC) ($relativeTime)"
                }
            } else {
                $errorTimeFormatted = "Unknown"
            }
            
            # Try to get detailed error info if this is a grouped error
            $errorDetails = $mostRecentError.Text
            if ($mostRecentError.Id -and $mostRecentError.Id -gt 0) {
                try {
                    $dataErrorGroup = @{
                        id      = 1
                        jsonrpc = "2.0"
                        method  = "EnumerateSessionErrorsGroupDetails"
                        params  = @{
                            accountId = [int]$repsvr.AccountId
                            sessionId = $sessionId
                            query     = "SessionId == $sessionId"
                            orderBy   = "Time DESC"
                            groupId   = [int]$mostRecentError.Id
                            account   = $repsvr.Name
                            token     = $repsvr.Token
                        }
                    }
                    
                    $responseGroupDetails = Invoke-RestMethod -Uri $repsvr.repurl -Method POST -ContentType 'application/json' -Body ($dataErrorGroup | ConvertTo-Json -Depth 10) -ErrorAction Stop
                    
                    if (-not $responseGroupDetails.error -and $responseGroupDetails.result.result -and $responseGroupDetails.result.result.Count -gt 0) {
                        $detailedError = $responseGroupDetails.result.result[0]
                        if ($detailedError.Filename) {
                            $errorDetails = "$($mostRecentError.Text) - File: $($detailedError.Filename)"
                            Write-Verbose "      Enhanced error with filename: $($detailedError.Filename)"
                        }
                    }
                }
                catch {
                    Write-Verbose "      Could not retrieve error group details: $($_.Exception.Message)"
                }
            }
            
            # Store error message (plain text for LastError field)
            # PHASE 4: Apply error cleaning for M365 Graph API errors
            if ($AccountType -eq 2) {
                # M365 error - apply Graph API error cleaning
                $cleanedErrorDetails = Format-CleanErrorMessage -ErrorMessage $errorDetails
                Write-Verbose "      Applied M365 error cleaning (Before: $($errorDetails.Length) chars, After: $($cleanedErrorDetails.Length) chars)"
                
                $errorMessage = if ($actualErrorCount -gt 1) {
                    "$cleanedErrorDetails [$actualErrorCount errors]"
                } else {
                    $cleanedErrorDetails
                }
            } else {
                # Systems error - use as-is (no Graph API wrapper)
                $errorMessage = if ($actualErrorCount -gt 1) {
                    "$errorDetails [$actualErrorCount errors]"
                } else {
                    $errorDetails
                }
            }
            
            # Get session status CODE (numeric, not text)
            $sessionStatusCode = if ($errorSession.StatusCode) { $errorSession.StatusCode } else { 2 }  # Default to Failed if not found
            
            # Store error with metadata (for LastSessionStatus and internal tracking)
            $errorWithMetadata = @{
                Message = $errorMessage
                Timestamp = $errorTimeFormatted
                SessionId = $sessionId
                Status = $sessionStatusCode
            }
            
            # Store error and session time for this datasource
            $datasourceErrors[$datasourceName] = $errorWithMetadata
            $datasourceSessionTimes[$datasourceName] = $sessionStartTime
            
            if ($DebugCDP) {
                Write-Host " - ❌ Error found (" -ForegroundColor Red -NoNewline
                Write-Host "$([math]::Round($dsQueryElapsed))ms" -ForegroundColor Gray -NoNewline
                Write-Host ")" -ForegroundColor Red
            }
            Write-Verbose "      Stored error for $datasourceName : $errorDetails"
            Write-Verbose "      Session BackupStartTime: $(Get-Date $sessionStartTime -Format 'yyyy-MM-dd HH:mm:ss')"
        }
        
        # Display session query timing summary
        $sessionQueryElapsed = ((Get-Date) - $sessionQueryStartTime).TotalSeconds
        if ($DebugCDP) {
            Write-Host "    ⏱️  Session queries completed in " -ForegroundColor Cyan -NoNewline
            Write-Host "$([math]::Round($sessionQueryElapsed, 2))s" -ForegroundColor White -NoNewline
            Write-Host " ($datasourceCount datasource(s), $($datasourceErrors.Count) with errors)" -ForegroundColor Gray -NoNewline
            Write-Host ""
        }
        
        if ($datasourceErrors.Count -eq 0) {
            Write-Verbose "DeviceId# $DeviceId - No errors found in recent sessions for any datasource"
            return $null
        }

        Write-Verbose "DeviceId# $DeviceId - Collected errors for $($datasourceErrors.Count) datasource(s):"
        foreach ($ds in $datasourceErrors.Keys) {
            Write-Verbose "  $ds : $($datasourceErrors[$ds])"
        }
        
        # Return both error messages and session times for proper comparison
        return @{
            Errors = $datasourceErrors
            SessionTimes = $datasourceSessionTimes
        }
    }
    catch {
        Write-Verbose "Get-LatestSessionError API Error for DeviceId $DeviceId : $($_.Exception.Message)"
        return $null
    }
}

Function Get-LastCompletedSessionErrors {
    <#
    .SYNOPSIS
        Checks last 30 days of completed sessions for errors in each datasource
    .DESCRIPTION
        Used when current session is InProcess - queries session history
        to find last completed session and check if it had errors.
        Uses EnumerateSessions for M365, QuerySessions for Systems.
    .PARAMETER AccountId
        Device Account ID
    .PARAMETER AccountToken
        Account token (required for M365 only)
    .PARAMETER DatasourcesToCheck
        Hashtable of datasource names to Plugin IDs (e.g., @{'Files'=@(28); 'Exchange'=@(4,19)})
    .PARAMETER IsM365
        True if checking M365 device, False for Systems
    .RETURNS
        Hashtable with datasource names as keys and boolean indicating if last session had errors
    #>
    Param (
        [Parameter(Mandatory=$True)] [int]$AccountId,
        [Parameter(Mandatory=$False)] [string]$AccountToken = "",
        [Parameter(Mandatory=$True)] [hashtable]$DatasourcesToCheck,
        [Parameter(Mandatory=$True)] [bool]$IsM365
    )
    
    try {
        $datasourceHadErrors = @{}
        $filterDate = (Get-Date).AddDays(-30)
        
        Write-Verbose "  InProcess Check - Querying last 30 days of sessions for Account $AccountId (M365: $IsM365)"
        
        if ($IsM365) {
            # M365: Use EnumerateSessions via management_api endpoint
            $now = [int][Math]::Floor((Get-Date).ToUniversalTime().Subtract((Get-Date "1970-01-01 00:00:00Z")).TotalSeconds)
            $createdAfter = $now - (30 * 86400)  # 30 days in seconds
            
            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $session.Cookies.Add((New-Object System.Net.Cookie("visa", $Script:Visa, "/", ".backup.management")))
            
            # Map datasource names to DataSources filter values
            $m365DataSourceMap = @{
                'Exchange' = 'Exchange'
                'OneDrive' = 'OneDrive'
                'SharePoint' = 'SharePoint'
                'Teams' = 'Teams'
            }
            
            foreach ($dsName in $DatasourcesToCheck.Keys) {
                $dataSourceFilter = $m365DataSourceMap[$dsName]
                
                $body = @{
                    method  = "EnumerateSessions"
                    params  = @{
                        accountToken = $AccountToken
                        range = @{
                            Offset = 0
                            Size = 1  # Only need last session
                        }
                        filter = @{
                            SessionTypes = @("Backup")
                            DataSources = @($dataSourceFilter)
                            CreatedAfter = $createdAfter
                            CreatedBefore = $now
                        }
                    }
                    jsonrpc = "2.0"
                    id      = "jsonrpc"
                }
                
                $response = Invoke-RestMethod `
                    -Uri "https://api.backup.management/management_api" `
                    -Method "POST" `
                    -WebSession $session `
                    -ContentType "application/json" `
                    -Body ($body | ConvertTo-Json -Depth 10) `
                    -ErrorAction Stop
                
                if (-not $response.result.result -or $response.result.result.Count -eq 0) {
                    Write-Verbose "    InProcess Check - No M365 sessions found for $dsName"
                    $datasourceHadErrors[$dsName] = $false
                    continue
                }
                
                $lastSession = $response.result.result[0]
                $sessionStatus = $lastSession.Status
                
                # Check if session had errors
                $hadErrors = ($sessionStatus -eq "CompletedWithErrors" -or $sessionStatus -eq "Failed")
                $datasourceHadErrors[$dsName] = $hadErrors
                
                Write-Verbose "    InProcess Check - M365 $dsName last session status: $sessionStatus (Errors: $hadErrors)"
            }
        } else {
            # Systems: Use QuerySessions via repserv endpoint (existing logic)
            $repsvr = Get-AccountInfoById -accountid $AccountId
            if (-not $repsvr -or -not $repsvr.Token) {
                Write-Verbose "  InProcess Check - Failed to get device token for Account $AccountId"
                return @{}
            }
            
            # Query each datasource separately
            foreach ($datasourceName in $DatasourcesToCheck.Keys) {
            $pluginIds = $DatasourcesToCheck[$datasourceName]
            
            # Build query for this datasource
            if ($pluginIds.Count -eq 1) {
                $pluginQuery = "Plugin == $($pluginIds[0])"
            } else {
                $pluginQuery = "(" + (($pluginIds | ForEach-Object { "Plugin == $_" }) -join " or ") + ")"
            }
            
            # Query recent sessions, excluding InProcess (Type 1)
            # Get last completed session only
            $query = "0 != 1 and $pluginQuery and Status != 1"
            
            $dataSessionsQuery = @{
                id      = 1
                jsonrpc = "2.0"
                method  = "QuerySessions"
                params  = @{
                    accountId = [int]$repsvr.AccountId
                    query     = $query
                    orderBy   = "BackupStartTime DESC"
                    range     = @{
                        Offset = 0
                        Size   = 1  # Only need last completed session
                    }
                    account   = $repsvr.Name
                    token     = $repsvr.Token
                }
            }
            
            $response = Invoke-RestMethod -Uri $repsvr.repurl -Method POST -ContentType 'application/json' -Body ($dataSessionsQuery | ConvertTo-Json -Depth 10) -ErrorAction Stop
            
            if ($response.error) {
                Write-Verbose "    InProcess Check - QuerySessions error for $datasourceName : $($response.error.message)"
                $datasourceHadErrors[$datasourceName] = $false
                continue
            }
            
            if (-not $response.result.result -or $response.result.result.Count -eq 0) {
                Write-Verbose "    InProcess Check - No completed sessions found for $datasourceName"
                $datasourceHadErrors[$datasourceName] = $false
                continue
            }
            
            $lastSession = $response.result.result[0]
            $sessionId = [int]$lastSession.Id
            $sessionStatus = $lastSession.Status
            $sessionErrorCount = if ($lastSession.ErrorsCount) { [int]$lastSession.ErrorsCount } else { 0 }
            
            # Check if session is within 30 day window
            $startTimeValue = if ($lastSession.BackupStartTime) { $lastSession.BackupStartTime } else { $lastSession.StartTime }
            $sessionStartTime = Convert-UnixTimeToDateTime $startTimeValue
            
            if ($sessionStartTime -lt $filterDate) {
                Write-Verbose "    InProcess Check - Last session for $datasourceName is older than 30 days"
                $datasourceHadErrors[$datasourceName] = $false
                continue
            }
            
            # Check if session had errors - verify with QueryErrors (ErrorsCount field unreliable)
            $dataErrorCheck = @{
                id      = 1
                jsonrpc = "2.0"
                method  = "QueryErrors"
                params  = @{
                    accountId = [int]$repsvr.AccountId
                    sessionId = $sessionId
                    query     = "SessionId == $sessionId"
                    orderBy   = "Time DESC"
                    groupId   = 0
                    account   = $repsvr.Name
                    token     = $repsvr.Token
                }
            }
            
            $errorResponse = Invoke-RestMethod -Uri $repsvr.repurl -Method POST -ContentType 'application/json' -Body ($dataErrorCheck | ConvertTo-Json -Depth 10) -ErrorAction Stop
            
            $hadErrors = ($errorResponse.result.result -and $errorResponse.result.result.Count -gt 0)
            $datasourceHadErrors[$datasourceName] = $hadErrors
            
            if ($hadErrors) {
                $errorCount = $errorResponse.result.result.Count
                Write-Verbose "    InProcess Check - $datasourceName last session (ID: $sessionId, Status: $sessionStatus) had $errorCount error(s)"
            } else {
                Write-Verbose "    InProcess Check - $datasourceName last session (ID: $sessionId, Status: $sessionStatus) had no errors"
            }
            }  # End foreach datasource
        }  # End else (Systems)
        
        return $datasourceHadErrors
    }
    catch {
        Write-Verbose "  InProcess Check - Error querying session history: $($_.Exception.Message)"
        return @{}
    }
}

#endregion ----- Error Retrieval Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator
    Write-Output "  Setting Backup API Credentials"
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential Path Present"
    } else {
        New-Item -ItemType Directory -Path $APIcredpath
    }

    Write-Output "  Enter Exact, Case Sensitive Partner Name for Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for Backup.Management API'
    
    # Create credential object in v10 XML format (matches CWM pattern)
    $CDPCredentials = [PSCustomObject]@{
        PartnerName = $PartnerName
        Username = $BackupCred.UserName
        Password = ($BackupCred.Password | ConvertFrom-SecureString)
    }

    # Save to XML (DPAPI encrypted)
    $CDPCredentials | Export-Clixml -Path $APIcredfile -Force
    
    Write-Output "  ✓ Credentials saved to: $APIcredfile"

    Start-Sleep -milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    $Script:APIcredfile = join-path -Path $True_Path -ChildPath "${env:computername}_${env:username}_API_Credentials.Secure.xml"
    $Script:APIcredpath = Split-path -path $APIcredfile

    if (($ClearCDPCredentials) -and (Test-Path $APIcredfile)) {
        Remove-Item -Path $Script:APIcredfile
        $ClearCDPCredentials = $Null
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential File Cleared"
        Send-APICredentialsCookie  ## Retry Authentication
    } else {
        Write-Output $Script:strLineSeparator
        Write-Output "  Getting Backup API Credentials"

        if (Test-Path $APIcredfile) {
            Write-Output    $Script:strLineSeparator
            "  Backup API Credential File Present"
            
            try {
                # Load XML credential file
                $APIcredentials = Import-Clixml -Path $APIcredfile -ErrorAction Stop

                $Script:cred0 = $APIcredentials.PartnerName
                $Script:cred1 = $APIcredentials.Username
                $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR(($APIcredentials.Password | ConvertTo-SecureString)))

                Write-Output    $Script:strLineSeparator
                Write-Output "  Stored Backup API Partner  = $Script:cred0"
                Write-Output "  Stored Backup API User     = $Script:cred1"
                Write-Output "  Stored Backup API Password = Encrypted"
            } catch {
                Write-Output    $Script:strLineSeparator
                Write-Warning "  Backup API Credential File is corrupted or invalid"
                Write-Output "  Error: $($_.Exception.Message)"
                Write-Output "  Removing corrupted file and prompting for new credentials..."
                
                # Delete corrupted file
                Remove-Item -Path $APIcredfile -Force -ErrorAction SilentlyContinue
                
                # Prompt for new credentials
                Set-APICredentials
            }

        } else {
            Write-Output    $Script:strLineSeparator
            Write-Output "  Backup API Credential File Not Present"

            Set-APICredentials  ## Create API Credential File if Not Found
        }
    }
}  ## Get API credentials if present

Function Send-APICredentialsCookie {

    Get-APICredentials  ## Read API Credential File before Authentication

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.partner = $Script:cred0
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    try {
        $Script:Authenticate = Invoke-RestMethod -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data) `
            -Uri $url `
            -TimeoutSec 30 `
            -SessionVariable Script:websession `
            -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession

        if ($DebugCDP) {
            Write-Host "`n[DEBUG] API Authentication Response:" -ForegroundColor Yellow
            $Script:Authenticate | ConvertTo-Json -Depth 5 | Write-Host
        }

        # Check for API error response
        if ($Script:Authenticate.error) {
            Write-Output    $Script:strLineSeparator
            Write-Output "  Authentication API Error:"
            Write-Output "  Code: $($Script:Authenticate.error.code)"
            Write-Output "  Message: $($Script:Authenticate.error.message)"
            Write-Output    $Script:strLineSeparator
            Set-APICredentials  ## Create API Credential File if Authentication Fails
            return
        }

        # Check for valid visa token
        if ($Script:Authenticate.visa) {
            $Script:visa = $Script:Authenticate.visa
            $Script:UserId = $Script:Authenticate.result.result.id
            Write-Output "  ✓ Authentication successful"
        } else {
            Write-Output    $Script:strLineSeparator
            Write-Output "  Authentication Failed: No visa token received"
            Write-Output "  Please confirm your Backup.Management Partner Name and Credentials"
            Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator

            if ($DebugCDP) {
                Write-Host "`n[DEBUG] Full authentication response:" -ForegroundColor Red
                $Script:Authenticate | ConvertTo-Json -Depth 10 | Write-Host
            }

            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }
    } catch {
        Write-Output    $Script:strLineSeparator
        Write-Output "  Authentication Network Error:"
        Write-Output "  $($_.Exception.Message)"
        Write-Output    $Script:strLineSeparator
        
        if ($DebugCDP) {
            Write-Host "`n[DEBUG] Exception Details:" -ForegroundColor Red
            $_ | Format-List * -Force | Out-String | Write-Host
        }
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Get-VisaTime {
    if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
            Send-APICredentialsCookie
        }
    }
}  ## Renew Visa

Function Get-EnumerateColumns {
    # Query API for available column codes and their descriptions
    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'EnumerateColumns'
    $data.params = @{}
    $data.params.partnerId = [int]$Script:PartnerId

    $jsondata = (ConvertTo-Json $data -depth 6)
    
    Write-Host "`n[EnumerateColumns] Querying available column codes..." -ForegroundColor Cyan
    
    $response = Invoke-RestMethod -Uri $url -Method POST -Body $jsondata -ContentType 'application/json; charset=utf-8'
    
    if ($response.result) {
        Write-Host "`nAvailable Columns (Total: $($response.result.Count)):" -ForegroundColor Green
        
        # Filter for timestamp-related columns
        Write-Host "`nTimestamp-related columns:" -ForegroundColor Yellow
        $response.result | Where-Object { $_.Description -match 'time|date|last|success|completed' } | 
            Sort-Object Code | Format-Table Code, Description -AutoSize | Out-String | Write-Host
        
        return $response.result
    } else {
        Write-Host "Error: $($response.error.message)" -ForegroundColor Red
        return $null
    }
}  ## Enumerate available column codes

#endregion ----- Authentication ----

#region ----- Data Conversion ----
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
        try {
            # Unix timestamps are typically 10 digits (seconds since 1970)
            # Valid range: ~0 to ~2147483647 (year 2038 for 32-bit systems)
            # Extended range: up to ~253402300799 (year 9999)
            if ($inputUnixTime -gt 253402300799) {
                Write-Verbose "Unix timestamp out of valid range: $inputUnixTime"
                return "N/A"
            }
            $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
            $epoch = $epoch.ToUniversalTime()
            $DateTime = $epoch.AddSeconds($inputUnixTime)
            return $DateTime
        }
        catch {
            Write-Verbose "Error converting Unix time $inputUnixTime : $_"
            return "N/A"
        }
    }else{ return "N/A"}
}  ## Convert epoch time to date time

Function Convert-DateTimeToUnixTime($DateToConvert) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $NewExtensionDate = Get-Date -Date $DateToConvert
    [int64]$NewEpoch = (New-TimeSpan -Start $epoch -End $NewExtensionDate).TotalSeconds
    Return $NewEpoch
}  ## Convert date time to epoch time

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

Function Send-GetPartnerInfo ($PartnerName) {

    if ($DebugCDP) {
        Write-Host "[DEBUG] Send-GetPartnerInfo called with: '$PartnerName'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Current Script:PartnerName='$Script:PartnerName'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Current Script:OriginalPartnerName='$Script:OriginalPartnerName'" -ForegroundColor Magenta
    }

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfo'
    $data.params = @{}
    $data.params.name = [String]$PartnerName

    $Script:Partner = Invoke-RestMethod -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
    $Script:websession = $websession

    if ($DebugCDP) { Write-Host "[DEBUG] API returned: Name='$($Partner.result.result.Name)', Id='$($Partner.result.result.Id)', Level='$($Partner.result.result.Level)'" -ForegroundColor Magenta }

    $RestrictedPartnerLevel = @("Root","SubRoot","Distributor")

    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$Script:Level = $Partner.result.result.Level
        # Only update PartnerName if not explicitly provided as script parameter
        if ([string]::IsNullOrWhiteSpace($Script:OriginalPartnerName)) {
            if ($DebugCDP) { Write-Host "[DEBUG]   OriginalPartnerName is empty - setting Script:PartnerName to API result: '$($Partner.result.result.Name)'" -ForegroundColor Magenta }
            [String]$Script:PartnerName = $Partner.result.result.Name
        } else {
            if ($DebugCDP) { Write-Host "[DEBUG]   OriginalPartnerName='$Script:OriginalPartnerName' - NOT overwriting Script:PartnerName" -ForegroundColor Magenta }
        }

        Write-Output $Script:strLineSeparator
        Write-Output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
    } else {
        Write-Output $Script:strLineSeparator
        Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
    }

    if ($partner.error) {
        Write-Output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
    }

    return $Partner
}  ## get PartnerID and Partner Level

Function Send-EnumerateAncestorPartners ($PartnerId) {
    # Get ancestor partners (walk up the hierarchy)
    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'EnumerateAncestorPartners'
    $data.params = @{}
    $data.params.partnerId = [int]$PartnerId

    $jsondata = (ConvertTo-Json $data -depth 6)
    
    $params = @{
        Uri         = $url
        Method      = 'POST'
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }
    
    $Script:Ancestors = Invoke-RestMethod @params
    
    return $Script:Ancestors
}  ## Get ancestor partners in hierarchy

Function Get-PartnerHierarchyInfo ($PartnerID, $PartnerName) {
    # Walk up partner tree to find End Customer and Site levels
    # Returns hashtable with EndCustomer and Site names
    # Uses cache to avoid redundant API calls
    
    if ($DebugCDP) {
        Write-Host "[DEBUG] Get-PartnerHierarchyInfo called: PartnerID=$PartnerID, PartnerName='$PartnerName'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Cache has $($Script:PartnerHierarchyCache.Count) entries" -ForegroundColor Magenta
    }
    
    # Check cache first
    if ($Script:PartnerHierarchyCache.ContainsKey($PartnerID)) {
        $cachedResult = $Script:PartnerHierarchyCache[$PartnerID]
        if ($DebugCDP) { Write-Host "[DEBUG]   CACHE HIT! Returning: EndCustomer='$($cachedResult.EndCustomer)', Site='$($cachedResult.Site)'" -ForegroundColor Yellow }
        return $cachedResult
    }
    
    if ($DebugCDP) { Write-Host "[DEBUG]   CACHE MISS - will query API" -ForegroundColor Cyan }
    
    $result = @{
        EndCustomer = $PartnerName  # Default to current partner
        Site = $null                # Site name if partner is a Site
        EndCustomerPartnerId = $PartnerID  # Partner ID of End Customer (for ExternalCode updates)
        EndCustomerLevel = $null    # Level of the EndCustomer partner (for validation)
    }
    
    try {
        $ancestors = Send-EnumerateAncestorPartners -PartnerId $PartnerID
        
        if ($DebugCDP) {
            Write-Host "[DEBUG]   EnumerateAncestorPartners returned $($ancestors.result.result.Count) ancestor(s)" -ForegroundColor Magenta
            foreach ($anc in $ancestors.result.result) {
                Write-Host "[DEBUG]     Ancestor: ID=$($anc.Id), Name='$($anc.Name)', Level='$($anc.Level)'" -ForegroundColor Magenta
            }
        }
        
        if ($ancestors.result.result) {
            # Ancestors are returned from immediate parent upward
            # Find the first EndCustomer level (skip Site levels)
            foreach ($ancestor in $ancestors.result.result) {
                if ($ancestor.Level -eq 'EndCustomer') {
                    # This partner is the End Customer
                    if ($DebugCDP) { Write-Host "[DEBUG]     SELECTING EndCustomer: '$($ancestor.Name)' (ID=$($ancestor.Id))" -ForegroundColor Yellow }
                    $result.EndCustomer = $ancestor.Name
                    $result.EndCustomerPartnerId = $ancestor.Id
                    $result.EndCustomerLevel = $ancestor.Level
                    # If we're walking up from a Site, the original partner is the Site
                    if ($PartnerName -ne $ancestor.Name) {
                        $result.Site = $PartnerName
                    }
                    break
                }
            }
        }
    }
    catch {
        # If API call fails, use original partner name
        Write-Verbose "Could not retrieve ancestor partners for $PartnerID"
    }
    
    # Cache the result
    $Script:PartnerHierarchyCache[$PartnerID] = $result
    if ($DebugCDP) { Write-Host "[DEBUG]   Caching for PartnerID $PartnerID : EndCustomer='$($result.EndCustomer)', Site='$($result.Site)'" -ForegroundColor Cyan }
    
    return $result
}  ## Resolve partner hierarchy to find End Customer and Site

Function Send-ModifyPartner {
    param(
        [Parameter(Mandatory=$true)] [int]$PartnerIdToModify,
        [Parameter(Mandatory=$false)] [string]$ExternalCode
    )
    
    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'ModifyPartner'
    $data.params = @{}
    $data.params.partnerInfo = @{}
    $data.params.partnerInfo.Id = $PartnerIdToModify
    
    if ($PSBoundParameters.ContainsKey('ExternalCode')) {
        $data.params.partnerInfo.ExternalCode = $ExternalCode
    }

    $jsondata = ConvertTo-Json $data -depth 5
    
    $params = @{
        Uri         = $url
        Method      = 'POST'
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }
    
    $result = Invoke-RestMethod @params
    return $result
}  ## Update partner ExternalCode with CWM Company ID

Function Update-CovePartnerExternalCode {
    param(
        [Parameter(Mandatory=$true)] [int]$CovePartnerId,
        [Parameter(Mandatory=$true)] [int]$CWMCompanyId,
        [Parameter(Mandatory=$true)] [string]$CWMCompanyIdentifier,
        [Parameter(Mandatory=$false)] [string]$CurrentExternalCode = "",
        [Parameter(Mandatory=$false)] [switch]$UpdateEnabled = $false
    )
    
    # Format: CWM:12345:TheGraphGroup (ID:Identifier)
    $cwmTag = "CWM:${CWMCompanyId}:${CWMCompanyIdentifier}"
    
    # Check if CWM ID already exists in ExternalCode
    if ($CurrentExternalCode -match "CWM:(\d+):?([^|]*)") {
        $existingCWMId = $matches[1]
        if ($existingCWMId -eq $CWMCompanyId) {
            # Check if identifier is also present and correct
            $existingIdentifier = $matches[2]
            if ($existingIdentifier -eq $CWMCompanyIdentifier) {
                # Already has correct CWM ID and identifier, no update needed
                return $true
            } else {
                # Update to include identifier or fix it (only if UpdateEnabled)
                if (-not $UpdateEnabled) {
                    return $true  # Match found, but updates disabled
                }
                $newExternalCode = $CurrentExternalCode -replace "CWM:\d+:?[^|]*", $cwmTag
            }
        } else {
            # Replace old CWM ID with new one (including identifier) - only if UpdateEnabled
            if (-not $UpdateEnabled) {
                return $true  # Match found, but updates disabled
            }
            $newExternalCode = $CurrentExternalCode -replace "CWM:\d+:?[^|]*", $cwmTag
        }
    } else {
        # Append CWM ID with identifier - only if UpdateEnabled
        if (-not $UpdateEnabled) {
            return $true  # No match, but updates disabled
        }
        if ([string]::IsNullOrWhiteSpace($CurrentExternalCode)) {
            $newExternalCode = $cwmTag
        } else {
            $newExternalCode = "$CurrentExternalCode | $cwmTag"
        }
    }
    
    try {
        $result = Send-ModifyPartner -PartnerIdToModify $CovePartnerId -ExternalCode $newExternalCode
        if ($result.error) {
            Write-Warning "Failed to update ExternalCode for Cove Partner ${CovePartnerId}: $($result.error.message)"
            return $false
        }
        
        # Track successful reference update
        $Script:ReferencesUpdated += [PSCustomObject]@{
            CovePartnerId = $CovePartnerId
            CWMCompanyId = $CWMCompanyId
            CWMCompanyIdentifier = $CWMCompanyIdentifier
            NewExternalCode = $newExternalCode
        }
        
        return $true
    }
    catch {
        Write-Warning "Failed to update ExternalCode for Cove Partner ${CovePartnerId}: $($_.Exception.Message)"
        return $false
    }
}  ## Update Cove partner ExternalCode with CWM Company ID link

Function Send-EnumeratePartners ($PartnerId) {
    # ----- Get Partners via EnumeratePartners -----
    
    # (Create the JSON object to call the EnumeratePartners function)
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumeratePartners'
        $data.params = @{}
        $data.params.parentPartnerId = $PartnerId
        
        $Script:EnumeratePartnersSession = Invoke-RestMethod -Method POST `
            -ContentType 'application/json; charset=utf-8' `
            -Body (ConvertTo-Json $data -depth 6) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
    # (Save the JSON response with the list of partners)
        $Script:Partnerslist = $EnumeratePartnersSession.result.result | Select-Object Id,Level,ExternalCode,PartnerId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
        
        $Script:Partnerslist | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
        $Script:Partnerslist | ForEach-Object { if ($_.TrialExpirationTime -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
        $Script:Partnerslist | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
    
        $Script:SelectedPartners = $Script:Partnerslist | Select-Object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | Out-GridView -Title "Current Partner | $($Partner.result.result.Level) | $($Partner.result.result.Name) | (Select Partners for Processing, or click Cancel to use the Current Partner)" -OutputMode Multiple
    
        if($null -eq $Script:SelectedPartners) {
            # (Manually set selected Partnrers)
            $Script:SelectedPartners = @()

            $Script:SelectedPartners += $Partnerslist | Select-Object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | Where-Object {$_.name -eq $PartnerName}
            
        }
    
}  ## EnumeratePartners API Call

Function Send-GetDevices {
    param(
        [Parameter(Mandatory=$true)][int]$AccountType,  # 1=Servers AND Workstations, 2=M365
        [Parameter(Mandatory=$true)][int]$StaleHours,
        [Parameter(Mandatory=$false)][int]$StaleHoursServers = 48,     # Used for server/workstation differentiation
        [Parameter(Mandatory=$false)][int]$StaleHoursWorkstations = 72, # Used for server/workstation differentiation
        [Parameter(Mandatory=$false)][int]$LookbackDays = 60
    )
    
    $deviceTypeName = if ($AccountType -eq 1) { "Servers and Workstations" } else { "M365 Tenants" }
    Write-Host "  Querying $deviceTypeName (AT==$AccountType)...`n" -ForegroundColor Cyan
    
    if ($DebugCDP) {
        Write-Host "[DEBUG] Send-GetDevices called:" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Script:PartnerId='$Script:PartnerId'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Script:PartnerName='$Script:PartnerName'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   Local PartnerId='$PartnerId'" -ForegroundColor Magenta
    }

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $visa
    $data.method = 'EnumerateAccountStatistics'
    $data.params = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId = [int]$PartnerId
    
    if ($DebugCDP) { Write-Host "[DEBUG]   Querying API with PartnerId=$PartnerId" -ForegroundColor Magenta }
    
    # OPTIMIZATION v08: Filter API query to only return devices with issues
    # BUILD FILTER PIECE BY PIECE:
    # Step 4: Add OR conditions for failures and stale devices (wrapped in outer parentheses)
    # Base filter: Account type and recent activity (TS timestamp within lookback period)
    $baseFilter = "(AT==$AccountType AND TS > ${LookbackDays}.day().ago())"
    
    # Build failure conditions based on device type
    # CRITICAL FIX: Only check datasource status if datasource is configured (I78)
    # Otherwise NULL/0 values for unconfigured datasources cause "!= 5" to match everything
    if ($AccountType -eq 1) {
        # System devices: Status failures OR errors OR stale
        $failureConditions = @(
            # Files status - check if D01 exists
            "(I78 =~ '*D01*' AND F0 != 5)",
            # SystemState status - check if D02 exists
            "(I78 =~ '*D02*' AND S0 != 5)",
            # SQL VSS - check if D03 exists
            "(I78 =~ '*D03*' AND Z0 != 5)",
            # Exchange - check if D04 exists
            "(I78 =~ '*D04*' AND X0 != 5)",
            # Network Shares - check if D06 exists
            "(I78 =~ '*D06*' AND N0 != 5)",
            # VMware - check if D08 exists
            "(I78 =~ '*D08*' AND W0 != 5)",
            # Hyper-V - check if D14 exists
            "(I78 =~ '*D14*' AND H0 != 5)",
            # MySQL - check if D15 exists
            "(I78 =~ '*D15*' AND L0 != 5)",
            # Error count checks (same logic)
            "(I78 =~ '*D01*' AND F7 > 0)",
            "(I78 =~ '*D02*' AND S7 > 0)",
            "(I78 =~ '*D03*' AND Z7 > 0)",
            "(I78 =~ '*D04*' AND X7 > 0)",
            "(I78 =~ '*D06*' AND N7 > 0)",
            "(I78 =~ '*D08*' AND W7 > 0)",
            "(I78 =~ '*D14*' AND H7 > 0)",
            "(I78 =~ '*D15*' AND L7 > 0)",
            # Stale checks - use actual parameter values
            "(OT == 2 AND TL < $StaleHoursServers.hour().ago())",
            "(OT == 1 AND TL < $StaleHoursWorkstations.hour().ago())"
        )
        $failureFilter = "(" + ($failureConditions -join ' OR ') + ")"
    } else {
        # M365 devices: Status failures OR errors OR stale
        $failureConditions = @(
            # Status failures (only if datasource configured)
            "(I78 =~ '*D19*' AND G0 != 5)",
            "(I78 =~ '*D20*' AND J0 != 5)",
            "(I78 =~ '*D5*' AND D5F0 != 5)",
            "(I78 =~ '*D23*' AND D23F0 != 5)",
            # Error checks
            "(I78 =~ '*D19*' AND G7 > 0)",
            "(I78 =~ '*D20*' AND J7 > 0)",
            "(I78 =~ '*D5*' AND D5F6 > 0)",
            "(I78 =~ '*D23*' AND D23F6 > 0)",
            # Stale - use actual parameter value (M365 uses $StaleHours parameter)
            "(TL < $StaleHours.hour().ago())"
        )
        $failureFilter = "(" + ($failureConditions -join ' OR ') + ")"
    }
    
    # Add test device filter if specified
    if ($TestDeviceName) {
        Write-Host "`n  TEST MODE: Filtering to single device: $TestDeviceName" -ForegroundColor Magenta
        $baseFilter = "$baseFilter AND (AN == `"$TestDeviceName`")"
    }
    
    $data.params.query.Filter = "$baseFilter AND $failureFilter"
    
    # Display the actual filter being used with detailed explanation
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                  COVE API DEVICE FILTER                       ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host "  Filter: $($data.params.query.Filter)" -ForegroundColor Yellow
    Write-Host "`n  Explanation:" -ForegroundColor White
    Write-Host "    • Base: AT==$AccountType (Account Type) AND TS > ${LookbackDays}.day().ago() (Activity in last $LookbackDays days)" -ForegroundColor Gray
    
    if ($AccountType -eq 1) {
        Write-Host "
    • System Failures (only checks datasources in I78):" -ForegroundColor White
        Write-Host "      - Status != 5 (Not Completed): F0(Files) S0(SystemState) Z0(SQL) X0(Exchange) N0(Shares) W0(VMware) H0(Hyper-V) L0(MySQL)" -ForegroundColor Gray
        Write-Host "      - Errors > 0: F7 S7 Z7 X7 N7 W7 H7 L7" -ForegroundColor Gray
        Write-Host "
    • Stale Thresholds:" -ForegroundColor White
        Write-Host "      - Servers (OT=2): > $StaleHoursServers hours since last backup" -ForegroundColor Gray
        Write-Host "      - Workstations (OT=1): > $StaleHoursWorkstations hours since last backup" -ForegroundColor Gray
    } else {
        Write-Host "
    • M365 Failures (only checks datasources in I78):" -ForegroundColor White
        Write-Host "      - Status != 5: G0(Exchange) J0(OneDrive) D5F0(SharePoint) D23F0(Teams)" -ForegroundColor Gray
        Write-Host "      - Errors > 0: G7 J7 D5F6 D23F6" -ForegroundColor Gray
        Write-Host "
    • Stale Threshold:" -ForegroundColor White
        Write-Host "      - M365 Tenants: > $StaleHoursM365 hours since last backup" -ForegroundColor Gray
    }
    Write-Host ""
    
    $data.params.query.Columns = @("AU","AR","AN","MN","AL","LN","OP","OI","OS","OT","PD","I78","PF","PN","CD","TS","TL","TZ","T3","US","I81","I84","I85","MF","MO","AA843","T0","T7","ER","F0","FL","FO","FQ","FJ","F3","F4","F5","F6","F7","FA","S0","SL","SO","SQ","SJ","S3","S4","S5","S6","S7","SA","Z0","ZL","ZO","ZQ","ZJ","Z3","Z4","Z5","Z6","Z7","ZA","D5F0","D5F1","D5F9","D5F16","D5F17","D5F18","D5F20","D5F22","D5F3","D5F4","D5F5","D5F6","D5F12","D19F0","G1","GM","G@","GL","GO","GQ","GJ","G3","G4","G5","G6","G7","GA","D20F0","J1","JM","JL","JO","JQ","JJ","J3","J4","J5","J6","J7","JA","D23F0","D23F1","D23F9","D23F11","D23F17","D23F18","D23F20","D23F23","D23F24","D23F25","D23F3","D23F4","D23F5","D23F6","D23F12","XL","XO","XQ","XJ","XA","X3","X4","X5","X7","X0","NL","NO","NQ","NJ","NA","N3","N4","N5","N7","N0","WL","WO","WQ","WJ","WA","W3","W4","W5","W7","W0","HL","HO","HQ","HJ","HA","H3","H4","H5","H7","H0","LL","LO","LQ","LJ","LA","L3","L4","L5","L7","L0","IP","EI","AA3147","AA3347","AA3148","AA3149","AA3150")<# Oracle deprecated: "YL","YO","YQ","YJ","YA","Y3","Y4","Y5","Y7","Y0" #><# Linux SystemState deprecated: "K0","KL","KO","KQ","KJ","K3","K4","K5","K6","K7","KA" #>
    $data.params.query.OrderBy = "CD DESC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = $DeviceCount
    $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")

    $jsondata = (ConvertTo-Json $data -depth 6)
    
    # PERFORMANCE TRACKING v06: Measure device statistics query time
    $deviceQueryStart = Get-Date
    
    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }
    $Script:DeviceResponse = Invoke-RestMethod @params

    # Initialize DeviceDetail array if not already initialized (first call)
    if ($null -eq $Script:DeviceDetail) {
        $Script:DeviceDetail = @()
    }

    # Track starting count for this section
    $StartingCount = $Script:DeviceDetail.Count

    # Track device index for test mode simulation
    $Script:DeviceIndex = 0
    
    # Initialize partner hierarchy cache (reset each run to prevent stale data from previous executions)
    # CRITICAL v09: Force new hashtable object creation to clear any persisted data
    if ($DebugCDP) { Write-Host "[DEBUG] Clearing PartnerHierarchyCache (had $($Script:PartnerHierarchyCache.Count) entries)" -ForegroundColor Magenta }
    $Script:PartnerHierarchyCache = @{}
    if ($DebugCDP) {
        Write-Host "[DEBUG] PartnerHierarchyCache cleared - now has $($Script:PartnerHierarchyCache.Count) entries" -ForegroundColor Magenta
        Write-Host "[DEBUG] PartnerHierarchyCache now has $($Script:PartnerHierarchyCache.Count) entries" -ForegroundColor Magenta
    }
    
    # PERFORMANCE TRACKING v06: Record device query completion time
    $Script:PerformanceMetrics.DeviceQueryTime = ((Get-Date) - $deviceQueryStart).TotalMilliseconds
    
    # DEBUG: List all devices returned by API filter for manual verification
    if ($debugCDP) {
        Write-Host "`n  === DEVICES RETURNED BY API FILTER ===" -ForegroundColor Cyan
        Write-Host "  Total devices returned: $($DeviceResponse.result.result.Count)" -ForegroundColor White
        foreach ($dev in $DeviceResponse.result.result) {
            $devName = $dev.Settings.AN -join ''
            $devF0 = $dev.Settings.F0 -join ''
            $devS0 = $dev.Settings.S0 -join ''
            $devT0 = $dev.Settings.T0 -join ''
            $devF7 = $dev.Settings.F7 -join ''
            $devS7 = $dev.Settings.S7 -join ''
            $devTL = $dev.Settings.TL -join ''
            $devI78 = $dev.Settings.I78 -join ''
            $devOT = $dev.Settings.OT -join ''
            
            # Calculate hours since last session
            $tlDate = if ($devTL -and $devTL -ne '' -and [int64]$devTL -gt 0) {
                $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
                $epoch.ToUniversalTime().AddSeconds([int64]$devTL)
            } else { $null }
            $hoursSinceTL = if ($tlDate) {
                [Math]::Round(((Get-Date).ToUniversalTime() - $tlDate).TotalHours, 1)
            } else { 'N/A' }
            
            Write-Host "    - $devName | T0=$devT0 F0=$devF0 S0=$devS0 | F7=$devF7 S7=$devS7 | TL=$hoursSinceTL hrs | DS=$devI78 | OT=$devOT" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Sort devices by PartnerID to maximize cache hits
    $SortedDevices = $DeviceResponse.result.result | Sort-Object -Property PartnerId
    
    # Initialize device counter for progress display
    $currentDeviceNum = 0
    $totalDevices = $SortedDevices.Count
    
    ForEach ( $DeviceResult in $SortedDevices ) {
        $currentDeviceNum++
        if ($debugCDP) {Write-Output "Getting statistics data for deviceid | $($DeviceResult.AccountId)"}
        
        # CRITICAL: Initialize $partnerHierarchy to prevent carrying over values from previous device iteration
        $partnerHierarchy = @{
            EndCustomer = ''
            EndCustomerLevel = ''
            Site = ''
            EndCustomerPartnerId = ''
        }
        
        # Determine device status and severity
        $LastSuccess = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '')
        $LastSessionStatusCode = $DeviceResult.Settings.T0 -join ''
        $LastSessionStatus = Get-SessionStatusText -StatusCode $LastSessionStatusCode
        $LastError = $DeviceResult.Settings.ER -join ''
        $DataSourceString = $DeviceResult.Settings.I78 -join ''  # Configured datasources
        
        # Get creation date early - needed for stale check on never-backed-up devices
        $creationDate = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '')
        
        # Check each datasource individually for stale status (both M365 and Systems)
        $HoursSinceSuccess = 0
        $StaleDatasources = @()
        
        if ($AccountType -eq 2) {
            # M365: Check each datasource individually
            # Format: DatasourceName = @{ LastSuccess = 'ColumnCode'; LastCompleted = 'ColumnCode'; Duration = 'ColumnCode'; DatasourceCode = 'D##' }
            # CRITICAL: Only evaluate datasources that are actually configured (present in I78)
            $allM365Datasources = @{
                'Exchange' = @{ LastSuccess = 'GL'; LastCompleted = 'GO'; Duration = 'GA'; DatasourceCode = 'D19' }
                'OneDrive' = @{ LastSuccess = 'JL'; LastCompleted = 'JO'; Duration = 'JA'; DatasourceCode = 'D20' }
                'SharePoint' = @{ LastSuccess = 'D5F9'; LastCompleted = 'D5F18'; Duration = 'D5F12'; DatasourceCode = 'D5' }
                'Teams' = @{ LastSuccess = 'D23F9'; LastCompleted = 'D23F18'; Duration = 'D23F12'; DatasourceCode = 'D23' }
            }
            
            # Filter to only configured datasources (present in I78)
            $m365Datasources = @{}
            foreach ($dsName in $allM365Datasources.Keys) {
                $dsCode = $allM365Datasources[$dsName].DatasourceCode
                if ($DataSourceString -match $dsCode) {
                    $m365Datasources[$dsName] = $allM365Datasources[$dsName]
                }
            }
            
            # If no M365 datasources are configured, skip M365 evaluation entirely
            if ($m365Datasources.Count -eq 0) {
                if ($DebugCDP) { Write-Host "  [DEBUG] M365 device has no configured datasources in I78: $DataSourceString" -ForegroundColor Magenta }
                # Set as successful with no issues
                $IssueDetected = $false
                $LastSessionStatusText = "No datasources configured"
                continue  # Skip to next device
            }
            
            $maxHoursSinceSuccess = 0
            $minHoursSinceSuccess = 999999
            $worstDatasourceName = "Unknown"
            $bestDatasourceName = "Unknown"
            $worstDatasourceTime = $null
            $bestDatasourceTime = $null
            $worstDatasourceStatus = $null
            $bestDatasourceStatus = $null
            $StaleDatasources = @()
            
            # HYBRID INPROCESS CHECK: If session is InProcess, check last completed session
            $isInProcess = ($LastSessionStatusCode -eq '1')
            $sessionHistoryErrors = @{}  # Will store per-datasource error flags from history
            
            if ($isInProcess) {
                Write-Verbose "  M365 Device $($DeviceResult.Settings.AN -join '') is InProcess - checking last completed sessions"
                
                # Get M365 AccountToken from GetAccountInfoById
                $accountId = ($DeviceResult.Settings.AU -join '')
                $accountInfo = Get-AccountInfoById -accountid $accountId
                
                if ($accountInfo.result -and $accountInfo.result.result.Token) {
                    $m365AccountToken = $accountInfo.result.result.Token
                    
                    # Map M365 datasources to names for EnumerateSessions filter
                    $m365DataSourceMap = @{
                        'Exchange' = 'Exchange'
                        'OneDrive' = 'OneDrive'
                        'SharePoint' = 'SharePoint'
                        'Teams' = 'Teams'
                    }
                    
                    $sessionHistoryErrors = Get-LastCompletedSessionErrors -AccountId $accountId -AccountToken $m365AccountToken -DatasourcesToCheck $m365DataSourceMap -IsM365 $true
                } else {
                    Write-Verbose "  Failed to get AccountToken for M365 device $accountId - skipping InProcess check"
                }
            }
            
            foreach ($dsName in $m365Datasources.Keys) {
                $dsConfig = $m365Datasources[$dsName]
                $dsLastSuccessColumn = $dsConfig.LastSuccess
                $dsLastCompletedColumn = $dsConfig.LastCompleted
                $dsDurationColumn = $dsConfig.Duration
                $dsLastSuccess = Convert-UnixTimeToDateTime ($DeviceResult.Settings.$dsLastSuccessColumn -join '')
                $dsLastCompleted = Convert-UnixTimeToDateTime ($DeviceResult.Settings.$dsLastCompletedColumn -join '')
                $dsDurationSeconds = [int]($DeviceResult.Settings.$dsDurationColumn -join '')
                
                # Get status code for this datasource to verify it's truly successful
                $dsStatusColumn = switch ($dsName) {
                    'Exchange' { 'D19F0' }
                    'OneDrive' { 'D20F0' }
                    'SharePoint' { 'D5F0' }
                    'Teams' { 'D23F0' }
                }
                $dsStatusCode = $DeviceResult.Settings.$dsStatusColumn -join ''
                
                # Determine if datasource is truly successful
                # - If status == 5 (Completed): TRUE success
                # - If status == 8 (CompletedWithErrors): FAILURE
                # - If status == 1 (InProcess) with 0 current errors: Check session history
                $isTrueSuccess = $false
                
                if ($dsStatusCode -eq '5') {
                    # Explicitly successful
                    $isTrueSuccess = $true
                } elseif ($dsStatusCode -eq '1' -and $sessionHistoryErrors.ContainsKey($dsName)) {
                    # InProcess - use session history result
                    $lastSessionHadErrors = $sessionHistoryErrors[$dsName]
                    $isTrueSuccess = -not $lastSessionHadErrors
                    
                    if ($lastSessionHadErrors) {
                        Write-Verbose "  M365 $dsName is InProcess but last completed session had errors - treating as failure"
                    } else {
                        Write-Verbose "  M365 $dsName is InProcess and last completed session was clean - treating as success"
                    }
                } elseif ($dsStatusCode -eq '1') {
                    # InProcess but no history check (shouldn't happen if T7==0) - treat as success
                    $isTrueSuccess = $true
                }
                # All other statuses (2,3,4,6,8,9,etc.) remain false (failures)
                
                if ($dsLastSuccess -and $dsLastSuccess -ne 'N/A' -and $dsLastSuccess -is [DateTime]) {
                    $dsHoursSinceSuccess = ((Get-Date).ToUniversalTime() - $dsLastSuccess).TotalHours
                    
                    # Track which datasources are stale (just name, no hours - time shown in summary)
                    if ($dsHoursSinceSuccess -gt $StaleHours) {
                        $StaleDatasources += $dsName
                    }
                    
                    # Always track the maximum hours for stale detection (regardless of status)
                    if ($dsHoursSinceSuccess -gt $maxHoursSinceSuccess) {
                        $maxHoursSinceSuccess = $dsHoursSinceSuccess
                    }
                    
                    # Track best/worst datasource NAMES, TIMES, and STATUSES based on current status
                    if ($isTrueSuccess) {
                        # This datasource is healthy - track as potential "best"
                        if ($dsHoursSinceSuccess -lt $minHoursSinceSuccess) {
                            $minHoursSinceSuccess = $dsHoursSinceSuccess
                            $bestDatasourceName = $dsName
                            $bestDatasourceTime = $dsLastSuccess
                            $bestDatasourceStatus = $dsStatusCode
                        }
                    } else {
                        # This datasource has problems (status != 5) - track as potential "worst" using Last Completed time
                        $worstDatasourceName = $dsName
                        $worstDatasourceTime = if ($dsLastCompleted -and $dsLastCompleted -ne 'N/A') { $dsLastCompleted } else { $dsLastSuccess }
                        $worstDatasourceStatus = $dsStatusCode
                    }
                } else {
                    # Datasource not configured (no last success) - check if it has a session start time
                    $dsStatusColumn = switch ($dsName) {
                        'Exchange' { 'D19F0' }
                        'OneDrive' { 'D20F0' }
                        'SharePoint' { 'D5F0' }
                        'Teams' { 'D23F0' }
                    }
                    $dsStatus = $DeviceResult.Settings.$dsStatusColumn -join ''
                    
                    # Only flag as stale if datasource is configured but has no success
                    if ($dsStatus -and $dsStatus -ne '0') {
                        $StaleDatasources += "$dsName (never successful)"
                        $maxHoursSinceSuccess = 999999  # Force stale classification
                    }
                }
            }
            
            $HoursSinceSuccess = $maxHoursSinceSuccess
        } else {
            # Servers/Workstations: Check each datasource individually
            # Format: DatasourceName = @{ LastSuccess = 'ColumnCode'; LastCompleted = 'ColumnCode'; Duration = 'ColumnCode' }
            $systemDatasources = @{
                'FileSystem' = @{ LastSuccess = 'FL'; LastCompleted = 'FO'; Duration = 'FA' }
                'SystemState' = @{ LastSuccess = 'SL'; LastCompleted = 'SO'; Duration = 'SA' }
                'MSSQL' = @{ LastSuccess = 'ZL'; LastCompleted = 'ZO'; Duration = 'ZA' }
                'ExchangeVSS' = @{ LastSuccess = 'XL'; LastCompleted = 'XO'; Duration = 'XA' }
                'NetworkShares' = @{ LastSuccess = 'NL'; LastCompleted = 'NO'; Duration = 'NA' }
                'VMware' = @{ LastSuccess = 'WL'; LastCompleted = 'WO'; Duration = 'WA' }
                'HyperV' = @{ LastSuccess = 'HL'; LastCompleted = 'HO'; Duration = 'HA' }
                'MySQL' = @{ LastSuccess = 'LL'; LastCompleted = 'LO'; Duration = 'LA' }
            }
            
            $maxHoursSinceSuccess = 0
            $minHoursSinceSuccess = 999999
            $worstDatasourceName = "Unknown"
            $bestDatasourceName = "Unknown"
            $worstDatasourceTime = $null
            $bestDatasourceTime = $null
            $worstDatasourceStatus = $null
            $bestDatasourceStatus = $null
            $hasAnyDatasource = $false
            
            # HYBRID INPROCESS CHECK: If session is InProcess, check last completed session
            $isInProcess = ($LastSessionStatusCode -eq '1')
            $sessionHistoryErrors = @{}  # Will store per-datasource error flags from history
            
            if ($isInProcess) {
                Write-Verbose "  Systems Device $($DeviceResult.Settings.AN -join '') is InProcess - checking last completed sessions"
                
                # Map Systems datasources to Plugin IDs for QuerySessions
                $systemsPluginMap = @{
                    'FileSystem' = @(28)
                    'SystemState' = @(1)
                    'MSSQL' = @(10, 11)  # Direct or VSS
                    'ExchangeVSS' = @(3)
                    'NetworkShares' = @(6)
                    'VMware' = @(8)
                    'HyperV' = @(14)
                    'MySQL' = @(15)
                }
                
                # OPTIMIZATION v08: Filter to only configured datasources from I78
                $i78Value = $DeviceResult.Settings.I78 -join ''
                if ($i78Value -and $i78Value -ne "") {
                    $dsCodeToName = @{
                        "D01" = "FileSystem"
                        "D02" = "SystemState"
                        "D03" = "MSSQL"
                        "D04" = "ExchangeVSS"
                        "D06" = "NetworkShares"
                        "D08" = "VMware"
                        "D14" = "HyperV"
                        "D15" = "MySQL"
                    }
                    
                    $configuredDatasources = @{}
                    foreach ($code in $dsCodeToName.Keys) {
                        if ($i78Value -like "*$code*") {
                            $dsName = $dsCodeToName[$code]
                            if ($systemsPluginMap.ContainsKey($dsName)) {
                                $configuredDatasources[$dsName] = $systemsPluginMap[$dsName]
                            }
                        }
                    }
                    $systemsPluginMap = $configuredDatasources
                    Write-Verbose "    Filtered to $($systemsPluginMap.Count) configured datasource(s) from I78: $i78Value"
                }
                
                $sessionHistoryErrors = Get-LastCompletedSessionErrors -AccountId ($DeviceResult.Settings.AU -join '') -DatasourcesToCheck $systemsPluginMap -IsM365 $false
            }
            
            foreach ($dsName in $systemDatasources.Keys) {
                $dsConfig = $systemDatasources[$dsName]
                $dsLastSuccessColumn = $dsConfig.LastSuccess
                $dsLastCompletedColumn = $dsConfig.LastCompleted
                $dsDurationColumn = $dsConfig.Duration
                $dsLastSuccess = Convert-UnixTimeToDateTime ($DeviceResult.Settings.$dsLastSuccessColumn -join '')
                $dsLastCompleted = Convert-UnixTimeToDateTime ($DeviceResult.Settings.$dsLastCompletedColumn -join '')
                $dsDurationSeconds = [int]($DeviceResult.Settings.$dsDurationColumn -join '')
                
                # Get status code for this datasource to verify it's truly successful
                $dsStatusColumn = switch ($dsName) {
                    'FileSystem' { 'F0' }
                    'SystemState' { 'S0' }
                    'MSSQL' { 'Z0' }
                    'ExchangeVSS' { 'X0' }
                    'NetworkShares' { 'N0' }
                    'VMware' { 'W0' }
                    'HyperV' { 'H0' }
                    'MySQL' { 'L0' }
                }
                $dsStatusCode = $DeviceResult.Settings.$dsStatusColumn -join ''
                
                # Determine if datasource is truly successful
                # - If status == 5 (Completed): TRUE success
                # - If status == 8 (CompletedWithErrors): FAILURE
                # - If status == 1 (InProcess) with 0 current errors: Check session history
                $isTrueSuccess = $false
                
                if ($dsStatusCode -eq '5') {
                    # Explicitly successful
                    $isTrueSuccess = $true
                } elseif ($dsStatusCode -eq '1' -and $sessionHistoryErrors.ContainsKey($dsName)) {
                    # InProcess - use session history result
                    $lastSessionHadErrors = $sessionHistoryErrors[$dsName]
                    $isTrueSuccess = -not $lastSessionHadErrors
                    
                    if ($lastSessionHadErrors) {
                        Write-Verbose "  Systems $dsName is InProcess but last completed session had errors - treating as failure"
                    } else {
                        Write-Verbose "  Systems $dsName is InProcess and last completed session was clean - treating as success"
                    }
                } elseif ($dsStatusCode -eq '1') {
                    # InProcess but no history check (shouldn't happen if T7==0) - treat as success
                    $isTrueSuccess = $true
                }
                # All other statuses (2,3,4,6,8,9,etc.) remain false (failures)
                
                if ($dsLastSuccess -and $dsLastSuccess -ne 'N/A' -and $dsLastSuccess -is [DateTime]) {
                    $hasAnyDatasource = $true
                    $dsHoursSinceSuccess = ((Get-Date).ToUniversalTime() - $dsLastSuccess).TotalHours
                    
                    # Always track the maximum hours for stale detection (regardless of status)
                    if ($dsHoursSinceSuccess -gt $maxHoursSinceSuccess) {
                        $maxHoursSinceSuccess = $dsHoursSinceSuccess
                    }
                    
                    # Track best/worst datasource NAMES, TIMES, and STATUSES based on current status
                    if ($isTrueSuccess) {
                        # This datasource is healthy - track as potential "best" using Last Success time
                        if ($dsHoursSinceSuccess -lt $minHoursSinceSuccess) {
                            $minHoursSinceSuccess = $dsHoursSinceSuccess
                            $bestDatasourceName = $dsName
                            $bestDatasourceTime = $dsLastSuccess
                            $bestDatasourceStatus = $dsStatusCode
                        }
                    } else {
                        # This datasource has problems (status != 5) - track as potential "worst" using Last Completed time
                        $worstDatasourceName = $dsName
                        $worstDatasourceTime = if ($dsLastCompleted -and $dsLastCompleted -ne 'N/A') { $dsLastCompleted } else { $dsLastSuccess }
                        $worstDatasourceStatus = $dsStatusCode
                    }
                }
            }
            
            # If no datasources found, fall back to device-level LastSuccess
            if ($hasAnyDatasource) {
                $HoursSinceSuccess = $maxHoursSinceSuccess
            } else {
                $HoursSinceSuccess = if ($LastSuccess -and $LastSuccess -is [DateTime]) { 
                    ((Get-Date).ToUniversalTime() - $LastSuccess).TotalHours 
                } elseif ($creationDate -and $creationDate -is [DateTime]) {
                    # No prior success - calculate from Creation Date instead of using 999999
                    ((Get-Date).ToUniversalTime() - $creationDate).TotalHours
                } else { 
                    999999 
                }
            }
        }
        
        # Determine appropriate stale threshold based on device type (server vs workstation)
        $DeviceStaleHours = $StaleHours  # Default to the passed-in value (M365 uses this)
        if ($AccountType -eq 1) {
            # For servers/workstations, check if it's a server or workstation
            $DeviceOSType = $DeviceResult.Settings.OT -join ''
            $DeviceOS = $DeviceResult.Settings.OS -join ''
            $isDeviceServer = ($DeviceOSType -eq "2") -or ($DeviceOS -match "Server")
            
            $DeviceStaleHours = if ($isDeviceServer) { $StaleHoursServers } else { $StaleHoursWorkstations }
        }
        
        # Determine issue severity with device-type-based defaults
        # Servers default to Critical, M365 to Warning (Medium), Workstations to Stale (Low)
        $IssueSeverity = "Success"
        $IssueDescription = "Backup successful"
        
        # Check if any datasource is InProcess (status code "1")
        # Track this but don't change severity yet - let error detection upgrade it if needed
        $anyInProcess = $false
        
        if ($AccountType -eq 2) {
            # M365: Check G0, J0, D5F0, D23F0
            $inProcessStatuses = @(
                ($DeviceResult.Settings.G0 -join ''),
                ($DeviceResult.Settings.J0 -join ''),
                ($DeviceResult.Settings.D5F0 -join ''),
                ($DeviceResult.Settings.D23F0 -join '')
            )
            $anyInProcess = $inProcessStatuses -contains '1'
        } else {
            # Servers/Workstations: Check F0, S0, Q0, X0, W0, H0
            $inProcessStatuses = @(
                ($DeviceResult.Settings.F0 -join ''),
                ($DeviceResult.Settings.S0 -join ''),
                ($DeviceResult.Settings.Q0 -join ''),
                ($DeviceResult.Settings.X0 -join ''),
                ($DeviceResult.Settings.W0 -join ''),
                ($DeviceResult.Settings.H0 -join '')
            )
            $anyInProcess = $inProcessStatuses -contains '1'
        }
        
        # BUGFIX v10.3: For stale check, use the WORST datasource time ($HoursSinceSuccess = $maxHoursSinceSuccess)
        # NOT the device-level $LastSuccess (TL), which may be newer than individual datasources
        # This ensures the summary time matches the worst individual datasource shown in ticket details
        $hoursToCheck = if ($HoursSinceSuccess -lt 999999) {
            $HoursSinceSuccess  # Use calculated worst datasource time
        } elseif ($creationDate -and $creationDate -is [DateTime]) {
            ((Get-Date).ToUniversalTime() - $creationDate).TotalHours  # Never backed up: hours since creation
        } else {
            $HoursSinceSuccess  # Fallback to 999999 (will trigger stale)
        }
        
        if ($hoursToCheck -gt $DeviceStaleHours) {
            # Determine base severity based on device type
            # NOTE: "Stale" means no successful backup in X hours, NOT an actual backup failure
            # Actual failures are determined by session status codes (2,4,6,8,9) checked later
            
            # Check if this is a never-backed-up device
            $isNeverBackedUp = (-not $LastSuccess -or $LastSuccess -eq 'N/A')
            
            # Use unified Format-HoursAsRelativeTime function for consistency
            # Pass -NeverBackedUp flag for appropriate messaging, -IncludeParentheses:$false for description text
            $staleTimeFormatted = Format-HoursAsRelativeTime -Hours $hoursToCheck -NeverBackedUp:$isNeverBackedUp -IncludeParentheses:$false
            
            # Build datasource list for all device types
            $datasourceList = @()
            if ($AccountType -eq 2) {
                # M365: Use stale datasource list
                $datasourceList = $StaleDatasources
            } else {
                # Server/Workstation: Extract stale datasource names
                foreach ($dsName in $systemDatasources.Keys) {
                    $colCode = $systemDatasources[$dsName]
                    $dsLastSuccess = $Device.$colCode
                    if ($dsLastSuccess) {
                        $dsHoursSinceSuccess = [Math]::Round(((Get-Date).ToUniversalTime() - $dsLastSuccess).TotalHours, 1)
                        if ($dsHoursSinceSuccess -ge $StaleHours) {
                            $datasourceList += $dsName
                        }
                    }
                }
            }
            
            # Generate unified issue description
            if ($datasourceList.Count -eq 1) {
                $IssueDescription = "Stale backup - $($datasourceList[0]) ($staleTimeFormatted)"
            } elseif ($datasourceList.Count -gt 1) {
                $IssueDescription = "Stale backup - Multiple datasources ($staleTimeFormatted)"
            } else {
                # Fallback if no datasources identified
                $IssueDescription = "Stale backup ($staleTimeFormatted)"
            }
            
            # Set severity based on device type
            if ($isDeviceServer) {
                # Servers: Default to Stale priority (will be upgraded to Critical if session actually failed)
                $IssueSeverity = "Stale"
            }
            elseif ($AccountType -eq 2) {
                # M365 Tenants: Default to Warning (Medium priority)
                $IssueSeverity = "Warning"
            }
            else {
                # Workstations: Default to Stale (Low priority)
                $IssueSeverity = "Stale"
            }
        }
        
        # EFFICIENT ERROR DETECTION: Check datasource error counts first to pre-filter which devices need session queries
        # Sum all datasource error counts from EnumerateAccountStatistics
        $totalDatasourceErrors = 0
        @('F7', 'S7', 'Z7', 'G7', 'J7', 'D5F6', 'D23F6', 'D4F12', 'D6F12', 'D8F12', 'D10F12', 'D12F12', 'D14F12', 'D15F12') | ForEach-Object {
            if ($DeviceResult.Settings.$_ -and (($DeviceResult.Settings.$_ -join '') -gt 0)) {
                $totalDatasourceErrors += [int]($DeviceResult.Settings.$_ -join '')
            }
        }
        
        # Set initial warning status if session shows warning (will be updated with error counts later)
        if ($LastSessionStatusCode -eq "3") {  # Warning
            $IssueSeverity = "Warning"
            $IssueDescription = if ($totalDatasourceErrors -gt 0) {
                "Backup completed with warnings ($totalDatasourceErrors errors)"
            } else {
                "Backup completed with warnings"
            }
        }
        
        # CRITICAL: Devices with datasource errors within the lookback period are NOT successful
        # Override Success classification if any datasource has errors
        if ($totalDatasourceErrors -gt 0 -and $IssueSeverity -eq "Success") {
            $IssueSeverity = "Warning"
            $IssueDescription = "Backup completed with $totalDatasourceErrors error(s)"
        }
        
        # Only query session errors if datasource error counts indicate errors OR session shows failed/error status
        $datasourceErrorData = $null
        if ($totalDatasourceErrors -gt 0 -or $LastSessionStatusCode -in @("2","3","4","6","8","9")) {
            Write-Verbose "Device $($DeviceResult.AccountId) has $totalDatasourceErrors datasource error(s) or failed session status ($LastSessionStatusCode) - querying session details..."
            # OPTIMIZATION v08: Pass I78 datasources to only query configured datasources
            $i78Value = $DeviceResult.Settings.I78 -join ''
            $datasourceErrorData = Get-LatestSessionError -DeviceId $DeviceResult.AccountId -Days $LookbackDays -AccountType $AccountType -DataSources $i78Value -DeviceNum $currentDeviceNum -TotalDevices $totalDevices
        } else {
            Write-Verbose "Device $($DeviceResult.AccountId) has no datasource errors and successful session status - skipping session query"
        }
        
        # If we found errors in recent sessions, update severity even if status was "Completed"
        if ($Script:DebugCDP) {
            Write-Host "  [DEBUG] datasourceErrorData result: " -ForegroundColor Magenta -NoNewline
            if ($datasourceErrorData) {
                Write-Host "FOUND (Errors: $($datasourceErrorData.Errors.Count), SessionTimes: $($datasourceErrorData.SessionTimes.Count))" -ForegroundColor Green
            } else {
                Write-Host "NULL" -ForegroundColor Yellow
            }
        }
        
        if ($datasourceErrorData) {
            $datasourceErrorMessages = $datasourceErrorData.Errors
            $datasourceSessionTimes = $datasourceErrorData.SessionTimes
            
            Write-Verbose "Found $($datasourceErrorMessages.Count) datasource(s) with recent errors for device $($DeviceResult.AccountId)"
            
            # If session status shows failed/error/CompletedWithErrors, use Critical
            if ($LastSessionStatusCode -in @("2","4","6","8","9")) {  # Failed/Error/CompletedWithErrors states
                $IssueSeverity = "Critical"
                
                # Select error from datasource with most recent session BackupStartTime
                $mostRecentError = $null
                $mostRecentSessionTime = [DateTime]::MinValue
                
                foreach ($datasourceName in $datasourceErrorMessages.Keys) {
                    $sessionTime = $datasourceSessionTimes[$datasourceName]
                    if ($sessionTime -gt $mostRecentSessionTime) {
                        $mostRecentSessionTime = $sessionTime
                        $mostRecentError = $datasourceErrorMessages[$datasourceName]
                    }
                }
                
                # Fallback to first error if no session time found
                if ($null -eq $mostRecentError) {
                    $mostRecentError = ($datasourceErrorMessages.Values | Select-Object -First 1)
                }
                
                # Get proper status text from the error's status code
                $errorStatusText = Get-SessionStatusText $mostRecentError.Status
                
                # Update LastSessionStatus to include session ID and timestamp
                $LastSessionStatus = "$($mostRecentError.SessionId) | $errorStatusText | $($mostRecentError.Timestamp)"
                # LastError contains only the error message
                $LastError = $mostRecentError.Message
                # Clean error message for M365 devices (removes Graph API JSON wrapper)
                $cleanedErrorMsg = if ($AccountType -eq 2) { Format-CleanErrorMessage $mostRecentError.Message } else { $mostRecentError.Message }
                $IssueDescription = "Backup failed - $($mostRecentError.Timestamp) - $cleanedErrorMsg"
            }
            # If status shows completed but we have errors, upgrade to Warning
            elseif ($IssueSeverity -eq "Success") {
                $IssueSeverity = "Warning"
                
                # Select error from datasource with most recent session BackupStartTime
                $mostRecentError = $null
                $mostRecentSessionTime = [DateTime]::MinValue
                
                foreach ($datasourceName in $datasourceErrorMessages.Keys) {
                    $sessionTime = $datasourceSessionTimes[$datasourceName]
                    if ($sessionTime -gt $mostRecentSessionTime) {
                        $mostRecentSessionTime = $sessionTime
                        $mostRecentError = $datasourceErrorMessages[$datasourceName]
                    }
                }
                
                # Fallback to first error if no session time found
                if ($null -eq $mostRecentError) {
                    $mostRecentError = ($datasourceErrorMessages.Values | Select-Object -First 1)
                }
                
                # Get proper status text from the error's status code
                $errorStatusText = Get-SessionStatusText $mostRecentError.Status
                
                # Update LastSessionStatus to include session ID and timestamp
                $LastSessionStatus = "$($mostRecentError.SessionId) | $errorStatusText | $($mostRecentError.Timestamp)"
                # LastError contains only the error message
                $LastError = $mostRecentError.Message
                # Clean error message for M365 devices (removes Graph API JSON wrapper)
                $cleanedErrorMsg = if ($AccountType -eq 2) { Format-CleanErrorMessage $mostRecentError.Message } else { $mostRecentError.Message }
                $IssueDescription = "Backup completed with errors - $($mostRecentError.Timestamp) - $cleanedErrorMsg"
            }
        }
        # Fallback if no detailed errors but session shows failure/CompletedWithErrors
        elseif ($LastSessionStatusCode -in @("2","4","6","8","9")) {  # Failed/Error/CompletedWithErrors states
            $IssueSeverity = "Critical"
            $LastError = if ($DeviceResult.Settings.ER -join '') { 
                $DeviceResult.Settings.ER -join '' 
            } else { 
                "Backup failed (no error details available)" 
            }
            $IssueDescription = "Backup failed - $LastError"
        }
        
        # Check for datasource-specific errors (from error count columns)
        $totalErrors = 0
        $errorDatasources = @()
        
        # Files and Folders (F7)
        if ($DeviceResult.Settings.F7 -and (($DeviceResult.Settings.F7 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.F7 -join '')
            $errorDatasources += "Files ($($DeviceResult.Settings.F7 -join '') errors)"
            $hasDataSourceErrors = $true
        }
        
        # System State (S7)
        if ($DeviceResult.Settings.S7 -and (($DeviceResult.Settings.S7 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.S7 -join '')
            $errorDatasources += "SystemState ($($DeviceResult.Settings.S7 -join '') errors)"
            $hasDataSourceErrors = $true
        }
        
        # MS SQL (Z7)
        if ($DeviceResult.Settings.Z7 -and (($DeviceResult.Settings.Z7 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.Z7 -join '')
            $errorDatasources += "SQL ($($DeviceResult.Settings.Z7 -join '') errors)"
            $hasDataSourceErrors = $true
        }
        
        # M365 Exchange (G7)
        if ($DeviceResult.Settings.G7 -and (($DeviceResult.Settings.G7 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.G7 -join '')
            $errorDatasources += "Exchange ($($DeviceResult.Settings.G7 -join '') errors)"
        }
        
        # M365 OneDrive (J7)
        if ($DeviceResult.Settings.J7 -and (($DeviceResult.Settings.J7 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.J7 -join '')
            $errorDatasources += "OneDrive ($($DeviceResult.Settings.J7 -join '') errors)"
        }
        
        # M365 SharePoint (D5F6)
        if ($DeviceResult.Settings.D5F6 -and (($DeviceResult.Settings.D5F6 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.D5F6 -join '')
            $errorDatasources += "SharePoint ($($DeviceResult.Settings.D5F6 -join '') errors)"
        }
        
        # M365 Teams (D23F6)
        if ($DeviceResult.Settings.D23F6 -and (($DeviceResult.Settings.D23F6 -join '') -gt 0)) {
            $totalErrors += [int]($DeviceResult.Settings.D23F6 -join '')
            $errorDatasources += "Teams ($($DeviceResult.Settings.D23F6 -join '') errors)"
        }
        
        # If we have errors but status isn't already Critical/Warning, set to Warning
        if ($totalErrors -gt 0 -and $IssueSeverity -eq "Success") {
            $IssueSeverity = "Warning"
            $IssueDescription = "Backup completed with $totalErrors error(s) in: $($errorDatasources -join ', ')"
        }
        
        # TEST MODE: Create fake failures for first 2 SUCCESS devices, then on next run convert first 2 ISSUE devices to success
        if ($TestMode) {
            # Check if we should create failures or close tickets
            $testTicketFile = Join-Path $ExportPath "TestModeState.txt"
            
            if (Test-Path $testTicketFile) {
                # Second run: Close tickets by converting issues to success
                if ($Script:DeviceIndex -lt 2 -and $IssueSeverity -ne "Success") {
                    Write-Host "  [TEST MODE] Converting issue to success for ticket close test: $($DeviceResult.Settings.AN -join '')" -ForegroundColor Magenta
                    $IssueSeverity = "Success"
                    $IssueDescription = "Backup successful (simulated for ticket close)"
                    $HoursSinceSuccess = 1.5  # Recent success
                }
                $Script:DeviceIndex++
            } else {
                # First run: Create failures for first 2 successful devices
                if ($Script:DeviceIndex -lt 2 -and $IssueSeverity -eq "Success") {
                    Write-Host "  [TEST MODE] Creating fake failure for ticket creation test: $($DeviceResult.Settings.AN -join '')" -ForegroundColor Magenta
                    $IssueSeverity = "Critical"
                    $IssueDescription = "TEST MODE: Simulated backup failure for ticket testing"
                    $LastSessionStatus = "Failed"
                    $LastError = "Simulated error for testing"
                    $HoursSinceSuccess = 72.5  # Simulate 3 days without success
                    
                    # Create state file after processing first 2 devices
                    if ($Script:DeviceIndex -eq 1) {
                        "TestMode: Tickets created" | Set-Content $testTicketFile
                        Write-Host "  [TEST MODE] State file created. Run script again to test ticket closure." -ForegroundColor Magenta
                    }
                }
                $Script:DeviceIndex++
            }
        }
        
        # Get datasource-specific status
        $datasourceStatusString = Get-DatasourceStatus -DeviceSettings $DeviceResult.Settings -DataSourceString ($DeviceResult.Settings.I78 -join '')
        
        # Debug: Show timezone offset value from API
        $tzValue = $DeviceResult.Settings.TZ -join ''
        if ($debugCDP) { Write-Host "  Device: $($DeviceResult.Settings.AN -join '') | TZ from API: '$tzValue'" -ForegroundColor Cyan }
        
        # OPTIMIZATION v06: Defer partner hierarchy lookup until needed (only for devices with issues)
        # Store PartnerID and PartnerName for later resolution
        # This eliminates 87% of unnecessary API calls (750 devices → 96 with issues)
        
        # Calculate additional display properties for ticket templates
        # Get worst datasource time (for "Oldest Problem" field)
        $worstDatasourceTimeFormatted = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            if ($UseLocalTime) {
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($worstDatasourceTime, [System.TimeZoneInfo]::Local)
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else {
                "$($worstDatasourceTime.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            }
        } else { 'N/A' }
        
        # Get best datasource time (most recent success)
        $bestDatasourceTimeFormatted = if ($bestDatasourceTime -and $bestDatasourceTime -ne 'N/A' -and $bestDatasourceTime -is [DateTime]) {
            if ($UseLocalTime) {
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($bestDatasourceTime, [System.TimeZoneInfo]::Local)
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else {
                "$($bestDatasourceTime.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            }
        } else { 'N/A' }
        
        # Get most recent session time (device level TL) - same as TimeStamp
        $mostRecentSessionTime = if ($UseLocalTime) {
            $utcTime = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '')
            if ($utcTime -and $utcTime -ne 'N/A') {
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcTime, [System.TimeZoneInfo]::Local)
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else { 'N/A' }
        } else {
            $utcTime = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '')
            if ($utcTime -and $utcTime -ne 'N/A') {
                "$($utcTime.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            } else { 'N/A' }
        }
        
        # Get status names for display
        $lastSessionStatusName = Get-SessionStatusText -StatusCode $LastSessionStatusCode
        $lastSuccessStatusCode = $DeviceResult.Settings.TQ -join ''  # TQ = Last success status
        $lastSuccessStatusName = Get-SessionStatusText -StatusCode $lastSuccessStatusCode
        
        # Get worst and best datasource status names
        $worstDatasourceStatusName = if ($worstDatasourceStatus) { Get-SessionStatusText -StatusCode $worstDatasourceStatus } else { 'Unknown' }
        $bestDatasourceStatusName = if ($bestDatasourceStatus) { Get-SessionStatusText -StatusCode $bestDatasourceStatus } else { 'Unknown' }
        
        # MostRecent fields = actual most recent session from device (not overwritten with best datasource)
        # NOTE: $mostRecentSessionTime was already calculated above (lines 3626-3645) from TS column
        # Do NOT overwrite it with $bestDatasourceTimeFormatted - they serve different purposes:
        #   - MostRecentSessionTime = Last backup ATTEMPT (success or failure) - from TS/TimeStamp
        #   - BestDatasourceTime = Last SUCCESSFUL backup - from datasource analysis
        # Keep the mostRecentSessionTime value that was already calculated from TS
        $mostRecentStatusName = $lastSessionStatusName  # Use device-level status, not best datasource status
        $mostRecentDatasource = if ($bestDatasourceName -ne 'Unknown') { $bestDatasourceName } elseif ($worstDatasourceName -ne 'Unknown') { $worstDatasourceName } else { 'Unknown' }
        
        # Calculate time ago for best datasource
        $bestDatasourceTimeAgo = if ($bestDatasourceTime -and $bestDatasourceTime -ne 'N/A' -and $bestDatasourceTime -is [DateTime]) {
            $dsHoursSinceSuccess = ((Get-Date).ToUniversalTime() - $bestDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $dsHoursSinceSuccess
        } else { '' }
        
        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ 
            AccountID      = [Int]$DeviceResult.AccountId
            PartnerID      = [string]$DeviceResult.PartnerId
            AccountType    = $AccountType  # 1=Servers AND Workstations, 2=M365
            DeviceName     = $DeviceResult.Settings.AN -join ''
            ComputerName   = $DeviceResult.Settings.MN -join ''
            DeviceAlias    = $DeviceResult.Settings.AL -join ''
            PartnerName    = $DeviceResult.Settings.AR -join ''
            EndCustomer    = $partnerHierarchy.EndCustomer
            EndCustomerLevel = $partnerHierarchy.EndCustomerLevel
            Site           = $partnerHierarchy.Site
            Reference      = $DeviceResult.Settings.PF -join ''
            PartnerIdForHierarchy = $partnerHierarchy.EndCustomerPartnerId
            Creation       = $($creationDate = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join ''); $creationDate)
            TimeStamp      = if ($UseLocalTime) {
                $utcTime = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '')
                if ($utcTime -and $utcTime -ne 'N/A') {
                    $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcTime, [System.TimeZoneInfo]::Local)
                    $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                        [System.TimeZoneInfo]::Local.DaylightName
                    } else {
                        [System.TimeZoneInfo]::Local.StandardName
                    }
                    $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                    "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
                } else { $utcTime }
            } else {
                $utcTime = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '')
                if ($utcTime -and $utcTime -ne 'N/A') {
                    "$($utcTime.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
                } else { $utcTime }
            }
            LastSuccess    = if ($UseLocalTime) {
                if ($LastSuccess -and $LastSuccess -ne 'N/A') {
                    $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($LastSuccess, [System.TimeZoneInfo]::Local)
                    $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localTime)) {
                        [System.TimeZoneInfo]::Local.DaylightName
                    } else {
                        [System.TimeZoneInfo]::Local.StandardName
                    }
                    $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                    "$($localTime.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
                } else { $LastSuccess }
            } else {
                if ($LastSuccess -and $LastSuccess -ne 'N/A') {
                    "$($LastSuccess.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
                } else { $LastSuccess }
            }
            LastSuccessDeviceLevel = $LastSuccess  # Device-level TL timestamp
            HoursSinceSuccess = [Math]::Round($HoursSinceSuccess,1)
            HoursSinceSuccessDeviceLevel = if ($LastSuccess -and $LastSuccess -is [DateTime]) { 
                [Math]::Round(((Get-Date).ToUniversalTime() - $LastSuccess).TotalHours, 1) 
            } elseif ($creationDate -and $creationDate -is [DateTime]) {
                # No prior success - calculate relative to Creation Date instead of Unix epoch
                [Math]::Round(((Get-Date).ToUniversalTime() - $creationDate).TotalHours, 1)
            } else { 
                $null 
            }
            WorstDatasourceName = if ($worstDatasourceName) { $worstDatasourceName } else { "Unknown" }
            WorstDatasourceTime = $worstDatasourceTimeFormatted
            WorstDatasourceStatusName = $worstDatasourceStatusName
            BestDatasourceName = if ($bestDatasourceName) { $bestDatasourceName } else { "Unknown" }
            BestDatasourceTime = $bestDatasourceTimeFormatted
            BestDatasourceTimeAgo = $bestDatasourceTimeAgo
            BestDatasourceStatusName = $bestDatasourceStatusName
            MostRecentSessionTime = $mostRecentSessionTime
            MostRecentDatasource = $mostRecentDatasource
            MostRecentStatusName = $mostRecentStatusName
            LastSessionStatus = $LastSessionStatus
            LastSessionStatusName = $lastSessionStatusName
            LastSuccessStatusName = $lastSuccessStatusName
            DatasourceStatus = $datasourceStatusString
            LastError      = $LastError
            IsInProcess    = $anyInProcess  # True if any datasource is currently running (status code "1")
            IssueSeverity  = $IssueSeverity
            IssueDescription = $IssueDescription
            SelectedSize   = Format-DataSize $DeviceResult.Settings.T3
            UsedStorage    = Format-DataSize $DeviceResult.Settings.US
            DataSources    = $DeviceResult.Settings.I78 -join ''
            Account        = $DeviceResult.Settings.AU -join ''
            Location       = $DeviceResult.Settings.LN -join ''
            Notes          = $DeviceResult.Settings.AA843 -join ''
            Product        = $DeviceResult.Settings.PN -join ''
            ProductID      = $DeviceResult.Settings.PD -join ''
            Profile        = $DeviceResult.Settings.OP -join ''
            OS             = $DeviceResult.Settings.OS -join ''
            OSType         = $DeviceResult.Settings.OT -join ''
            Physicality    = $DeviceResult.Settings.I81 -join ''
            CPUCores       = $DeviceResult.Settings.I84 -join ''
            RAMSizeGB      = if ($DeviceResult.Settings.I85 -join '') { [math]::Round(([int64]($DeviceResult.Settings.I85 -join '')) / 1GB, 2) } else { 'N/A' }
            Manufacturer   = $DeviceResult.Settings.MF -join ''
            Model          = $DeviceResult.Settings.MO -join ''
            ProfileID      = $DeviceResult.Settings.OI -join ''
            TimezoneOffset = Format-TimezoneOffset ($DeviceResult.Settings.TZ -join '')
            IPAddress      = $DeviceResult.Settings.IP -join ''
            ExternalIP     = $DeviceResult.Settings.EI -join ''
            DeviceSettings = $DeviceResult.Settings  # Store full settings for datasource details
            ErrorMessages  = $datasourceErrorMessages  # Store detailed error messages for datasource display
        }
    }

    # Section summary
    $SectionDeviceCount = $Script:DeviceDetail.Count - $StartingCount
    $SectionIssues = ($Script:DeviceDetail | Select-Object -Skip $StartingCount | Where-Object { $_.IssueSeverity -ne "Success" }).Count
    $deviceTypeLabel = if ($AccountType -eq 1) { "Servers and Workstations" } else { "M365 Tenants" }
    Write-Host "  $deviceTypeLabel`: $SectionDeviceCount devices processed, $SectionIssues with issues" -ForegroundColor $(if ($SectionIssues -gt 0) { 'Yellow' } else { 'Green' })

}  ## EnumerateAccountStatistics API Call

#endregion ----- Backup.Management JSON Calls ----

#region ----- ConnectWise Manage Functions ----

Function InstallCWMPSModule {
    if (Get-Module -ListAvailable -Name "ConnectWiseManageAPI") {
        if ($debugCWM) {Write-Host "  ConnectWise Manage PowerShell Module Already Installed" -ForegroundColor Green}
    } else {
        Write-Host "  Installing ConnectWise Manage PowerShell Module" -ForegroundColor Yellow
        Install-Module -Name ConnectWiseManageAPI -Force -AllowClobber
        Write-Host "  ConnectWise Manage PowerShell Module Installed" -ForegroundColor Green
    }

    if (Get-Module -ListAvailable -Name "ConnectWiseManageAPI") {
        Import-Module ConnectWiseManageAPI
        if ($debugCWM) {Write-Host "  ConnectWise Manage PowerShell Module Imported" -ForegroundColor Green}
    } else {
        Write-Error "ConnectWise Manage PowerShell Module Installation Failed"
        Break
    }
}  ## Install and Import the ConnectWise Manage PowerShell module

Function Prompt-CWMAPICreds {
    $CWMAPICreds =$null
    $CWMAPICreds = @{}

    $CWMAPICreds.Server = Read-Host "Enter the DNS Name of your ConnectWise Manage server (e.g., staging.connectwisedev.com)"
    $CWMAPICreds.Company = Read-Host "Enter the Company you use when logging on to ConnectWise Manage"
    $CWMAPICreds.pubKey = Read-Host "Enter the Public key created for this integration"
    $CWMAPICreds.privateKey = Read-Host "Enter the Private key created for this integration"
    $CWMAPICreds.clientId = Read-Host "Enter your ClientID (You can create/retrieve your ClientID at https://developer.connectwise.com/ClientID)"

    # Encrypt sensitive fields before exporting
    $CWMAPICreds.privateKey = ConvertTo-SecureString -String $CWMAPICreds.privateKey -AsPlainText -Force | ConvertFrom-SecureString
    $CWMAPICreds.pubKey = ConvertTo-SecureString -String $CWMAPICreds.pubKey -AsPlainText -Force | ConvertFrom-SecureString
    $CWMAPICreds.clientId = ConvertTo-SecureString -String $CWMAPICreds.clientId -AsPlainText -Force | ConvertFrom-SecureString

    $CWMAPICreds | Export-Clixml -Path $CWMAPICredsFile -Force
}

Function Get-CWMAPICreds {
    if (Test-Path $CWMAPICredsFile) {
        $CWMAPICreds = Import-Clixml -Path $CWMAPICredsFile

        # Decrypt sensitive fields after importing
        $securePrivateKey = $CWMAPICreds.privateKey | ConvertTo-SecureString -Force
        $CWMAPICreds.privateKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePrivateKey))

        $securePubKey = $CWMAPICreds.pubKey | ConvertTo-SecureString -Force
        $CWMAPICreds.pubKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePubKey))
     
        $secureClientId = $CWMAPICreds.clientId | ConvertTo-SecureString -Force
        $CWMAPICreds.clientId = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureClientId))
   
    } else {
        Prompt-CWMAPICreds
        $CWMAPICreds = Get-CWMAPICreds
    }
    return $CWMAPICreds
}

Function AuthenticateCWM {
    # CRITICAL: Set bypass BEFORE Connect-CWM for staging environments
    if ($AllowInsecureSSL) {
        # Bypass SSL certificate validation
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        
        # For PowerShell 6+ (Core)
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck'] = $true
            $PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck'] = $true
        }
        
        Write-Host "  [WARNING] SSL certificate validation bypassed for staging environment" -ForegroundColor Yellow
    }
    
    $CWMAPICreds = Get-CWMAPICreds
    if ($null -eq $CWMAPICreds) {
        Prompt-CWMAPICreds      
        $CWMAPICreds = Get-CWMAPICreds
    }
    
    # CRITICAL v09: Disconnect existing session to clear module cache (prevents stale data in VSCode runs)
    # Only disconnect if we're already connected (prevents "Connection object not found" warning)
    try {
        $null = Get-CWMServiceBoard -pageSize 1 -ErrorAction Stop
        # If we got here, we're connected - disconnect to clear cache
        Disconnect-CWM -ErrorAction SilentlyContinue
    } catch {
        # Not connected yet - ignore (this is expected on first run)
    }
    
    # Connect to ConnectWise Manage
    Write-Host "`n  Connecting to ConnectWise Manage..." -ForegroundColor Cyan
    $null = Connect-CWM @CWMAPICreds
    
    # Verify connection was successful before validating parameters
    Write-Host "`n  Verifying ConnectWise connection..." -ForegroundColor Cyan
    try {
        $testConnection = Get-CWMServiceBoard -pageSize 1 -ErrorAction Stop
        Write-Host "  Connected to ConnectWise Manage" -ForegroundColor Green
    } catch {
        Write-Host "  Failed to connect to ConnectWise Manage" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "`n  Cannot validate ticket parameters without a valid connection." -ForegroundColor Yellow
        Write-Host "  Please check your credentials and try again." -ForegroundColor Yellow
        exit 1
    }
    
    if ($debugCWM) {
        Write-Output "`nDebugging ConnectWise Manage API Credentials:`n"
        Write-Output "Server: $($CWMAPICreds.Server)"
        Write-Output "Company: $($CWMAPICreds.Company)"
        Write-Output "Public Key: $($CWMAPICreds.pubKey)"
        Write-Output "Private Key: $($CWMAPICreds.privateKey)"
        Write-Output "Client ID: $($CWMAPICreds.clientId)"
    }
    
    # Validate ticket parameters against actual ConnectWise configuration
    Write-Host "`n  Validating ConnectWise ticket parameters..." -ForegroundColor Cyan
    $invalidParams = @()
    
    # Check if Board exists
    try {
        $boards = Get-CWMServiceBoard -all
        $boardExists = $boards | Where-Object { $_.name -eq $TicketBoard }
        if (-not $boardExists) {
            $invalidParams += @{
                Parameter = 'TicketBoard'
                Value = $TicketBoard
                Type = 'Board'
            }
        } else {
            Write-Host "  Board '$TicketBoard' found" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Could not validate Service Board"
    }
    
    # Check if Priorities exist (all three types)
    try {
        $priorities = Get-CWMPriority -all
        $priorityServerExists = $priorities | Where-Object { $_.name -eq $TicketPriorityServer }
        $priorityWorkstationExists = $priorities | Where-Object { $_.name -eq $TicketPriorityWorkstation }
        $priorityM365Exists = $priorities | Where-Object { $_.name -eq $TicketPriorityM365 }
        
        if (-not $priorityServerExists) {
            $invalidParams += @{
                Parameter = 'TicketPriorityServer'
                Value = $TicketPriorityServer
                Type = 'Priority'
            }
        } else {
            Write-Host "  Priority 'Server: $TicketPriorityServer' found" -ForegroundColor Green
        }
        
        if (-not $priorityWorkstationExists) {
            $invalidParams += @{
                Parameter = 'TicketPriorityWorkstation'
                Value = $TicketPriorityWorkstation
                Type = 'Priority'
            }
        } else {
            Write-Host "  Priority 'Workstation: $TicketPriorityWorkstation' found" -ForegroundColor Green
        }
        
        if (-not $priorityM365Exists) {
            $invalidParams += @{
                Parameter = 'TicketPriorityM365'
                Value = $TicketPriorityM365
                Type = 'Priority'
            }
        } else {
            Write-Host "  Priority 'M365: $TicketPriorityM365' found" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Could not validate Priorities"
    }
    
    # Check if Status exists (if board was found)
    if ($boardExists) {
        try {
            $statuses = Get-CWMBoardStatus -parentId $boardExists.id -all
            $newStatusExists = $statuses | Where-Object { $_.name -eq $TicketStatus }
            $closedStatusExists = $statuses | Where-Object { $_.name -eq $TicketClosedStatus }
            
            if (-not $newStatusExists) {
                $invalidParams += @{
                    Parameter = 'TicketStatus'
                    Value = $TicketStatus
                    Type = 'Status'
                }
            } else {
                Write-Host "  New Status '$TicketStatus' found" -ForegroundColor Green
            }
            
            if (-not $closedStatusExists) {
                $invalidParams += @{
                    Parameter = 'TicketClosedStatus'
                    Value = $TicketClosedStatus
                    Type = 'Status'
                }
            } else {
                Write-Host "  Closed Status '$TicketClosedStatus' found" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Could not validate Statuses - unable to query ConnectWise API"
            
            # Treat validation failure as invalid parameters - offer to run options script
            $invalidParams += @{
                Parameter = 'TicketStatus & TicketClosedStatus'
                Value = 'Validation Failed'
                Type = 'Status'
            }
        }
    }
    
    # If any parameters are invalid, offer to run the options script
    if ($invalidParams.Count -gt 0) {
        Write-Host ""
        Write-Host "  WARNING: The following ticket parameters do not exist in ConnectWise:" -ForegroundColor Yellow
        foreach ($param in $invalidParams) {
            Write-Host "           -$($param.Parameter) = '$($param.Value)' ($($param.Type) not found)" -ForegroundColor Red
        }
        Write-Host ""
        
        $optionsScriptPath = Join-Path $PSScriptRoot "Cove2CWM-SetTicketsConfig.v10.ps1"
        
        if (-not (Test-Path $optionsScriptPath)) {
            # Try to find any version of the script
            $allVersions = Get-ChildItem -Path $PSScriptRoot -Filter "Cove2CWM-SetTicketsConfig*.ps1" -File | Sort-Object Name -Descending
            
            if ($allVersions.Count -eq 1) {
                $optionsScriptPath = $allVersions[0].FullName
            }
            elseif ($allVersions.Count -gt 1) {
                $optionsScriptPath = $allVersions[0].FullName
            }
            else {
                $optionsScriptPath = $null
            }
        }
        
        if ($optionsScriptPath -and (Test-Path $optionsScriptPath)) {
            $scriptName = Split-Path $optionsScriptPath -Leaf
            Write-Host "  Would you like to run $scriptName to select valid options?" -ForegroundColor Cyan
            $response = Read-Host "  Enter Y to update parameters, or N to exit"
            
            if ($response -eq 'Y' -or $response -eq 'y') {
                Write-Host ""
                Write-Host "  Launching Cove2CWM-SetTicketsConfig.ps1..." -ForegroundColor Cyan
                Write-Host ""
                
                $monitoringScriptPath = $PSCommandPath
                & $optionsScriptPath -MonitoringScriptPath $monitoringScriptPath
                
                Write-Host ""
                Write-Host "  Checking if parameters were updated..." -ForegroundColor Cyan
                
                # Re-read the script to check if parameters were updated
                $updatedContent = Get-Content $monitoringScriptPath -Raw
                $boardMatch = if ($updatedContent -match '\$TicketBoard\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                $statusMatch = if ($updatedContent -match '\$TicketStatus\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                $priorityServerMatch = if ($updatedContent -match '\$TicketPriorityServer\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                $priorityWorkstationMatch = if ($updatedContent -match '\$TicketPriorityWorkstation\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                $priorityM365Match = if ($updatedContent -match '\$TicketPriorityM365\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                $closedMatch = if ($updatedContent -match '\$TicketClosedStatus\s*=\s*"([^"]*)"') { $matches[1] } else { $null }
                
                if ($boardMatch -and $statusMatch -and $priorityServerMatch -and $priorityWorkstationMatch -and $priorityM365Match -and $closedMatch) {
                    Write-Host "  Parameters were successfully updated!" -ForegroundColor Green
                    Write-Host ""
                    Write-Host "  Updated values:" -ForegroundColor Cyan
                    Write-Host "    -TicketBoard               : $boardMatch" -ForegroundColor White
                    Write-Host "    -TicketStatus              : $statusMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityServer      : $priorityServerMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityWorkstation : $priorityWorkstationMatch" -ForegroundColor White
                    Write-Host "    -TicketPriorityM365        : $priorityM365Match" -ForegroundColor White
                    Write-Host "    -TicketClosedStatus        : $closedMatch" -ForegroundColor White
                    Write-Host ""
                    Write-Host "  Updating runtime parameters..." -ForegroundColor Green
                    
                    # Force update the parameter variables at script scope (overrides command-line parameters)
                    Set-Variable -Name 'TicketBoard' -Value $boardMatch -Scope Script -Force
                    Set-Variable -Name 'TicketStatus' -Value $statusMatch -Scope Script -Force
                    Set-Variable -Name 'TicketPriorityServer' -Value $priorityServerMatch -Scope Script -Force
                    Set-Variable -Name 'TicketPriorityWorkstation' -Value $priorityWorkstationMatch -Scope Script -Force
                    Set-Variable -Name 'TicketPriorityM365' -Value $priorityM365Match -Scope Script -Force
                    Set-Variable -Name 'TicketClosedStatus' -Value $closedMatch -Scope Script -Force
                    
                    # Also update local scope for immediate use
                    $TicketBoard = $boardMatch
                    $TicketStatus = $statusMatch
                    $TicketPriorityServer = $priorityServerMatch
                    $TicketPriorityWorkstation = $priorityWorkstationMatch
                    $TicketPriorityM365 = $priorityM365Match
                    $TicketClosedStatus = $closedMatch
                    
                    Write-Host "  Runtime parameters updated successfully!" -ForegroundColor Green
                    Write-Host ""
                    
                    # Re-validate the updated parameters are now correct
                    Write-Host "  Re-validating updated parameters..." -ForegroundColor Cyan
                    $revalidationFailed = $false
                    
                    try {
                        $boards = Get-CWMServiceBoard -all
                        $boardExists = $boards | Where-Object { $_.name -eq $TicketBoard }
                        if ($boardExists) {
                            Write-Host "  Board '$TicketBoard' validated" -ForegroundColor Green
                            
                            $statuses = Get-CWMBoardStatus -parentId $boardExists.id -all
                            $newStatusExists = $statuses | Where-Object { $_.name -eq $TicketStatus }
                            $closedStatusExists = $statuses | Where-Object { $_.name -eq $TicketClosedStatus }
                            
                            if ($newStatusExists) {
                                Write-Host "  New Status '$TicketStatus' validated" -ForegroundColor Green
                            } else {
                                Write-Host "  New Status '$TicketStatus' NOT FOUND" -ForegroundColor Red
                                $revalidationFailed = $true
                            }
                            
                            if ($closedStatusExists) {
                                Write-Host "  Closed Status '$TicketClosedStatus' validated" -ForegroundColor Green
                            } else {
                                Write-Host "  Closed Status '$TicketClosedStatus' NOT FOUND" -ForegroundColor Red
                                $revalidationFailed = $true
                            }
                        } else {
                            Write-Host "  Board '$TicketBoard' NOT FOUND" -ForegroundColor Red
                            $revalidationFailed = $true
                        }
                        
                        $priorities = Get-CWMPriority -all
                        
                        # Validate Server Priority
                        $priorityServerExists = $priorities | Where-Object { $_.name -eq $TicketPriorityServer }
                        if ($priorityServerExists) {
                            Write-Host "  Priority 'Server: $TicketPriorityServer' validated" -ForegroundColor Green
                        } else {
                            Write-Host "  Priority 'Server: $TicketPriorityServer' NOT FOUND" -ForegroundColor Red
                            $revalidationFailed = $true
                        }
                        
                        # Validate Workstation Priority
                        $priorityWorkstationExists = $priorities | Where-Object { $_.name -eq $TicketPriorityWorkstation }
                        if ($priorityWorkstationExists) {
                            Write-Host "  Priority 'Workstation: $TicketPriorityWorkstation' validated" -ForegroundColor Green
                        } else {
                            Write-Host "  Priority 'Workstation: $TicketPriorityWorkstation' NOT FOUND" -ForegroundColor Red
                            $revalidationFailed = $true
                        }
                        
                        # Validate M365 Priority
                        $priorityM365Exists = $priorities | Where-Object { $_.name -eq $TicketPriorityM365 }
                        if ($priorityM365Exists) {
                            Write-Host "  Priority 'M365: $TicketPriorityM365' validated" -ForegroundColor Green
                        } else {
                            Write-Host "  Priority 'M365: $TicketPriorityM365' NOT FOUND" -ForegroundColor Red
                            $revalidationFailed = $true
                        }
                    } catch {
                        Write-Warning "Re-validation failed: $($_.Exception.Message)"
                        $revalidationFailed = $true
                    }
                    
                    if ($revalidationFailed) {
                        Write-Host ""
                        Write-Host "  ERROR: Updated parameters failed validation!" -ForegroundColor Red
                        Write-Host "  The script file was updated, but the values don't exist in ConnectWise." -ForegroundColor Yellow
                        Write-Host "  Please run Cove2CWM-SetTicketsConfig.ps1 again to select valid options." -ForegroundColor Yellow
                        exit 1
                    }
                    
                    Write-Host ""
                    Write-Host "  All parameters validated successfully! Continuing execution..." -ForegroundColor Green
                    Write-Host ""
                } else {
                    Write-Host "  Parameters were not updated. Exiting..." -ForegroundColor Yellow
                    exit 1
                }
            } else {
                Write-Host "  Script execution cancelled." -ForegroundColor Yellow
                exit 1
            }
        }
        
        if (-not $optionsScriptPath -or -not (Test-Path $optionsScriptPath)) {
            Write-Host "  Cove2CWM-SetTicketsConfig.ps1 not found. Please update parameters manually." -ForegroundColor Yellow
            Write-Host "  Script execution cancelled." -ForegroundColor Yellow
            exit 1
        }
    } else {
        Write-Host "  All ticket parameters validated successfully!" -ForegroundColor Green
    }
    Write-Host ""
}  ## Connect to ConnectWise Manage instance

Function Get-CWMTicketURL {
    param(
        [int]$TicketID
    )
    
    $CWMAPICreds = Get-CWMAPICreds
    $baseUrl = "https://$($CWMAPICreds.Server)"
    $ticketUrl = "$baseUrl/v4_6_release/services/system_io/Service/fv_sr100_request.rails?service_recid=$TicketID&companyName=$($CWMAPICreds.Company)"
    
    return $ticketUrl
}  ## Build direct URL to ConnectWise ticket in web UI

Function Get-CWMCompanyForDevice {
    param(
        [PSObject]$Device
    )
    
    $CWMcompany = $null
    
    # Determine which level to match at:
    # - If UseDevicePartner: Match at device's immediate partner level
    # - Otherwise: Match at End Customer level (parent company)
    if ($UseDevicePartner) {
        $matchName = $Device.PartnerName
        $matchLevel = "Device Partner"
    } else {
        $matchName = if ($Device.EndCustomer) { $Device.EndCustomer } else { $Device.PartnerName }
        $matchLevel = "End Customer"
    }
    
    if ($debugCWM) {
        Write-Host "[DEBUG] Get-CWMCompanyForDevice: DeviceName='$($Device.DeviceName)', Partner='$($Device.PartnerName)', EndCustomer='$($Device.EndCustomer)', Reference='$($Device.Reference)'" -ForegroundColor Magenta
        Write-Host "[DEBUG]   UseDevicePartner=$UseDevicePartner, Matching at: $matchLevel, matchName='$matchName'" -ForegroundColor Cyan
    }
    
    # OPTIMIZATION v10: Check EndCustomer cache first (before any CWM API calls)
    $Script:EndCustomerCacheStats.TotalLookups++
    
    if ($Device.EndCustomer -and $Script:EndCustomerToCWMCompanyCache.ContainsKey($Device.EndCustomer)) {
        $Script:EndCustomerCacheStats.CacheHits++
        $cachedCompany = $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer]
        
        # Track unique customer (update on every lookup, not just creation)
        if (-not $Script:EndCustomerCacheStats.UniqueCustomersCached.Contains($Device.EndCustomer)) {
            $Script:EndCustomerCacheStats.UniqueCustomersCached.Add($Device.EndCustomer) | Out-Null
        }
        
        # Return cached company object directly (no CWM API call needed)
        if ($debugCWM) { Write-Host "[DEBUG] Cache HIT: '$($Device.EndCustomer)' → [$($cachedCompany.ID)] $($cachedCompany.Name)" -ForegroundColor Green }
        return $cachedCompany
    } else {
        $Script:EndCustomerCacheStats.CacheMisses++
        if ($debugCWM) { Write-Host "[DEBUG] Cache MISS: '$($Device.EndCustomer)'" -ForegroundColor Yellow }
    }
    
    # Strategy 0: Check session cache first (prevents duplicate company creation)
    # Check for BOTH full name (pre-truncation) and truncated name (post-creation)
    $alreadyCreated = $Script:CompanyLookupCache | Where-Object { 
        $_.CompanyName -eq $matchName -or 
        $_.MatchName -eq $matchName -or
        ($_.TruncatedName -and $_.TruncatedName -eq $matchName)
    }
    if ($alreadyCreated) {
        # Check if this is a placeholder (company creation in progress)
        if ($alreadyCreated.IsPlaceholder) {
            Write-Host "[DEBUG] Strategy 0: Found PLACEHOLDER for '$matchName' - another device is creating this company, waiting..." -ForegroundColor Yellow
            
            # Wait for placeholder to be updated with actual company ID (max 30 seconds)
            $waitTime = 0
            while ($alreadyCreated.IsPlaceholder -and $waitTime -lt 30) {
                Start-Sleep -Seconds 2
                $waitTime += 2
                Write-Host "[DEBUG]   Waiting for company creation to complete... ($waitTime seconds)" -ForegroundColor Gray
            }
            
            if ($alreadyCreated.IsPlaceholder) {
                Write-Warning "Timeout waiting for company '$matchName' to be created - proceeding with lookup"
            } else {
                if ($debugCWM) { Write-Host "[DEBUG]   Placeholder resolved! Company created with ID: $($alreadyCreated.CompanyId)" -ForegroundColor Green }
            }
        }
        
        # If we have a real company ID now (not placeholder), return it
        if (-not $alreadyCreated.IsPlaceholder -and $alreadyCreated.CompanyId -gt 0) {
            if ($debugCWM) { Write-Host "[DEBUG] Strategy 0: Found '$matchName' in session cache [ID: $($alreadyCreated.CompanyId)]" -ForegroundColor Green }
            
            # Return cached company object directly (no CWM API call needed)
            if ($alreadyCreated.CompanyObject) {
                if ($debugCWM) { Write-Host "[DEBUG]   Session cache hit - returning cached company object: [$($alreadyCreated.CompanyObject.ID)] $($alreadyCreated.CompanyObject.Name) (no API call)" -ForegroundColor Green }
                return $alreadyCreated.CompanyObject
            } else {
                # Fallback: If CompanyObject property doesn't exist (old cache format), re-query
                if ($debugCWM) { Write-Host "[DEBUG]   ⚠ Session cache has ID but no object - re-querying CWM" -ForegroundColor Yellow }
                try {
                    $CWMcompanyResults = Get-CWMcompany -condition "id=$($alreadyCreated.CompanyId) and deletedFlag = false"
                    $Script:PerformanceMetrics.CompanyLookupCount++
                    if ($CWMcompanyResults) {
                        $CWMcompany = $CWMcompanyResults | Select-Object -First 1
                        # Update cache with full object for future lookups
                        $alreadyCreated.CompanyObject = $CWMcompany
                        if ($debugCWM) { Write-Host "[DEBUG]   Re-queried and updated session cache with company object" -ForegroundColor Green }
                        return $CWMcompany
                    }
                } catch {
                    if ($debugCWM) { Write-Host "[DEBUG]   ⚠ Company ID $($alreadyCreated.CompanyId) in cache but failed to re-query: $_" -ForegroundColor Yellow }
                    # Continue to other lookup strategies
                }
            }
        }
    } else {
        if ($debugCWM) { Write-Host "[DEBUG] Strategy 0: '$matchName' not in session cache - proceeding to CWM lookup" -ForegroundColor Gray }
    }
    
    # Strategy 1: Check ExternalCode (PF field) for CWM:ID reference
    if ($Device.Reference -match 'CWM:(\d+)') {
        $extractedCWMId = $matches[1]
        try {
            $CWMcompanyResults = Get-CWMcompany -condition "id=$extractedCWMId and deletedFlag = false"
            $Script:PerformanceMetrics.CompanyLookupCount++
            if ($CWMcompanyResults) {
                # Ensure single company (query may return array)
                $CWMcompany = $CWMcompanyResults | Select-Object -First 1
                
                # Check if company is deleted
                if ($CWMcompany.deletedFlag) {
                    if ($debugCWM) {
                        Write-Warning "  CWM:$extractedCWMId found but company is marked as DELETED - skipping"
                    }
                    $CWMcompany = $null
                } else {
                    if ($debugCWM) {
                        Write-Host "  Matched via ExternalCode CWM:$extractedCWMId - [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                    }
                    
                    # Add to session cache for future lookups within this run
                    $Script:CompanyLookupCache += [PSCustomObject]@{
                        CompanyId = [int]$CWMcompany.ID
                        CompanyName = $CWMcompany.Name
                        MatchName = $matchName
                        Identifier = $CWMcompany.identifier
                        EndCustomer = $Device.EndCustomer
                        IsPlaceholder = $false
                        CompanyObject = $CWMcompany
                    }
                    
                    # OPTIMIZATION v10: Add to EndCustomer cache (store full company object)
                    if ($Device.EndCustomer) {
                        $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer] = $CWMcompany
                        if ($debugCWM) { Write-Host "[DEBUG]   [CACHE] Added EndCustomer '$($Device.EndCustomer)' → CWM Company [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                    }
                    return $CWMcompany
                }
            }
        } catch {
            if ($debugCWM) {
                Write-Warning "  CWM:$extractedCWMId found in ExternalCode but company not found in ConnectWise"
            }
        }
    }
    
    # Strategy 2: Try to extract CWM Company ID from partner name (format: "Name | CWMCompanyName ~ CWMCompanyID")
    $CWMCompanyID = (($matchName -split '\| ') -split ' ~')[2]
    if ($CWMCompanyID) {
        try {
            $CWMcompanyResults = Get-CWMcompany -condition "id=$CWMCompanyID and deletedFlag = false"
            $Script:PerformanceMetrics.CompanyLookupCount++
            if ($CWMcompanyResults) {
                # Ensure single company (query may return array)
                $CWMcompany = $CWMcompanyResults | Select-Object -First 1
                
                # Check if company is deleted
                if ($CWMcompany.deletedFlag) {
                    if ($debugCWM) {
                        Write-Warning "  Company ID $CWMCompanyID found but marked as DELETED - skipping"
                    }
                    $CWMcompany = $null
                } else {
                    if ($debugCWM) {
                        Write-Host "  Matched via Partner Name Format: [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                    }
                    # Update Cove ExternalCode with CWM ID link
                    if ($Device.PartnerIdForHierarchy) {
                        Update-CovePartnerExternalCode -CovePartnerId $Device.PartnerIdForHierarchy -CWMCompanyId ([int]$CWMcompany.ID) -CWMCompanyIdentifier $CWMcompany.identifier -CurrentExternalCode $Device.Reference -UpdateEnabled:$UpdateCoveReferences
                    }
                    
                    # Add to session cache for future lookups within this run
                    $Script:CompanyLookupCache += [PSCustomObject]@{
                        CompanyId = [int]$CWMcompany.ID
                        CompanyName = $CWMcompany.Name
                        MatchName = $matchName
                        Identifier = $CWMcompany.identifier
                        EndCustomer = $Device.EndCustomer
                        IsPlaceholder = $false
                        CompanyObject = $CWMcompany
                    }
                    
                    # OPTIMIZATION v10: Add to EndCustomer cache (store full company object)
                    if ($Device.EndCustomer) {
                        $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer] = $CWMcompany
                        if ($debugCWM) { Write-Host "[DEBUG]   [CACHE] Added EndCustomer '$($Device.EndCustomer)' → CWM Company [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                    }
                    return $CWMcompany
                }
            }
        } catch {
            # Continue to other methods
        }
    }

    # Strategy 3: Try End Customer exact name match
    # CRITICAL: Truncate to 50 chars to match what's actually stored in CWM
    $searchName = if ($matchName.Length -gt 50) { $matchName.Substring(0, 50) } else { $matchName }
    try {
        if ($debugCWM) {
            Write-Host "[DEBUG] Strategy 3: Query name=`"$searchName`" (active companies only)" -ForegroundColor Cyan
        }
        $CWMcompanyResults = Get-CWMcompany -condition "name=`"$searchName`" and deletedFlag = false"
        $Script:PerformanceMetrics.CompanyLookupCount++
        if ($debugCWM) {
            Write-Host "[DEBUG]   Get-CWMcompany returned $($CWMcompanyResults.Count) result(s)" -ForegroundColor Magenta
            if ($CWMcompanyResults) {
                Write-Host "[DEBUG]   First result: ID=$($CWMcompanyResults[0].ID), Name='$($CWMcompanyResults[0].Name)', deletedFlag=$($CWMcompanyResults[0].deletedFlag)" -ForegroundColor Magenta
            }
        }
        if ($CWMcompanyResults) {
            # Ensure single company (query may return array)
            $CWMcompany = $CWMcompanyResults | Select-Object -First 1
            
            # Check if company is deleted
            if ($CWMcompany.deletedFlag) {
                if ($debugCWM) {
                    Write-Warning "  Company '$matchName' found but marked as DELETED - skipping"
                }
                $CWMcompany = $null
            } else {
                if ($debugCWM) {
                    Write-Host "  Matched via End Customer Name: [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                }
                # Update Cove ExternalCode with CWM ID link
                if ($Device.PartnerIdForHierarchy) {
                    Update-CovePartnerExternalCode -CovePartnerId $Device.PartnerIdForHierarchy -CWMCompanyId ([int]$CWMcompany.ID) -CWMCompanyIdentifier $CWMcompany.identifier -CurrentExternalCode $Device.Reference -UpdateEnabled:$UpdateCoveReferences
                }
                
                # Add to session cache for future lookups within this run
                $Script:CompanyLookupCache += [PSCustomObject]@{
                    CompanyId = [int]$CWMcompany.ID
                    CompanyName = $CWMcompany.Name
                    MatchName = $matchName
                    Identifier = $CWMcompany.identifier
                    EndCustomer = $Device.EndCustomer
                    IsPlaceholder = $false
                    CompanyObject = $CWMcompany
                }
                
                # OPTIMIZATION v10: Add to EndCustomer cache (store full company object)
                if ($Device.EndCustomer) {
                    $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer] = $CWMcompany
                    if ($debugCWM) { Write-Host "[DEBUG]   [CACHE] Added EndCustomer '$($Device.EndCustomer)' → CWM Company [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                }
                return $CWMcompany
            }
        }
    } catch {
        # Continue to other methods
    }

    # Strategy 4: Try Reference match (from PF field - legacy reference match, not CWM:ID)
    if ($Device.Reference -and $Device.Reference -notmatch 'CWM:\d+') {
        try {
            $CWMcompanyResults = Get-CWMcompany -condition "name=`"$($Device.Reference)`" and deletedFlag = false"
            $Script:PerformanceMetrics.CompanyLookupCount++
            if ($CWMcompanyResults) {
                # Ensure single company (query may return array)
                $CWMcompany = $CWMcompanyResults | Select-Object -First 1
                
                # Check if company is deleted
                if ($CWMcompany.deletedFlag) {
                    if ($debugCWM) {
                        Write-Warning "  Company '$($Device.Reference)' found but marked as DELETED - skipping"
                    }
                    $CWMcompany = $null
                } else {
                    if ($debugCWM) {
                        Write-Host "  Matched via Reference Field: [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                    }
                    # Update Cove ExternalCode with CWM ID link
                    if ($Device.PartnerIdForHierarchy) {
                        Update-CovePartnerExternalCode -CovePartnerId $Device.PartnerIdForHierarchy -CWMCompanyId ([int]$CWMcompany.ID) -CWMCompanyIdentifier $CWMcompany.identifier -CurrentExternalCode $Device.Reference -UpdateEnabled:$UpdateCoveReferences
                    }
                    
                    # Add to session cache for future lookups within this run
                    $Script:CompanyLookupCache += [PSCustomObject]@{
                        CompanyId = [int]$CWMcompany.ID
                        CompanyName = $CWMcompany.Name
                        MatchName = $matchName
                        Identifier = $CWMcompany.identifier
                        EndCustomer = $Device.EndCustomer
                        IsPlaceholder = $false
                        CompanyObject = $CWMcompany
                    }
                    
                    # OPTIMIZATION v10: Add to EndCustomer cache (store full company object)
                    if ($Device.EndCustomer) {
                        $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer] = $CWMcompany
                        if ($debugCWM) { Write-Host "[DEBUG]   [CACHE] Added EndCustomer '$($Device.EndCustomer)' → CWM Company [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                    }
                    return $CWMcompany
                }
            }
        } catch {
            # Continue
        }
    }

    if ($debugCWM) {
        Write-Warning "No matching company found in ConnectWise for: $matchName | Ref: $($Device.Reference)"
    }
    
    # Only auto-create companies if parameter is enabled and this is an End Customer (not root/sub-root/site/reseller)
    if ($AutoCreateCompanies -and $Device.EndCustomer -and $Device.EndCustomerLevel -eq 'EndCustomer') {
        if ($debugCWM) {
            Write-Host "  Attempting to create new End Customer company via Sync script: $matchName" -ForegroundColor Yellow
        }
        
        # Call Cove2CWM-SyncCustomers script to create company with intelligent identifier
        try {
            $syncScriptPath = Join-Path $PSScriptRoot "Cove2CWM-SyncCustomers.v10.ps1"
            
            if (-not (Test-Path $syncScriptPath)) {
                # Try to find any version of the script
                $allVersions = Get-ChildItem -Path $PSScriptRoot -Filter "Cove2CWM-SyncCustomers*.ps1" -File | Sort-Object Name -Descending
                
                if ($allVersions.Count -ge 1) {
                    # Use latest version found
                    $syncScriptPath = $allVersions[0].FullName
                    if ($debugCWM) { Write-Host "  [SCRIPT] Using $($allVersions[0].Name) for company creation" -ForegroundColor Cyan }
                }
            }
            
            if (-not (Test-Path $syncScriptPath)) {
                Write-Warning "Cove2CWM-SyncCustomers.ps1 not found (any version)"
                Write-Warning "Cannot auto-create company. Please run Cove2CWM-SyncCustomers script manually first."
                return $null
            }
            
            # CRITICAL: Add placeholder to cache IMMEDIATELY to prevent duplicate creation attempts
            # This reserves the company name while sync script is running (prevents race condition)
            # Store BOTH full name (for Cove matching) and truncated name (for post-creation matching)
            $truncatedName = if ($matchName.Length -gt 50) { 
                $matchName.Substring(0, 50)
            } else { 
                $matchName 
            }
            
            $placeholderCompany = [PSCustomObject]@{
                CompanyId = -1  # Placeholder ID (will be updated after creation)
                CompanyName = $matchName  # Store FULL name for initial Cove matching
                MatchName = $matchName    # Store FULL name for initial Cove matching
                TruncatedName = $truncatedName  # Store truncated name for post-creation matching
                Identifier = ""
                EndCustomer = $Device.EndCustomer
                IsPlaceholder = $true
                CompanyObject = $null  # Will store full CWM company object after creation
            }
            $Script:CompanyLookupCache += $placeholderCompany
            if ($debugCWM) { Write-Host "  [CACHE] Added placeholder for '$matchName' (truncated: '$truncatedName') to prevent duplicate creation" -ForegroundColor Cyan }
            
            # Track this as a new company creation
            $Script:CompaniesCreatedViaHelper += [PSCustomObject]@{
                OriginalName = $matchName
                TruncatedName = $truncatedName
                EndCustomer = $Device.EndCustomer
                CreatedAt = Get-Date
            }
            
            # Call sync script with specific company creation (non-interactive mode)
            $syncResult = & $syncScriptPath -PartnerName $Script:PartnerName -CreateCompany $matchName -NonInteractive -ErrorAction Stop
            
            # Wait for CWM to fully propagate the new company before querying
            Write-Host "  [WAIT] Pausing 5 seconds for CWM to sync new company..." -ForegroundColor Gray
            Start-Sleep -Seconds 5
            
            # Re-query ConnectWise to get the newly created company
            # NOTE: If name was truncated (>50 chars), use exact match with truncated name
            if ($matchName.Length -gt 50) {
                # Use exact match with truncated name (no ... suffix)
                $searchName = $matchName.Substring(0, 50)
                $CWMcompanyResults = Get-CWMcompany -condition "name=`"$searchName`" and deletedFlag = false"
            } else {
                # Use exact match for non-truncated names
                $CWMcompanyResults = Get-CWMcompany -condition "name=`"$matchName`" and deletedFlag = false"
            }
            $Script:PerformanceMetrics.CompanyLookupCount++
            
            if ($CWMcompanyResults) {
                # Ensure single company (query may return array)
                $CWMcompany = $CWMcompanyResults | Select-Object -First 1
                
                Write-Host "  ✓ Company created successfully: [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                
                # Update placeholder with actual company data AND full object
                # CRITICAL: Keep MatchName as original full name, update CompanyName to actual CWM name (may be truncated)
                $placeholderCompany.CompanyId = [int]$CWMcompany.ID
                $placeholderCompany.CompanyName = $CWMcompany.Name  # Actual name from CWM (may be truncated to 50 chars)
                $placeholderCompany.TruncatedName = $CWMcompany.Name  # Store for future lookups
                $placeholderCompany.Identifier = $CWMcompany.identifier
                $placeholderCompany.CompanyObject = $CWMcompany  # Store full company object
                $placeholderCompany.IsPlaceholder = $false
                if ($debugCWM) { Write-Host "  [CACHE] Updated placeholder with full company object: [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                
                # Update Cove ExternalCode with CWM ID link
                if ($Device.PartnerIdForHierarchy) {
                    Update-CovePartnerExternalCode -CovePartnerId $Device.PartnerIdForHierarchy -CWMCompanyId ([int]$CWMcompany.ID) -CWMCompanyIdentifier $CWMcompany.identifier -CurrentExternalCode $Device.Reference -UpdateEnabled:$UpdateCoveReferences
                }
                
                # OPTIMIZATION v10: Add to EndCustomer cache for future lookups (store full company object)
                if ($Device.EndCustomer) {
                    $Script:EndCustomerToCWMCompanyCache[$Device.EndCustomer] = $CWMcompany
                    if ($debugCWM) { Write-Host "[DEBUG]   [CACHE] Added EndCustomer '$($Device.EndCustomer)' → CWM Company [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Cyan }
                }
                
                return $CWMcompany
            } else {
                Write-Warning "Company creation completed but company not found in CWM query"
            }
        }
        catch {
            Write-Warning "Failed to create company '$matchName' via Sync script: $($_.Exception.Message)"
            Write-Host "  Run this command manually: .\Cove2CWM-SyncCustomers.v10.ps1 -PartnerName `"$Script:PartnerName`" -CreateCompany `"$matchName`"" -ForegroundColor Yellow
        }
    }
    elseif (-not $AutoCreateCompanies -and $debugCWM) {
        Write-Host "  Auto-create companies disabled - skipping company creation" -ForegroundColor Yellow
        Write-Host "  To create missing companies, run: .\Cove2CWM-SyncCustomers.v10.ps1 -PartnerName `"$Script:PartnerName`"" -ForegroundColor Gray
    }
    elseif (-not $Device.EndCustomer -and $debugCWM) {
        Write-Host "  Skipping company creation - not an End Customer (Level: $($Device.EndCustomerLevel))" -ForegroundColor Yellow
    }
    elseif ($Device.EndCustomerLevel -ne 'EndCustomer' -and $debugCWM) {
        Write-Host "  Skipping company creation - partner is $($Device.EndCustomerLevel) level, not EndCustomer" -ForegroundColor Yellow
    }
    
    return $null
}  ## Match Cove partner/customer to ConnectWise company using multiple methods

Function Get-CWMTicketForDevice {
    param(
        [string]$DeviceName,
        [PSObject]$CWMCompany,
        [string]$IssueSeverity
    )
    
    # Search for open tickets with this device name in the summary
    # Format: "Cove: [DeviceType] DeviceName (ID: AccountId) - Severity - Description"
    # Exclude tickets already in closed status to prevent re-closure attempts
    $ticketFilter = "summary contains `"$DeviceName`" and closedFlag = false and board/name = `"$TicketBoard`" and status/name != `"$TicketClosedStatus`""
    
    if ($CWMCompany) {
        $ticketFilter += " and company/id = $($CWMCompany.id)"
    }
    
    try {
        $existingTickets = Get-CWMTicket -condition $ticketFilter -all
        $Script:PerformanceMetrics.TicketSearchCount++
        if ($existingTickets) {
            if ($debugCWM) {
                Write-Host "  Found $($existingTickets.Count) existing open ticket(s) for $DeviceName" -ForegroundColor Cyan
            }
            
            # Filter to exact device name match in summary (avoid partial matches)
            # Match format: "COVE DeviceType DeviceName #AccountId - ..."
            # Device type can be: 365, SVR, or WKS
            $escapedName = [regex]::Escape($DeviceName)
            $exactMatches = $existingTickets | Where-Object { $_.summary -match "COVE\s+(?:365|SVR|WKS)\s+$escapedName\s+#\d+\s+-" }
            
            if ($exactMatches) {
                if ($debugCWM) {
                    Write-Host "  Found exact match: ticket #$($exactMatches[0].id)" -ForegroundColor Green
                }
                return $exactMatches[0]  # Return first exact match
            } else {
                # No exact match found - don't return partial matches as they could be for different devices
                if ($debugCWM) {
                    Write-Host "  No exact match found for '$DeviceName' - no ticket to close/update" -ForegroundColor Gray
                }
                return $null
            }
        }
    } catch {
        Write-Warning "Error searching for existing tickets: $_"
    }
    
    return $null
}  ## Search for existing open tickets for a device

Function New-CWMTicketForDevice {
    param(
        [PSObject]$Device,
        [PSObject]$CWMCompany
    )
    
    if (-not $CWMCompany) {
        Write-Warning "Cannot create ticket - no matching ConnectWise company found for $($Device.PartnerName)"
        
        # Export ticket body to text file for review
        $exportFileName = "TicketBody_$($Device.DeviceName -replace '[^a-zA-Z0-9_-]','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $exportFilePath = Join-Path $ExportPath $exportFileName
        
        # Build ticket body (same logic as below)
        if ($Device.AccountType -eq 2) {
            # M365: Use cloud-properties structure
            $deviceUrl = "https://backup.management/#/device/$($Device.AccountID)/cloud-properties/office365/history"
        } else {
            # Standard backup devices: Use backup/overview view
            $viewId = '1662'
            $deviceUrl = "https://backup.management/#/backup/overview/view/$viewId(panel:device-properties/$($Device.AccountID)/summary)"
        }
        $readableDataSources = Convert-DataSourceCodes -DataSourceString $Device.DataSources
        $deviceType = if ($Device.AccountType -eq 2) { "Microsoft 365 Tenant" } else { "Server or Workstation" }
        $datasourceDetails = "  Datasource details not available (device detail lookup required)"
        if ($Device.DeviceSettings) {
            $errorMsgs = if ($Device.ErrorMessages) { $Device.ErrorMessages } else { @{} }
            $datasourceDetails = Get-DatasourceDetails -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources -ErrorMessages $errorMsgs -AccountType $Device.AccountType
        }
        $partnerDisplay = if ($Device.Site) {
            "$($Device.EndCustomer) | Site: $($Device.Site)"
        } else {
            $Device.EndCustomer
        }
        $isServer = ($Device.OSType -eq "2") -or ($Device.OS -match "Server")
        $deviceTypeDisplay = switch ($Device.Physicality) {
            "Virtual" { if ($isServer) { "Virtual Server" } else { "Virtual Workstation" } }
            "Physical" { if ($isServer) { "Physical Server" } else { "Physical Workstation" } }
            default { "Server or Workstation" }
        }
        
        $ticketBody = @"
========================================
TICKET BODY EXPORT - COMPANY NOT FOUND
========================================

ConnectWise Company Match Attempted For:
  EndCustomer: $($Device.EndCustomer)
  PartnerName: $($Device.PartnerName)
  Reference: $($Device.Reference)
  Site: $($Device.Site)

Result: NO MATCHING COMPANY FOUND IN CONNECTWISE

========================================
TICKET DETAILS THAT WOULD BE CREATED:
========================================

Cove Data Protection Backup Alert

$(if ($Device.AccountType -eq 2) {
    # M365 format
"M365 Tenant: $($Device.DeviceName)
Customer: $partnerDisplay"
} else {
    # System format
"Device: $($Device.DeviceName)
Computer Name: $($Device.ComputerName)
$(if ($Device.DeviceAlias) { "Alias: $($Device.DeviceAlias)`n" })Customer: $partnerDisplay"
})
Reference: $($Device.Reference)
Issue Severity: $($Device.IssueSeverity)
Issue Description: $($Device.IssueDescription)
Last Timestamp: $($Device.TimeStamp)

Device Details:

$(if ($Device.AccountType -eq 2) {
    # M365 tenant details (matching structure)
"Storage Usage:
  $(Format-AlignedLabel 'Selected Data' 20)$($Device.SelectedGB) GB
  $(Format-AlignedLabel 'Used Storage' 20)$($Device.UsedGB) GB
  $(Format-AlignedLabel 'Storage Location' 20)$($Device.Location)"
} else {
    # System details
"Hardware Information:
  $(Format-AlignedLabel 'OS' 20)$($Device.OS)
  $(Format-AlignedLabel 'Mfg\Model' 20)$($Device.Manufacturer) | $($Device.Model)
  $(Format-AlignedLabel 'Device Type' 20)$deviceTypeDisplay
  $(Format-AlignedLabel 'IP Address' 20)$($Device.IPAddress)
  $(Format-AlignedLabel 'External IP' 20)$($Device.ExternalIP)

Backup Configuration:
  $(Format-AlignedLabel 'Backup Profile' 20)$($Device.Profile) (ID: $($Device.ProfileID))
  $(Format-AlignedLabel 'Retention Policy' 20)$($Device.Product) (ID: $($Device.ProductID))
  $(Format-AlignedLabel 'Storage Location' 20)$($Device.Location)
  $(Format-AlignedLabel 'Timezone Offset' 20)$($Device.TimezoneOffset)"
})

Datasource Details:
$datasourceDetails

View Device in Cove Portal:
$deviceUrl

This ticket was automatically created by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)

========================================
"@
        
        try {
            $ticketBody | Set-Content -Path $exportFilePath -Encoding UTF8
            Write-Host "  Exported ticket body to: $exportFilePath" -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Failed to export ticket body: $($_.Exception.Message)"
        }
        
        return $null
    }
    
    $severityInfo = $Script:IssueSeverity[$Device.IssueSeverity]
    
    # Determine device type label for ticket summary (Option D format)
    $deviceTypeLabel = if ($Device.AccountType -eq 2) {
        "365"
    } else {
        # Server or Workstation - use OT field (2=Server, 1=Workstation)
        if ($Device.OSType -eq "2") { "SVR" } else { "WKS" }
    }

    # Get datasource codes for summary (Option D format)
    # For stale tickets: Check last success timestamps against threshold
    # For failed tickets: Check status codes for failures
    $staleHoursParam = if ($Device.IssueSeverity -eq 'Stale') {
        # Determine stale threshold based on device type
        if ($Device.OSType -eq "2" -or $Device.OS -match "Server") {
            $StaleHoursServers
        } elseif ($Device.AccountType -eq 2) {
            $StaleHoursM365
        } else {
            $StaleHoursWorkstations
        }
    } else { 0 }
    
    $datasourceCodes = Get-DatasourceCodes -DataSourceString $Device.DataSources -DeviceSettings $Device.DeviceSettings -StaleHours $staleHoursParam
    
    if ($DebugCDP) {
        Write-Host "[DEBUG] Get-DatasourceCodes called:" -ForegroundColor Magenta
        Write-Host "  Device: $($Device.DeviceName)" -ForegroundColor Gray
        Write-Host "  DataSources: $($Device.DataSources)" -ForegroundColor Gray
        Write-Host "  IssueSeverity: $($Device.IssueSeverity)" -ForegroundColor Gray
        Write-Host "  StaleHoursParam: $staleHoursParam" -ForegroundColor Gray
        Write-Host "  Returned codes: '$datasourceCodes'" -ForegroundColor Gray
    }
    
    $datasourcePrefix = if ($datasourceCodes) { "$datasourceCodes " } else { "" }
    
    # Extract and truncate error message for summary (Option D format)
    # Format: "COVE WKS devicename #AccountId - DS Failed - Error message"
    $errorSummary = ""
    if ($Device.LastError -and $Device.LastError -ne "N/A" -and $Device.LastError -ne "Backup successful") {
        $errorMsg = $Device.LastError
        
        # Smart truncation for common error patterns
        if ($errorMsg -match "Could not resolve hostname.*Error #(\d+)") {
            $errorSummary = "DNS error #$($matches[1])"
        }
        elseif ($errorMsg -match "The process cannot access the file.*because it is being used") {
            # Extract filename if possible
            if ($errorMsg -match "'([^']+)'") {
                $file = $matches[1]
                if ($file.Length -gt 30) { $file = "..." + $file.Substring($file.Length - 27) }
                $errorSummary = "File locked: $file"
            } else {
                $errorSummary = "File locked"
            }
        }
        elseif ($errorMsg -match "Mailbox (inactive|soft-deleted)") {
            $errorSummary = "Mailbox $($matches[1])"
        }
        elseif ($errorMsg -match "No data available") {
            $errorSummary = "No data available"
        }
        else {
            # Generic truncation at word boundary (max 50 chars for error portion)
            if ($errorMsg.Length -gt 50) {
                $truncated = $errorMsg.Substring(0, 47)
                $lastSpace = $truncated.LastIndexOf(' ')
                if ($lastSpace -gt 20) {
                    $errorSummary = $truncated.Substring(0, $lastSpace) + "..."
                } else {
                    $errorSummary = $truncated + "..."
                }
            } else {
                $errorSummary = $errorMsg
            }
        }
    }
    
    # Create abbreviated issue description for Option D format
    $abbreviatedIssue = if ($Device.IssueDescription -match "Backup failed" -and $errorSummary) {
        "${datasourcePrefix}Failed - $errorSummary"
    } elseif ($Device.IssueDescription -match "Stale backup") {
        if ($Device.IssueDescription -match "\((~[^)]+)\)") {
            $timeRef = $matches[1] -replace '~', '' -replace ' days?', 'd' -replace ' hr', 'h' -replace ' min', 'm' -replace ' ago', ''
            "${datasourcePrefix}Stale ($timeRef)"
        } else {
            "${datasourcePrefix}Stale"
        }
    } elseif ($Device.IssueDescription -match "No successful.*in ([\d.]+) hours") {
        $hours = [double]$matches[1]
        $timeText = Format-HoursAsRelativeTime -Hours $hours -IncludeParentheses:$false
        "${datasourcePrefix}No success ($timeText)"
    } elseif ($Device.IssueDescription -match "completed with.*error") {
        if ($Device.IssueDescription -match "(\d+) error") {
            "${datasourcePrefix}Failed($($matches[1]))"
        } else {
            "${datasourcePrefix}Errors"
        }
    } else {
        # Generic fallback
        "${datasourcePrefix}Issue"
    }

    # Store $abbreviatedIssue on Device object so it persists to ticket creation/update loop
    $Device | Add-Member -NotePropertyName "AbbreviatedIssue" -NotePropertyValue $abbreviatedIssue -Force
    
    # DEBUG: Verify property was set
    if ($DebugCDP) {
        Write-Host "  [DEBUG] Set AbbreviatedIssue on device $($Device.DeviceName): '$abbreviatedIssue'" -ForegroundColor Magenta
    }
    
    # Build ticket summary (Option D format - datasource-first)
    # Format: "COVE WKS devicename #AccountId - DS Failed - Error message"
    $summary = "COVE $deviceTypeLabel $($Device.DeviceName) #$($Device.AccountId) - $abbreviatedIssue"
    
    # Truncate to 100 chars max (ConnectWise limit)
    if ($summary.Length -gt 100) {
        $summary = $summary.Substring(0, 97) + "..."
    }
    
    # Build device URL for Cove portal
    if ($Device.AccountType -eq 2) {
        # M365: Use cloud-properties structure
        $deviceUrl = "https://backup.management/#/device/$($Device.AccountID)/cloud-properties/office365/history"
    } else {
        # Standard backup devices: Use backup/overview view
        $viewId = '1662'
        $deviceUrl = "https://backup.management/#/backup/overview/view/$viewId(panel:device-properties/$($Device.AccountID)/summary)"
    }
    
    # Convert data source codes to readable names
    $readableDataSources = Convert-DataSourceCodes -DataSourceString $Device.DataSources
    
    # Determine device type display for ticket body
    $deviceType = if ($Device.AccountType -eq 2) { "Microsoft 365 Tenant" } else { "Server or Workstation" }
    
    # Get datasource details (needs DeviceResult object with Settings)
    $datasourceDetails = "  Datasource details not available (device detail lookup required)"
    if ($Device.DeviceSettings) {
        $errorMsgs = if ($Device.ErrorMessages) { $Device.ErrorMessages } else { @{} }
        $datasourceDetails = Get-DatasourceDetails -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources -ErrorMessages $errorMsgs -AccountType $Device.AccountType
    }
    
    # Build Partner/Site display string
    $partnerDisplay = if ($Device.Site) {
        "$($Device.EndCustomer) | Site: $($Device.Site)"
    } else {
        $Device.EndCustomer
    }
    
    # Build ticket description based on AccountType (M365 vs Systems)
    if ($Device.AccountType -eq 2) {
        # M365 Tenant - Use M365-specific template (matching system format structure)
        # Pre-calculate formatted time strings for M365 ticket display (based on datasource-level times)
        
        # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
        $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $worstHoursSince
        } else { '' }
        
        # Build Oldest Problem row only if there's an actual problem (not "Unknown")
        $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
            "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
        } else { $null }
        
        # Calculate and format Creation Date for ticket display
        $creationDateFormatted = if ($Device.Creation -and $Device.Creation -is [DateTime]) {
            if ($UseLocalTime) {
                $localCreation = [System.TimeZoneInfo]::ConvertTimeFromUtc($Device.Creation, [System.TimeZoneInfo]::Local)
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localCreation)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                "$($localCreation.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else {
                "$($Device.Creation.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            }
        } else { 'N/A' }
        
        $creationDateRelative = if ($Device.Creation -and $Device.Creation -is [DateTime]) {
            $hoursSinceCreation = [Math]::Round(((Get-Date).ToUniversalTime() - $Device.Creation).TotalHours, 1)
            Format-HoursAsRelativeTime $hoursSinceCreation
        } else { '' }
        
        # Calculate relative time for Last Timestamp (M365 template)
        # Use TimeStamp field (TS column) which includes all sessions (InProcess, Success, Failed)
        # For devices with no timestamp at all, show creation-based message
        $lastTimestampDisplay = if ($Device.TimeStamp -and $Device.TimeStamp -ne 'N/A') {
            $Device.TimeStamp
        } elseif ($creationDateFormatted -and $creationDateFormatted -ne 'N/A') {
            "Never (device created: $creationDateFormatted)"
        } else {
            'N/A'
        }
        
        # Calculate relative time from TimeStamp field (parse back from formatted string to get DateTime)
        $lastTimestampRelative = if ($Device.TimeStamp -and $Device.TimeStamp -ne 'N/A') {
            # TimeStamp is already formatted string, need to get original DateTime from device data
            # Use TS column from raw device data
            $tsUnix = $Device.DeviceSettings.TS -join ''
            if ($tsUnix) {
                $tsDateTime = Convert-UnixTimeToDateTime $tsUnix
                if ($tsDateTime -and $tsDateTime -is [DateTime]) {
                    $hoursSinceTS = [Math]::Round(((Get-Date).ToUniversalTime() - $tsDateTime).TotalHours, 1)
                    Format-HoursAsRelativeTime $hoursSinceTS
                } else { '' }
            } else { '' }
        } else { '' }
        
        $description = @"
Cove Data Protection Backup Alert

M365 Tenant          : $($Device.DeviceName) (ID:$($Device.AccountID))
Customer             : $partnerDisplay
$(if ($Device.Reference) { "Reference            : $($Device.Reference)`n" })$(if ($Device.Notes) { "Notes                : $($Device.Notes)`n" })Severity             : $($Device.IssueSeverity)
Description          : $($Device.IssueDescription)
Ticket Criteria      : Lookback < $DaysBack days | Stale > $StaleHoursM365 hours
Creation Date        : $creationDateFormatted $creationDateRelative
Last Timestamp       : $lastTimestampDisplay $lastTimestampRelative
$(if ($oldestProblemRow) { "$oldestProblemRow`n" })
─────────────────────────────────────────────────────────────────
TENANT DETAILS

Storage Usage:
  Selected Data       : $($Device.SelectedSize)
  Used Storage        : $($Device.UsedStorage)
  Storage Location    : $($Device.Location)
  Timezone Offset     : $($Device.TimezoneOffset)

─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS
$datasourceDetails

─────────────────────────────────────────────────────────────────
View Device: $deviceUrl

This ticket was automatically created by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)
"@
    } else {
        # Server/Workstation - Use system template
        # Determine device type based on Physicality and OSType
        # OT values: 1=Workstation, 2=Server (verified from API)
        $isServer = ($Device.OSType -eq "2") -or ($Device.OS -match "Server")
        
        $deviceTypeDisplay = switch ($Device.Physicality) {
            "Virtual" { 
                if ($isServer) { "Virtual Server" } else { "Virtual Workstation" }
            }
            "Physical" {
                if ($isServer) { "Physical Server" } else { "Physical Workstation" }
            }
            default { "Server or Workstation" }
        }
        
        # Pre-calculate formatted time strings for ticket display (based on datasource-level times)
        
        # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
        $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $worstHoursSince
        } else { '' }
        
        # Build Oldest Problem row only if there's an actual problem (not "Unknown")
        $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
            "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
        } else { $null }
        
        # Calculate and format Creation Date for ticket display
        $creationDateFormatted = if ($Device.Creation -and $Device.Creation -is [DateTime]) {
            if ($UseLocalTime) {
                $localCreation = [System.TimeZoneInfo]::ConvertTimeFromUtc($Device.Creation, [System.TimeZoneInfo]::Local)
                $tzName = if ([System.TimeZoneInfo]::Local.IsDaylightSavingTime($localCreation)) {
                    [System.TimeZoneInfo]::Local.DaylightName
                } else {
                    [System.TimeZoneInfo]::Local.StandardName
                }
                $tzAbbr = ($tzName -split ' ' | ForEach-Object { $_[0] }) -join ''
                "$($localCreation.ToString('yyyy-MM-dd HH:mm:ss')) ($tzAbbr)"
            } else {
                "$($Device.Creation.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
            }
        } else { 'N/A' }
        
        $creationDateRelative = if ($Device.Creation -and $Device.Creation -is [DateTime]) {
            $hoursSinceCreation = [Math]::Round(((Get-Date).ToUniversalTime() - $Device.Creation).TotalHours, 1)
            Format-HoursAsRelativeTime $hoursSinceCreation
        } else { '' }
        
        # Calculate relative time for Last Timestamp (Server/Workstation template)
        # Use TimeStamp field (TS column) which includes all sessions (InProcess, Success, Failed)
        # For devices with no timestamp at all, show creation-based message
        $lastTimestampDisplay = if ($Device.TimeStamp -and $Device.TimeStamp -ne 'N/A') {
            $Device.TimeStamp
        } elseif ($creationDateFormatted -and $creationDateFormatted -ne 'N/A') {
            "Never (device created: $creationDateFormatted)"
        } else {
            'N/A'
        }
        
        # Calculate relative time from TimeStamp field (parse back from formatted string to get DateTime)
        $lastTimestampRelative = if ($Device.TimeStamp -and $Device.TimeStamp -ne 'N/A') {
            # TimeStamp is already formatted string, need to get original DateTime from device data
            # Use TS column from raw device data
            $tsUnix = $Device.DeviceSettings.TS -join ''
            if ($tsUnix) {
                $tsDateTime = Convert-UnixTimeToDateTime $tsUnix
                if ($tsDateTime -and $tsDateTime -is [DateTime]) {
                    $hoursSinceTS = [Math]::Round(((Get-Date).ToUniversalTime() - $tsDateTime).TotalHours, 1)
                    Format-HoursAsRelativeTime $hoursSinceTS
                } else { '' }
            } else { '' }
        } else { '' }
        
        # Parse OS to get clean format like "macOS Sonoma (14.7.7), Intel-based" or "Windows Server 2022"
        $osDisplay = $Device.OS
        if ($Device.OS -match 'macOS') {
            # Extract macOS version details
            $osDisplay = $Device.OS  # Keep full format from API
        }
        
        # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
        $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $worstHoursSince
        } else { '' }
        
        # Build Oldest Problem row only if there's an actual problem (not "Unknown")
        $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
            "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
        } else { $null }
        
        $description = @"
Cove Data Protection Backup Alert

Device               : $($Device.DeviceName) (ID:$($Device.AccountID))
Computer             : $($Device.ComputerName)
$(if ($Device.DeviceAlias) { "Alias                : $($Device.DeviceAlias)`n" })Customer             : $partnerDisplay
$(if ($Device.Reference) { "Reference            : $($Device.Reference)`n" })$(if ($Device.Notes) { "Notes                : $($Device.Notes)`n" })Severity             : $($Device.IssueSeverity)
Description          : $($Device.IssueDescription)
Ticket Criteria      : Lookback < $DaysBack days | Stale > $(if ($Device.OSType -eq '2' -or $Device.OS -match 'Server') { "$StaleHoursServers" } else { "$StaleHoursWorkstations" }) hours
Creation Date        : $creationDateFormatted $creationDateRelative
Last Timestamp       : $lastTimestampDisplay $lastTimestampRelative
$(if ($oldestProblemRow) { "$oldestProblemRow`n" })
─────────────────────────────────────────────────────────────────
DEVICE DETAILS

Hardware Information:
  OS                  : $osDisplay
  Mfg|Model           : $($Device.Manufacturer) | $($Device.Model)
  Device Type         : $deviceTypeDisplay$(if ($Device.CPUCores -and $Device.CPUCores -ne '') { " | $($Device.CPUCores) Cores" })$(if ($Device.RAMSizeGB -and $Device.RAMSizeGB -ne 'N/A') { " | $($Device.RAMSizeGB) GB RAM" })
  IP Address          : $($Device.IPAddress)
  External IP         : $($Device.ExternalIP)

Backup Configuration:
$(if ($Device.Profile -and $Device.ProfileID) { "  Backup Profile      : $($Device.Profile) (ID: $($Device.ProfileID))`n" })  Retention Policy    : $($Device.Product) (ID: $($Device.ProductID))
  Storage Location    : $($Device.Location)
  Timezone Offset     : $($Device.TimezoneOffset)

─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS
$datasourceDetails

─────────────────────────────────────────────────────────────────
View Device: $deviceUrl

This ticket was automatically created by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)
"@
    }

    # Build ticket object
    # Determine priority based on device type (Server/Workstation/M365)
    $ticketPriority = Get-TicketPriorityForDevice -Device $Device
    
    $ticketParams = @{
        summary = $summary
        company = @{ id = $CWMCompany.id }
        board = @{ name = $TicketBoard }
        status = @{ name = $TicketStatus }
        priority = @{ name = $ticketPriority }
        initialDescription = $description
    }
    
    try {
        if ($CreateTickets) {
            $newTicket = New-CWMTicket @ticketParams
            if ($newTicket -and $newTicket.id -and $newTicket.id -gt 0) {
                $Script:ProcessedTickets.Created += $newTicket.id  # Track to prevent duplicates
                $ticketUrl = Get-CWMTicketURL -TicketID $newTicket.id
                Write-Host "  Created ticket #$($newTicket.id) for $($Device.DeviceName) at company $($CWMCompany.Name)" -ForegroundColor Green
                Write-Host "    URL: $ticketUrl" -ForegroundColor Gray
                return $newTicket
            } else {
                Write-Warning "Ticket creation failed for $($Device.DeviceName) - no valid ticket ID returned"
                return $null
            }
        } else {
            Write-Host "  [TEST MODE] Would create ticket for $($Device.DeviceName) at company $($CWMCompany.Name) - $($Device.IssueDescription)" -ForegroundColor Yellow
            return $null
        }
    } catch {
        $errorMessage = $_.Exception.Message
        
        # Check if this is a retryable error (network/timeout issues)
        $isRetryable = $errorMessage -match '(timeout|forcibly closed|connection|unreachable|timed out|503|502|500)'
        
        if ($isRetryable) {
            Write-Warning "Network error creating ticket for $($Device.DeviceName): $errorMessage"
            
            # IMMEDIATE RETRY: Single retry with 2-second delay for transient network issues
            Write-Host "    ⏱️  Waiting 2 seconds before retry..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            
            try {
                Write-Host "    🔄 Retrying ticket creation..." -ForegroundColor Cyan
                $newTicket = New-CWMTicket @ticketParams
                
                if ($newTicket -and $newTicket.id -and $newTicket.id -gt 0) {
                    $Script:ProcessedTickets.Created += $newTicket.id  # Track to prevent duplicates
                    $ticketUrl = Get-CWMTicketURL -TicketID $newTicket.id
                    Write-Host "    ✓ Retry successful - Created ticket #$($newTicket.id)" -ForegroundColor Green
                    Write-Host "      URL: $ticketUrl" -ForegroundColor Gray
                    return $newTicket
                }
            } catch {
                Write-Warning "Retry also failed for $($Device.DeviceName): $($_.Exception.Message)"
            }
        } else {
            # Non-retryable error (bad data, permissions, etc.)
            Write-Warning "Non-retryable error creating ticket for $($Device.DeviceName): $errorMessage"
        }
        
        return $null
    }
}  ## Create a new ConnectWise ticket for a device issue

Function Update-CWMTicketForDevice {
    param(
        [PSObject]$Ticket,
        [PSObject]$Device
    )
    
    # Skip if already updated in this session
    if ($Script:ProcessedTickets.Updated -contains $Ticket.id) {
        if ($debugCWM) {
            Write-Host "  Skipping duplicate update for ticket #$($Ticket.id)" -ForegroundColor Gray
        }
        return $false
    }
    
    # Build device URL for Cove portal
    if ($Device.AccountType -eq 2) {
        # M365: Use cloud-properties structure
        $deviceUrl = "https://backup.management/#/device/$($Device.AccountID)/cloud-properties/office365/history"
    } else {
        # Standard backup devices: Use backup/overview view
        $viewId = '1662'
        $deviceUrl = "https://backup.management/#/backup/overview/view/$viewId(panel:device-properties/$($Device.AccountID)/summary)"
    }
    
    # Get datasource details
    $datasourceDetails = "  Datasource details not available"
    if ($Device.DeviceSettings) {
        $errorMsgs = if ($Device.ErrorMessages) { $Device.ErrorMessages } else { @{} }
        $datasourceDetails = Get-DatasourceDetails -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources -ErrorMessages $errorMsgs -AccountType $Device.AccountType
    }
    
    # Build new summary with current severity (Option D format)
    $deviceTypeLabel = if ($Device.AccountType -eq 2) {
        "365"
    } elseif ($Device.OSType -eq "2") {
        "SVR"
    } else {
        "WKS"
    }
    
    # Get datasource codes for summary (Option D format)
    # For stale tickets: Check last success timestamps against threshold
    # For failed tickets: Check status codes for failures
    $staleHoursParam = if ($Device.IssueSeverity -eq 'Stale') {
        # Determine stale threshold based on device type
        if ($Device.OSType -eq "2" -or $Device.OS -match "Server") {
            $StaleHoursServers
        } elseif ($Device.AccountType -eq 2) {
            $StaleHoursM365
        } else {
            $StaleHoursWorkstations
        }
    } else { 0 }
    
    $datasourceCodes = Get-DatasourceCodes -DataSourceString $Device.DataSources -DeviceSettings $Device.DeviceSettings -StaleHours $staleHoursParam
    $datasourcePrefix = if ($datasourceCodes) { "$datasourceCodes " } else { "" }
    
    # Extract and truncate error message for summary (Option D format)
    $errorSummary = ""
    if ($Device.LastError -and $Device.LastError -ne "N/A" -and $Device.LastError -ne "Backup successful") {
        $errorMsg = $Device.LastError
        
        # Smart truncation for common error patterns
        if ($errorMsg -match "Could not resolve hostname.*Error #(\d+)") {
            $errorSummary = "DNS error #$($matches[1])"
        }
        elseif ($errorMsg -match "The process cannot access the file.*because it is being used") {
            if ($errorMsg -match "'([^']+)'") {
                $file = $matches[1]
                if ($file.Length -gt 30) { $file = "..." + $file.Substring($file.Length - 27) }
                $errorSummary = "File locked: $file"
            } else {
                $errorSummary = "File locked"
            }
        }
        elseif ($errorMsg -match "Mailbox (inactive|soft-deleted)") {
            $errorSummary = "Mailbox $($matches[1])"
        }
        elseif ($errorMsg -match "No data available") {
            $errorSummary = "No data available"
        }
        else {
            # Generic truncation at word boundary (max 50 chars)
            if ($errorMsg.Length -gt 50) {
                $truncated = $errorMsg.Substring(0, 47)
                $lastSpace = $truncated.LastIndexOf(' ')
                if ($lastSpace -gt 20) {
                    $errorSummary = $truncated.Substring(0, $lastSpace) + "..."
                } else {
                    $errorSummary = $truncated + "..."
                }
            } else {
                $errorSummary = $errorMsg
            }
        }
    }
    
    # Create abbreviated issue description for Option D format
    $abbreviatedIssue = if ($Device.IssueDescription -match "Backup failed" -and $errorSummary) {
        "${datasourcePrefix}Failed - $errorSummary"
    } elseif ($Device.IssueDescription -match "Stale backup") {
        if ($Device.IssueDescription -match "\((~[^)]+)\)") {
            $timeRef = $matches[1] -replace '~', '' -replace ' days?', 'd' -replace ' hr', 'h' -replace ' min', 'm' -replace ' ago', ''
            "${datasourcePrefix}Stale ($timeRef)"
        } else {
            "${datasourcePrefix}Stale"
        }
    } elseif ($Device.IssueDescription -match "No successful.*in ([\d.]+) hours") {
        $hours = [double]$matches[1]
        $timeText = Format-HoursAsRelativeTime -Hours $hours -IncludeParentheses:$false
        "${datasourcePrefix}No success ($timeText)"
    } elseif ($Device.IssueDescription -match "completed with.*error") {
        if ($Device.IssueDescription -match "(\d+) error") {
            "${datasourcePrefix}Failed($($matches[1]))"
        } else {
            "${datasourcePrefix}Errors"
        }
    } else {
        # Generic fallback
        "${datasourcePrefix}Issue"
    }
    
    # Build new summary (Option D format)
    $newSummary = "COVE $deviceTypeLabel $($Device.DeviceName) #$($Device.AccountId) - $abbreviatedIssue"
    # Truncate to 100 chars max (ConnectWise limit)
    if ($newSummary.Length -gt 100) {
        $newSummary = $newSummary.Substring(0, 97) + "..."
    }
    
    # Check if summary needs updating (severity changed)
    $summaryNeedsUpdate = $false
    if ($Ticket.summary -ne $newSummary) {
        $summaryNeedsUpdate = $true
        if ($debugCWM) {
            Write-Host "  Summary change detected:" -ForegroundColor Yellow
            Write-Host "    Old: $($Ticket.summary)" -ForegroundColor Gray
            Write-Host "    New: $newSummary" -ForegroundColor Gray
        }
    }
    
    # Pre-calculate formatted time strings for note display (based on datasource-level times)
    
    # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
    $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
        $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
        Format-HoursAsRelativeTime $worstHoursSince
    } else { '' }
    
    # Build Oldest Problem row only if there's an actual problem (not "Unknown")
    $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
        "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
    } else { $null }
    
    $note = @"
Updated Status - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)

Severity             : $($Device.IssueSeverity)
Description          : $($Device.IssueDescription)
Ticket Criteria      : Lookback < $DaysBack days | Stale > $(if ($Device.AccountType -eq '2') { "$StaleHoursM365" } elseif ($Device.OSType -eq '2' -or $Device.OS -match 'Server') { "$StaleHoursServers" } else { "$StaleHoursWorkstations" }) hours
Last Timestamp       : $($Device.TimeStamp)
$(if ($oldestProblemRow) { "$oldestProblemRow`n" })
─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS
$datasourceDetails

─────────────────────────────────────────────────────────────────
View Device: $deviceUrl

This update was automatically generated by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)
"@

    try {
        if ($UpdateTickets) {
            # Add note to ticket
            $noteParams = @{
                ticketId = $Ticket.id
                text = $note
                detailDescriptionFlag = $true
                internalAnalysisFlag = $false
            }
            $noteResult = New-CWMTicketNote @noteParams -ErrorAction Stop
            
            if (-not $noteResult -or -not $noteResult.id) {
                Write-Warning "Failed to add note to ticket #$($Ticket.id)"
                return $false
            }
            
            # Update summary if severity changed
            if ($summaryNeedsUpdate) {
                try {
                    $summaryUpdateParams = @{
                        id = $Ticket.id
                        Operation = "replace"
                        Path = "summary"
                        Value = $newSummary
                    }
                    $summaryUpdateResult = Update-CWMTicket @summaryUpdateParams -ErrorAction Stop
                    
                    if ($summaryUpdateResult -and $summaryUpdateResult.id) {
                        if ($debugCWM) {
                            Write-Host "  ✓ Updated ticket summary to reflect current severity" -ForegroundColor Green
                        }
                    }
                } catch {
                    Write-Warning "Failed to update ticket summary for #$($Ticket.id): $_"
                    # Continue even if summary update fails - note was added successfully
                }
            }
            
            $Script:ProcessedTickets.Updated += $Ticket.id
            $ticketUrl = Get-CWMTicketURL -TicketID $Ticket.id
            Write-Host "  Updated ticket #$($Ticket.id) for $($Device.DeviceName)" -ForegroundColor Cyan
            Write-Host "    URL: $ticketUrl" -ForegroundColor Gray
            return $true
        } else {
            Write-Host "  [TEST MODE] Would update ticket #$($Ticket.id) for $($Device.DeviceName)" -ForegroundColor Yellow
            if ($summaryNeedsUpdate) {
                Write-Host "    Would update summary to: $newSummary" -ForegroundColor Gray
            }
            return $false
        }
    } catch {
        $errorMessage = $_.Exception.Message
        
        # Check if this is a retryable error (network/timeout issues)
        $isRetryable = $errorMessage -match '(timeout|forcibly closed|connection|unreachable|timed out|503|502|500)'
        
        if ($isRetryable) {
            Write-Warning "Network error updating ticket #$($Ticket.id): $errorMessage"
            
            # IMMEDIATE RETRY: Single retry with 2-second delay for transient network issues
            Write-Host "    ⏱️  Waiting 2 seconds before retry..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            
            try {
                Write-Host "    🔄 Retrying ticket update..." -ForegroundColor Cyan
                
                # Retry the note add
                $noteParams = @{
                    ticketId = $Ticket.id
                    text = $note
                    detailDescriptionFlag = $true
                    internalAnalysisFlag = $false
                }
                $noteResult = New-CWMTicketNote @noteParams -ErrorAction Stop
                
                if ($noteResult -and $noteResult.id) {
                    # Retry summary update if needed
                    if ($summaryNeedsUpdate) {
                        $summaryUpdateParams = @{
                            id = $Ticket.id
                            Operation = "replace"
                            Path = "summary"
                            Value = $newSummary
                        }
                        Update-CWMTicket @summaryUpdateParams -ErrorAction Stop | Out-Null
                    }
                    
                    $Script:ProcessedTickets.Updated += $Ticket.id
                    Write-Host "    ✓ Retry successful for ticket #$($Ticket.id)" -ForegroundColor Green
                    return $true
                }
            } catch {
                Write-Warning "Retry also failed for ticket #$($Ticket.id): $($_.Exception.Message)"
            }
        } else {
            # Non-retryable error (bad data, permissions, etc.)
            Write-Warning "Non-retryable error updating ticket #$($Ticket.id): $errorMessage"
        }
        
        return $false
    }
}  ## Update an existing ConnectWise ticket with new information

Function Close-CWMTicketForDevice {
    param(
        [PSObject]$Ticket,
        [PSObject]$Device
    )
    
    # Skip if already closed in this session
    if ($Script:ProcessedTickets.Closed -contains $Ticket.id) {
        if ($debugCWM) {
            Write-Host "  Skipping duplicate close for ticket #$($Ticket.id)" -ForegroundColor Gray
        }
        return $false
    }
    
    # Skip if ticket status is already the closed status
    if ($Ticket.status -and $Ticket.status.name -eq $TicketClosedStatus) {
        if ($debugCWM) {
            Write-Host "  Skipping ticket #$($Ticket.id) - already in '$TicketClosedStatus' status" -ForegroundColor Gray
        }
        return $false
    }
    
    # Skip if ticket is already marked as closed (closedFlag)
    if ($Ticket.closedFlag -eq $true) {
        if ($debugCWM) {
            Write-Host "  Skipping ticket #$($Ticket.id) - already marked as closed (closedFlag=true)" -ForegroundColor Gray
        }
        return $false
    }
    
    # Build device URL for portal link
    if ($Device.AccountType -eq 2) {
        # M365: Use cloud-properties structure
        $deviceUrl = "https://backup.management/#/device/$($Device.AccountID)/cloud-properties/office365/history"
    } else {
        # Standard backup devices: Use backup/overview view
        $viewId = '1662'
        $deviceUrl = "https://backup.management/#/backup/overview/view/$viewId(panel:device-properties/$($Device.AccountID)/summary)"
    }
    
    # Build close note based on AccountType (M365 vs Systems)
    if ($Device.AccountType -eq 2) {
        # M365 Tenant close note
        
        # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
        $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $worstHoursSince
        } else { '' }
        
        # Build Oldest Problem row only if there's an actual problem (not "Unknown")
        $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
            "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
        } else { $null }
        
        # Get datasource details for close note
        $datasourceDetails = "  Datasource details not available"
        if ($Device.DeviceSettings) {
            $errorMsgs = if ($Device.ErrorMessages) { $Device.ErrorMessages } else { @{} }
            $datasourceDetails = Get-DatasourceDetails -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources -ErrorMessages $errorMsgs -AccountType $Device.AccountType
        }
        
        # Calculate recovery statistics
        $recoveryStats = ""
        if ($Device.WorstDatasourceTime -and $Device.WorstDatasourceTime -ne 'N/A' -and $Device.WorstDatasourceTime -is [DateTime]) {
            $failureStart = $Device.WorstDatasourceTime
            $recoveryTime = ((Get-Date).ToUniversalTime() - $failureStart).TotalHours
            $recoveryDays = [math]::Floor($recoveryTime / 24)
            $recoveryHours = [math]::Floor($recoveryTime % 24)
            $recoveryStats = "`nRecovery Statistics:`n"
            $recoveryStats += "  Time in Failed State : $recoveryDays days, $recoveryHours hours`n"
            $recoveryStats += "  Issue First Detected : $($Device.WorstDatasourceTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC`n"
            $recoveryStats += "  Resolution Verified  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (System Time)`n"
        }
        
        # Get current datasource status summary - M365 ONLY (filter out any non-M365 datasources)
        $statusSummary = ""
        if ($Device.DeviceSettings -and $Device.DataSources) {
            # Only show M365 datasources: D19 (Exchange), D20 (OneDrive), D5 (SharePoint), D23 (Teams)
            # Keep only D19, D20, D5/D05, D23 - remove all others (D01, D02, D10, etc.)
            $m365Only = ""
            if ($Device.DataSources -match 'D19') { $m365Only += "D19" }
            if ($Device.DataSources -match 'D20') { $m365Only += "D20" }
            if ($Device.DataSources -match 'D0?5') { $m365Only += "D5" }
            if ($Device.DataSources -match 'D23') { $m365Only += "D23" }
            
            if ($m365Only) {
                $statusSummary = "`nCurrent Backup Status:`n"
                $statusSummary += "  " + (Get-DatasourceStatus -DeviceSettings $Device.DeviceSettings -DataSourceString $m365Only) + "`n"
            }
        }
        
        $closeNote = @"
Issue Resolved - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)

The backup issue for $($Device.DeviceName) has been resolved.$recoveryStats$statusSummary
Last Timestamp       : $($Device.TimeStamp)
$(if ($oldestProblemRow) { "$oldestProblemRow`n" })
─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS
$datasourceDetails

─────────────────────────────────────────────────────────────────
View Device: $deviceUrl

This ticket was automatically closed by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)
"@
    } else {
        # Server/Workstation close note
        
        # Calculate Oldest Problem time based on WorstDatasourceTime (datasource-level)
        $oldestProblemTime = if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $worstHoursSince = ((Get-Date).ToUniversalTime() - $worstDatasourceTime).TotalHours
            Format-HoursAsRelativeTime $worstHoursSince
        } else { '' }
        
        # Build Oldest Problem row only if there's an actual problem (not "Unknown")
        $oldestProblemRow = if ($Device.WorstDatasourceName -and $Device.WorstDatasourceName -ne 'Unknown') {
            "Oldest Problem       : $($Device.WorstDatasourceTime) $oldestProblemTime | $($Device.WorstDatasourceName) | $($Device.WorstDatasourceStatusName)"
        } else { $null }
        
        # Get datasource details for close note
        $datasourceDetails = "  Datasource details not available"
        if ($Device.DeviceSettings) {
            $errorMsgs = if ($Device.ErrorMessages) { $Device.ErrorMessages } else { @{} }
            $datasourceDetails = Get-DatasourceDetails -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources -ErrorMessages $errorMsgs -AccountType $Device.AccountType
        }
        
        # Calculate recovery statistics
        $recoveryStats = ""
        if ($worstDatasourceTime -and $worstDatasourceTime -ne 'N/A' -and $worstDatasourceTime -is [DateTime]) {
            $failureStart = $worstDatasourceTime
            $recoveryTime = ((Get-Date).ToUniversalTime() - $failureStart).TotalHours
            $recoveryDays = [math]::Floor($recoveryTime / 24)
            $recoveryHours = [math]::Floor($recoveryTime % 24)
            $recoveryStats = "`nRecovery Statistics:`n"
            $recoveryStats += "  Time in Failed State : $recoveryDays days, $recoveryHours hours`n"
            $recoveryStats += "  Issue First Detected : $($worstDatasourceTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC`n"
            $recoveryStats += "  Resolution Verified  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (System Time)`n"
        }
        
        # Get current datasource status summary
        $statusSummary = ""
        if ($Device.DeviceSettings -and $Device.DataSources) {
            $statusSummary = "`nCurrent Backup Status:`n"
            $statusSummary += "  " + (Get-DatasourceStatus -DeviceSettings $Device.DeviceSettings -DataSourceString $Device.DataSources) + "`n"
        }
        
        $closeNote = @"
Issue Resolved - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)

The backup issue for $($Device.DeviceName) has been resolved.$recoveryStats$statusSummary
Last Timestamp       : $($Device.TimeStamp)
$(if ($oldestProblemRow) { "$oldestProblemRow`n" })
─────────────────────────────────────────────────────────────────
DATASOURCE DETAILS
$datasourceDetails

─────────────────────────────────────────────────────────────────
View Device: $deviceUrl

This ticket was automatically closed by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)
"@
    }

    try {
        if ($CloseResolvedTickets) {
            # Add closing note
            $noteParams = @{
                ticketId = $Ticket.id
                text = $closeNote
                detailDescriptionFlag = $true
                internalAnalysisFlag = $false
            }
            
            try {
                $noteResult = New-CWMTicketNote @noteParams -ErrorAction Stop
            }
            catch {
                # Check if ticket doesn't exist (404)
                if ($_.Exception.Message -match '404|not found|does not exist') {
                    Write-Host "  Skipping ticket #$($Ticket.id) - ticket no longer exists in ConnectWise" -ForegroundColor Gray
                    return "NotFound"  # Return special value to indicate ticket doesn't exist
                }
                throw  # Re-throw other errors
            }
            
            if (-not $noteResult -or -not $noteResult.id) {
                Write-Warning "Failed to add closing note to ticket #$($Ticket.id)"
                return $false
            }
            
            # Close the ticket
            $updateParams = @{
                id = $Ticket.id
                Operation = "replace"
                Path = "status/name"
                Value = $TicketClosedStatus
            }
            
            try {
                $updateResult = Update-CWMTicket @updateParams -ErrorAction Stop
            }
            catch {
                # Check if ticket doesn't exist (404)
                if ($_.Exception.Message -match '404|not found|does not exist') {
                    Write-Host "  Skipping ticket #$($Ticket.id) - ticket no longer exists in ConnectWise" -ForegroundColor Gray
                    return "NotFound"  # Return special value to indicate ticket doesn't exist
                }
                throw  # Re-throw other errors
            }
            
            if ($updateResult -and $updateResult.id) {
                $Script:ProcessedTickets.Closed += $Ticket.id
                $ticketUrl = Get-CWMTicketURL -TicketID $Ticket.id
                Write-Host "  Closed ticket #$($Ticket.id) for $($Device.DeviceName) - Issue Resolved" -ForegroundColor Green
                Write-Host "    URL: $ticketUrl" -ForegroundColor Gray
                return $true
            } else {
                Write-Warning "Failed to close ticket #$($Ticket.id) - status update failed"
                return $false
            }
        } else {
            Write-Host "  [TEST MODE] Would close ticket #$($Ticket.id) for $($Device.DeviceName)" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Warning "Error closing ticket #$($Ticket.id): $_"
        return $false
    }
}  ## Close a ConnectWise ticket when issue is resolved

Function Remove-CoveTicketsLastDay {
    <#
    .SYNOPSIS
        Deletes all Cove-created tickets from the last 24 hours
    .DESCRIPTION
        WARNING: This permanently deletes tickets and cannot be undone!
        Used for cleanup during testing/development.
    #>
    
    Write-Host "`n$Script:strLineSeparator" -ForegroundColor Red
    Write-Host "  WARNING: TICKET DELETION MODE ACTIVE" -ForegroundColor Red
    Write-Host "$Script:strLineSeparator" -ForegroundColor Red
    Write-Host ""
    Write-Host "  This will PERMANENTLY DELETE all tickets matching:" -ForegroundColor Yellow
    Write-Host "    - Summary contains: 'Cove Backup Alert'" -ForegroundColor White
    Write-Host "    - Board: $TicketBoard" -ForegroundColor White
    Write-Host "    - Created in last 24 hours" -ForegroundColor White
    Write-Host ""
    Write-Host "  THIS CANNOT BE UNDONE!" -ForegroundColor Red
    Write-Host ""
    
    $confirmation = Read-Host "  Type 'DELETE' (all caps) to proceed"
    
    if ($confirmation -ne 'DELETE') {
        Write-Host "`n  Cleanup cancelled by user." -ForegroundColor Cyan
        return
    }
    
    try {
        # Calculate date 24 hours ago
        $yesterday = (Get-Date).AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
        
        Write-Host "`n  Searching for Cove tickets created since $yesterday..." -ForegroundColor Cyan
        
        # Search for tickets matching criteria
        $conditions = "summary contains 'Cove Backup Alert' and board/name='$TicketBoard' and dateEntered > [$yesterday]"
        
        $tickets = Get-CWMTicket -Condition $conditions -all -ErrorAction Stop
        $Script:PerformanceMetrics.TicketSearchCount++
        
        if (-not $tickets -or $tickets.Count -eq 0) {
            Write-Host "  No matching tickets found." -ForegroundColor Green
            return
        }
        
        Write-Host "  Found $($tickets.Count) ticket(s) to delete`n" -ForegroundColor Yellow
        
        $deletedCount = 0
        $failedCount = 0
        
        foreach ($ticket in $tickets) {
            try {
                Write-Host "    Deleting ticket #$($ticket.id): $($ticket.summary)" -ForegroundColor Gray
                Remove-CWMTicket -id $ticket.id -ErrorAction Stop
                $deletedCount++
            }
            catch {
                Write-Warning "    Failed to delete ticket #$($ticket.id): $_"
                $failedCount++
            }
        }
        
        Write-Host "`n$Script:strLineSeparator" -ForegroundColor Cyan
        Write-Host "  Cleanup Summary:" -ForegroundColor Cyan
        Write-Host "    Deleted: $deletedCount" -ForegroundColor $(if ($deletedCount -gt 0) { 'Green' } else { 'Gray' })
        Write-Host "    Failed:  $failedCount" -ForegroundColor $(if ($failedCount -gt 0) { 'Red' } else { 'Gray' })
        Write-Host "$Script:strLineSeparator`n" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Error during ticket cleanup: $_"
    }
}  ## Delete Cove-created tickets from last 24 hours (WARNING: Permanent!)

#endregion ----- ConnectWise Manage Functions ----

#endregion ----- Functions ----

#region ----- Main Script Execution ----

# Check if cleanup mode is requested
if ($CleanupCoveTickets) {
    # Initialize ConnectWise Manage connection (but skip Cove authentication)
    $Script:CWMAPICredsFile = join-path -Path $True_Path -ChildPath "${env:computername}_${env:username}_CWM_Ticketing_Credentials.Secure.xml"
    $Script:CWMAPICredsPath = Split-path -path $CWMAPICredsFile
    
    if ($ClearCWMCredentials) {
        if (Test-Path $CWMAPICredsFile) {
            Remove-Item -Path $CWMAPICredsFile -Force
            Write-Host "`n  Removed stored ConnectWise Manage API credentials`n" -ForegroundColor Green
        }
    }
    
    InstallCWMPSModule
    AuthenticateCWM
    
    # Run cleanup and exit
    Remove-CoveTicketsLastDay
    exit
}

# Authenticate to Cove Data Protection
Send-APICredentialsCookie | Out-Null

if ($DebugCDP) { Write-Host "[DEBUG] Before partner lookup: Script:PartnerName='$Script:PartnerName', OriginalPartnerName='$Script:OriginalPartnerName'" -ForegroundColor Magenta }

# Get Partner Info - use Script:PartnerName directly (already set from parameter)
if ($Script:PartnerName) {
    if ($DebugCDP) { Write-Host "[DEBUG] Looking up partner: '$Script:PartnerName'" -ForegroundColor Magenta }
    # Use specified partner name
    Send-GetPartnerInfo $Script:PartnerName | Out-Null
} else {
    if ($DebugCDP) { Write-Host "[DEBUG] No PartnerName parameter - using authenticated partner from credentials: '$cred0'" -ForegroundColor Magenta }
    # Use authenticated partner from credentials
    Send-GetPartnerInfo $cred0 | Out-Null
}

if ($DebugCDP) { Write-Host "[DEBUG] After Send-GetPartnerInfo: Script:PartnerName='$Script:PartnerName', Script:PartnerId='$Script:PartnerId'" -ForegroundColor Magenta }

# Process each selected partner (or current partner)
if ($null -eq $Script:SelectedPartners) {
    $Script:SelectedPartners = @($Script:Partner.result.result)
}

# Initialize ConnectWise Manage connection
$Script:CWMAPICredsFile = join-path -Path $True_Path -ChildPath "${env:computername}_${env:username}_CWM_Ticketing_Credentials.Secure.xml"
$Script:CWMAPICredsPath = Split-path -path $CWMAPICredsFile

if ($ClearCWMCredentials) {
    if (Test-Path $CWMAPICredsFile) {
        Remove-Item -Path $CWMAPICredsFile -Force
        Write-Output "ConnectWise Manage Ticketing API Credentials file removed successfully."
    } else {
        Write-Warning "ConnectWise Manage Ticketing API Credentials file not found at $CWMAPICredsPath."
    }
}

InstallCWMPSModule
AuthenticateCWM

# Process all selected partners
foreach ($partner in $Script:SelectedPartners) {
    Write-Output "`n$Script:strLineSeparator"
    Write-Output "Processing Partner: $($partner.Name)"
    Write-Output $Script:strLineSeparator
    
    $Script:PartnerId = $partner.Id
    
    # Reset DeviceDetail for this partner
    $Script:DeviceDetail = @()
    
    # Get devices for this partner based on monitoring settings
    if ($MonitorSystems) {
        Write-Host "`n  === Monitoring Servers and Workstations ===" -ForegroundColor Cyan
        Send-GetDevices -AccountType 1 -StaleHours $StaleHoursServers -StaleHoursServers $StaleHoursServers -StaleHoursWorkstations $StaleHoursWorkstations -LookbackDays $DaysBack
    }
    
    if ($MonitorM365) {
        Write-Host "`n  === Monitoring M365 Tenants ===" -ForegroundColor Cyan
        Send-GetDevices -AccountType 2 -StaleHours $StaleHoursM365 -LookbackDays $DaysBack
    }
    
    # Overall summary for this partner
    Write-Host "`n  === Overall Summary ===" -ForegroundColor Cyan
    $DevicesWithIssues = $Script:DeviceDetail | Where-Object { $_.IssueSeverity -ne "Success" }
    $ServersCount = ($Script:DeviceDetail | Where-Object { $_.AccountType -eq 1 }).Count
    $M365Count = ($Script:DeviceDetail | Where-Object { $_.AccountType -eq 2 }).Count
    
    Write-Output "  Total Devices: $($Script:DeviceDetail.Count) (Servers/Workstations: $ServersCount, M365: $M365Count)"
    Write-Output "  Devices with Issues: $($DevicesWithIssues.Count)"
    
    # OPTIMIZATION v06: Initialize hierarchy lookup tracking (reset each run to prevent accumulation)
    $Script:HierarchyLookupStats = @{
        TotalLookups = 0
        CacheHits = 0
        CacheMisses = 0
        TotalTime = 0
        UniquePartnerIDs = @{}
    }
    
    # OPTIMIZATION v07: Removed batch CWM optimization (was failing with timeouts)
    # Using indexed lookups instead (faster and more reliable than failed batch fetches)
    
    Write-Host "  Processing devices with issues...`n" -ForegroundColor Cyan
    
    $deviceProcessingStartTime = Get-Date
    $totalDevices = $DevicesWithIssues.Count
    $currentDeviceNum = 0
    
    foreach ($device in $DevicesWithIssues) {
        $currentDeviceNum++
        $deviceStartTime = Get-Date
        
        $severityInfo = $Script:IssueSeverity[$device.IssueSeverity]
        Write-Host "  [$currentDeviceNum/$totalDevices] " -ForegroundColor Gray -NoNewline
        Write-Host "$($device.DeviceName)" -ForegroundColor $severityInfo.Color -NoNewline
        Write-Host " - $($device.IssueSeverity) - $($device.IssueDescription)" -ForegroundColor $severityInfo.Color
        
        # OPTIMIZATION v06: Resolve partner hierarchy ONLY for devices with issues (not all 750 devices)
        if (-not $device.EndCustomer) {
            $partnerHierarchyStartTime = Get-Date
            $partnerHierarchy = Get-PartnerHierarchyInfo -PartnerID $device.PartnerId -PartnerName $device.PartnerName
            $partnerHierarchyElapsed = ((Get-Date) - $partnerHierarchyStartTime).TotalMilliseconds
            
            # Update device object with hierarchy info (use Add-Member -Force to avoid property errors)
            $device.EndCustomer = $partnerHierarchy.EndCustomer
            $device.Site = $partnerHierarchy.Site
            $device | Add-Member -MemberType NoteProperty -Name "PartnerIdForHierarchy" -Value $partnerHierarchy.EndCustomerPartnerId -Force
            $device.EndCustomerLevel = $partnerHierarchy.EndCustomerLevel
            
            # Track hierarchy lookup performance
            $Script:HierarchyLookupStats.TotalLookups++
            $Script:HierarchyLookupStats.TotalTime += $partnerHierarchyElapsed
            
            if ($Script:PartnerHierarchyCache.ContainsKey($device.PartnerId)) {
                if (-not $Script:HierarchyLookupStats.UniquePartnerIDs.ContainsKey($device.PartnerId)) {
                    $Script:HierarchyLookupStats.CacheMisses++
                    $Script:HierarchyLookupStats.UniquePartnerIDs[$device.PartnerId] = 1
                } else {
                    $Script:HierarchyLookupStats.CacheHits++
                    $Script:HierarchyLookupStats.UniquePartnerIDs[$device.PartnerId]++
                }
            }
        }
        
        # Get CWM company using indexed lookup (faster than failed batch cache)
        $companyLookupStart = Get-Date
        $CWMCompany = Get-CWMCompanyForDevice -Device $device
        if ($debugCWM) { Write-Host "[DEBUG] CWM Company Result: [$($CWMCompany.ID)] $($CWMCompany.Name)" -ForegroundColor Green }
        $Script:PerformanceMetrics.CompanyLookupTime += ((Get-Date) - $companyLookupStart).TotalMilliseconds
        
        if ($CWMCompany) {
            # Search for existing ticket using indexed query
            $ticketSearchStart = Get-Date
            $existingTicket = Get-CWMTicketForDevice -DeviceName $device.DeviceName -CWMCompany $CWMCompany -IssueSeverity $device.IssueSeverity
            $Script:PerformanceMetrics.TicketSearchTime += ((Get-Date) - $ticketSearchStart).TotalMilliseconds
            
            if ($existingTicket) {
                # Update existing ticket
                $ticketUpdateStart = Get-Date
                $updateSuccess = Update-CWMTicketForDevice -Ticket $existingTicket -Device $device
                $Script:PerformanceMetrics.TicketUpdateTime += ((Get-Date) - $ticketUpdateStart).TotalMilliseconds
                
                # Add to retry queue if update failed
                if (-not $updateSuccess) {
                    $Script:FailedTicketOperations += [PSCustomObject]@{
                        Operation = 'Update'
                        TicketId = $existingTicket.id
                        Device = $device
                        CWMCompany = $CWMCompany
                        Ticket = $existingTicket
                        Timestamp = Get-Date
                        DeviceNumber = "[$currentDeviceNum/$totalDevices]"
                    }
                    Write-Host "  Failed to update ticket #$($existingTicket.id) - added to retry queue" -ForegroundColor Yellow
                }
                
                # Determine device type label
                $deviceTypeLabel = if ($device.AccountType -eq 2) {
                    "M365"
                } else {
                    $isServer = ($device.OSType -eq "2") -or ($device.OS -match "Server")
                    if ($isServer) { "Server" } else { "Workstation" }
                }
                
                # Determine device type label
                $deviceTypeLabel = if ($device.AccountType -eq 2) {
                    "M365"
                } else {
                    $isServer = ($device.OSType -eq "2") -or ($device.OS -match "Server")
                    if ($isServer) { "Server" } else { "Workstation" }
                }
                
                # NOTE: $abbreviatedIssue was already calculated earlier with proper error message truncation (lines 5925-5947)
                # and stored on device object, so use that instead of rebuilding
                $abbreviatedIssue = if ($device.AbbreviatedIssue) {
                    $device.AbbreviatedIssue
                } else {
                    # Extreme fallback if somehow not set (should never happen)
                    if ($device.IssueDescription -match "Backup failed") {
                        "Backup Failed"
                    } else {
                        "Issue"
                    }
                }
                
                $Script:TicketActions += [PSCustomObject]@{
                    Action = "Updated"
                    TicketID = $existingTicket.id
                    TicketSummary = if (("COVE $deviceTypeLabel $($device.DeviceName) #$($device.AccountId) - $abbreviatedIssue").Length -gt 100) { ("COVE $deviceTypeLabel $($device.DeviceName) #$($device.AccountId) - $abbreviatedIssue").Substring(0, 97) + "..." } else { "COVE $deviceTypeLabel $($device.DeviceName) #$($device.AccountId) - $abbreviatedIssue" }
                    DeviceName = $device.DeviceName
                    Company = $CWMCompany.Name
                    IssueSeverity = $device.IssueSeverity
                    IssueDescription = $device.IssueDescription
                }
            } else {
                # Create new ticket
                # Check if we already created a ticket for this device in this session
                $alreadyCreated = $Script:TicketActions | Where-Object { 
                    $_.Action -eq 'Created' -and 
                    $_.DeviceName -eq $device.DeviceName -and 
                    $_.Company -eq $CWMCompany.Name 
                }
                
                if ($alreadyCreated) {
                    if ($debugCDP) { 
                        Write-Host "[DEBUG] Skipping duplicate ticket creation for $($device.DeviceName) - already created ticket #$($alreadyCreated.TicketID)" -ForegroundColor Magenta 
                    }
                    continue  # Skip to next device
                }
                
                $ticketCreateStart = Get-Date
                $newTicket = New-CWMTicketForDevice -Device $device -CWMCompany $CWMCompany
                $Script:PerformanceMetrics.TicketCreateTime += ((Get-Date) - $ticketCreateStart).TotalMilliseconds
                
                # Add to retry queue if creation failed
                if (-not $newTicket) {
                    $Script:FailedTicketOperations += [PSCustomObject]@{
                        Operation = 'Create'
                        Device = $device
                        CWMCompany = $CWMCompany
                        Timestamp = Get-Date
                        DeviceNumber = "[$currentDeviceNum/$totalDevices]"
                    }
                    Write-Host "  Failed to create ticket for $($device.DeviceName) - added to retry queue" -ForegroundColor Yellow
                }
                
                if ($newTicket) {
                    # Determine device type label
                    $deviceTypeLabel = if ($device.AccountType -eq 2) {
                        "M365"
                    } else {
                        $isServer = ($device.OSType -eq "2")
                        if ($isServer) { "Server" } else { "Workstation" }
                    }
                    
                    # NOTE: $abbreviatedIssue was already calculated earlier with proper error message truncation (lines 5925-5947)
                    # Only rebuild if it's somehow empty (defensive programming)
                    # NOTE: $abbreviatedIssue was already calculated earlier and stored on device object
                    $abbreviatedIssue = if ($device.AbbreviatedIssue) {
                        $device.AbbreviatedIssue
                    } else {
                        # Extreme fallback if somehow not set (should never happen)
                        if ($device.IssueDescription -match "Backup failed") {
                            "Backup Failed"
                        } else {
                            "Issue"
                        }
                    }
                    
                    # Build ticket summary and truncate if needed (Option D format)
                    $fullSummary = "COVE $deviceTypeLabel $($device.DeviceName) #$($device.AccountId) - $abbreviatedIssue"
                    $ticketSummary = if ($fullSummary.Length -gt 100) {
                        $fullSummary.Substring(0, 97) + "..."
                    } else {
                        $fullSummary
                    }
                    
                    $Script:TicketActions += [PSCustomObject]@{
                        Action = "Created"
                        TicketID = $newTicket.id
                        TicketSummary = $ticketSummary
                        DeviceName = $device.DeviceName
                        Company = $CWMCompany.Name
                        IssueSeverity = $device.IssueSeverity
                        IssueDescription = $device.IssueDescription
                    }
                }
            }
        } else {
            Write-Warning "  Skipping ticket creation - no matching CWM company for $($device.DeviceName)"
            
            # Export ticket body to timestamped subfolder for review
            $exportFileName = "TicketBody_$($device.DeviceName -replace '[^a-zA-Z0-9_-]','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $exportFilePath = Join-Path $Script:TicketExportFolder $exportFileName
            
            # Build ticket body
            if ($device.AccountType -eq 2) {
                # M365: Use cloud-properties structure
                $deviceUrl = "https://backup.management/#/device/$($device.AccountID)/cloud-properties/office365/history"
            } else {
                # Standard backup devices: Use backup/overview view
                $viewId = '1662'
                $deviceUrl = "https://backup.management/#/backup/overview/view/$viewId(panel:device-properties/$($device.AccountID)/summary)"
            }
            $readableDataSources = Convert-DataSourceCodes -DataSourceString $device.DataSources
            $deviceType = if ($device.AccountType -eq 2) { "Microsoft 365 Tenant" } else { "Server or Workstation" }
            $datasourceDetails = "  Datasource details not available (device detail lookup required)"
            if ($device.DeviceSettings) {
                $errorMsgs = if ($device.ErrorMessages) { $device.ErrorMessages } else { @{} }
                $datasourceDetails = Get-DatasourceDetails -DeviceSettings $device.DeviceSettings -DataSourceString $device.DataSources -ErrorMessages $errorMsgs -AccountType $device.AccountType
            }
            $partnerDisplay = if ($device.Site) {
                "$($device.EndCustomer) | Site: $($device.Site)"
            } else {
                $device.EndCustomer
            }
            $isServer = ($device.OSType -eq "2") -or ($device.OS -match "Server")
            $deviceTypeDisplay = switch ($device.Physicality) {
                "Virtual" { if ($isServer) { "Virtual Server" } else { "Virtual Workstation" } }
                "Physical" { if ($isServer) { "Physical Server" } else { "Physical Workstation" } }
                default { "Server or Workstation" }
            }
            
            $ticketBody = @"
========================================
TICKET BODY EXPORT - COMPANY NOT FOUND
========================================

ConnectWise Company Match Attempted For:
  EndCustomer: $($device.EndCustomer)
  PartnerName: $($device.PartnerName)
  Reference: $($device.Reference)
  Site: $($device.Site)

Result: NO MATCHING COMPANY FOUND IN CONNECTWISE

========================================
TICKET DETAILS THAT WOULD BE CREATED:
========================================

Cove Data Protection Backup Alert

Device: $($device.DeviceName)
Computer Name: $($device.ComputerName)
$(if ($device.DeviceAlias) { "Alias: $($device.DeviceAlias)`n" })Customer: $partnerDisplay
Reference: $($device.Reference)
Issue Severity: $($device.IssueSeverity)
Issue Description: $($device.IssueDescription)

View Device in Cove Portal:
$deviceUrl

Last Timestamp: $($device.TimeStamp)

Device Details:

$(if ($device.AccountType -eq 2) {
    # M365 tenant details
"Storage Usage:
  $(Format-AlignedLabel 'Selected Data' 20)$($device.SelectedGB) GB
  $(Format-AlignedLabel 'Used Storage' 20)$($device.UsedGB) GB
  $(Format-AlignedLabel 'Storage Location' 20)$($device.Location)"
} else {
    # Server/Workstation details
"Hardware Information:
  $(Format-AlignedLabel 'OS' 20)$($device.OS)
  $(Format-AlignedLabel 'Manufacturer' 20)$($device.Manufacturer)
  $(Format-AlignedLabel 'Model' 20)$($device.Model)
  $(Format-AlignedLabel 'Device Type' 20)$deviceTypeDisplay
  $(Format-AlignedLabel 'IP Address' 20)$($device.IPAddress)
  $(Format-AlignedLabel 'External IP' 20)$($device.ExternalIP)

Backup Configuration:
  $(Format-AlignedLabel 'Backup Profile' 20)$($device.Profile) (ID: $($device.ProfileID))
  $(Format-AlignedLabel 'Retention Policy' 20)$($device.Product) (ID: $($device.ProductID))
  $(Format-AlignedLabel 'Storage Location' 20)$($device.Location)
  $(Format-AlignedLabel 'Timezone Offset' 20)$($device.TimezoneOffset)"
})

Datasource Details:
$datasourceDetails

This ticket was automatically created by Cove Data Protection Monitoring $($Script:ScriptVersion) @ $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") (System Time)

========================================
"@
            
            try {
                $ticketBody | Set-Content -Path $exportFilePath -Encoding UTF8
                Write-Host "    Exported ticket body to: $exportFilePath" -ForegroundColor Cyan
            }
            catch {
                Write-Warning "    Failed to export ticket body: $($_.Exception.Message)"
            }
        }
        
        $Script:AllIssues += $device
        
        # Add blank line between devices for readability
        Write-Host ""
    }
    
    # FINAL RETRY QUEUE: Only contains operations that failed AFTER immediate retry
    # These are likely persistent issues, so we limit attempts and add delays
    if ($Script:FailedTicketOperations.Count -gt 0) {
        Write-Host "`n  ========================================" -ForegroundColor Yellow
        Write-Host "  FINAL RETRY ATTEMPT FOR FAILED OPERATIONS" -ForegroundColor Yellow
        Write-Host "  ========================================" -ForegroundColor Yellow
        Write-Host "  Found $($Script:FailedTicketOperations.Count) operation(s) that failed immediate retry`n" -ForegroundColor Yellow
        
        $retrySuccess = 0
        $retryFailed = 0
        $retrySkipped = 0
        $maxFinalRetries = 1  # Limit final retries (these already failed immediate retry)
        
        foreach ($failedOp in $Script:FailedTicketOperations) {
            # Skip if this operation has already been retried too many times
            if (-not $failedOp.RetryCount) { $failedOp.RetryCount = 0 }
            
            if ($failedOp.RetryCount -ge $maxFinalRetries) {
                Write-Host "  $($failedOp.DeviceNumber) Skipping $($failedOp.Device.DeviceName) - max retries ($maxFinalRetries) reached" -ForegroundColor Gray
                $retrySkipped++
                continue
            }
            
            $failedOp.RetryCount++
            Write-Host "  $($failedOp.DeviceNumber) Final retry for $($failedOp.Device.DeviceName)..." -ForegroundColor Cyan
            
            # Wait 5 seconds before final retry (API might be recovering)
            Write-Host "    ⏱️  Waiting 5 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
            
            try {
                if ($failedOp.Operation -eq 'Update') {
                    # Retry ticket update
                    $updateSuccess = Update-CWMTicketForDevice -Ticket $failedOp.Ticket -Device $failedOp.Device
                    
                    if ($updateSuccess) {
                        Write-Host "    ✓ Successfully updated ticket #$($failedOp.TicketId)" -ForegroundColor Green
                        $retrySuccess++
                    } else {
                        Write-Host "    Still failing for ticket #$($failedOp.TicketId)" -ForegroundColor Red
                        $retryFailed++
                    }
                }
                elseif ($failedOp.Operation -eq 'Create') {
                    # Retry ticket creation
                    $newTicket = New-CWMTicketForDevice -Device $failedOp.Device -CWMCompany $failedOp.CWMCompany
                    
                    if ($newTicket) {
                        Write-Host "    ✓ Successfully created ticket #$($newTicket.id)" -ForegroundColor Green
                        $retrySuccess++
                    } else {
                        Write-Host "    Still failing for ticket creation" -ForegroundColor Red
                        $retryFailed++
                    }
                }
            }
            catch {
                Write-Host "    Retry error: $($_.Exception.Message)" -ForegroundColor Red
                $retryFailed++
            }
        }
        
        Write-Host "`n  Final Retry Summary:" -ForegroundColor Cyan
        Write-Host "    Successful : $retrySuccess" -ForegroundColor $(if ($retrySuccess -gt 0) { "Green" } else { "Gray" })
        Write-Host "    Failed     : $retryFailed" -ForegroundColor $(if ($retryFailed -gt 0) { "Red" } else { "Gray" })
        Write-Host "    Skipped    : $retrySkipped (max retries reached)" -ForegroundColor Gray
        
        if ($retryFailed -gt 0) {
            Write-Host "`n  $retryFailed operation(s) could not be completed after multiple retries" -ForegroundColor Yellow
            Write-Host "  Check ConnectWise API connectivity and review error messages above" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # OPTIMIZED: Check for resolved issues using reverse logic (query tickets first, not devices)
    # This is vastly more efficient: queries ~10-100 tickets instead of checking 1000+ successful devices
    Write-Host "`n  Checking for tickets to close (reverse logic - querying open tickets)..." -ForegroundColor Cyan
    
    # Query all open Cove backup tickets from ConnectWise
    $openCoveTickets = @()
    try {
        $ticketFilter = "board/name='$TicketBoard' and closedFlag=false and summary like 'Cove%'"
        $openCoveTickets = Get-CWMTicket -condition $ticketFilter -all
        $Script:PerformanceMetrics.TicketSearchCount++
        
        if ($openCoveTickets) {
            Write-Host "  Found $($openCoveTickets.Count) open Cove backup tickets" -ForegroundColor Cyan
        } else {
            Write-Host "  No open Cove backup tickets found" -ForegroundColor Gray
        }
    }
    catch {
        Write-Warning "Failed to query open tickets from ConnectWise: $($_.Exception.Message)"
    }
    
    # Extract device IDs from tickets and build targeted Cove query
    $deviceIdsFromTickets = @()
    $ticketDeviceMap = @{}  # Map AccountId → Ticket for efficient lookup
    
    foreach ($ticket in $openCoveTickets) {
        # Extract device ID from Option D format: "COVE Type devicename #1234567 - ..."
        # Also handle old format: "Cove: [Type] DeviceName (ID: 1234567) - Severity - Issue"
        if ($ticket.summary -match '#(\d+)\s+-' -or $ticket.summary -match '\(ID:\s*(\d+)\)') {
            $deviceId = $matches[1]
            $deviceIdsFromTickets += $deviceId
            $ticketDeviceMap[$deviceId] = $ticket
            
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id): Extracted Device ID $deviceId" -ForegroundColor Magenta
            }
        }
    }
    
    # Deduplicate device IDs (multiple tickets may reference same device)
    $totalTicketsWithDeviceIds = $deviceIdsFromTickets.Count
    $deviceIdsFromTickets = @($deviceIdsFromTickets | Sort-Object -Unique)
    
    if ($debugCDP -and $totalTicketsWithDeviceIds -gt 0) {
        Write-Host "[DEBUG] Extracted device IDs: $totalTicketsWithDeviceIds total, $($deviceIdsFromTickets.Count) unique" -ForegroundColor Magenta
        $sampleIds = $deviceIdsFromTickets | Select-Object -First 5
        Write-Host "[DEBUG] Sample device IDs: $($sampleIds -join ', ')..." -ForegroundColor Magenta
    }
    
    # Query Cove for only the devices that have open tickets
    # Using top-level partner returns all nested devices
    $ticketDeviceStatuses = @()
    if ($deviceIdsFromTickets.Count -gt 0) {
        Write-Host "  Querying Cove for $($deviceIdsFromTickets.Count) devices with open tickets..." -ForegroundColor Cyan
        
        # Build filter: AU in (id1,id2,id3,...)
        $deviceIdFilter = "AU in (" + ($deviceIdsFromTickets -join ',') + ")"
        
        if ($debugCDP) {
            $filterPreview = if ($deviceIdFilter.Length -gt 100) { $deviceIdFilter.Substring(0, 100) + "..." } else { $deviceIdFilter }
            Write-Host "[DEBUG] Cove Device ID Filter (preview): $filterPreview" -ForegroundColor Magenta
            Write-Host "[DEBUG] Filter contains $($deviceIdsFromTickets.Count) unique device IDs" -ForegroundColor Magenta
        }
        
        try {
            # Use direct API call with custom filter
            # Query using the PartnerName parameter's partner to get all nested devices
            $url = "https://api.backup.management/jsonapi"
            $data = @{}
            $data.jsonrpc = '2.0'
            $data.id = '2'
            $data.visa = $visa
            $data.method = 'EnumerateAccountStatistics'
            $data.params = @{}
            $data.params.query = @{}
            $data.params.query.PartnerId = [int]$Script:PartnerId
            $data.params.query.Filter = $deviceIdFilter
            # Request all columns needed for datasource details in closure notes
            $data.params.query.Columns = @('AU','AN','TL','T0','I78','AT','OT','TZ',
                'F0','F3','F4','F5','F7','FL','FO','FA','FJ','FQ',  # Files
                'S0','S3','S4','S5','S7','SL','SO','SA','SJ','SQ',  # System State
                'G0','G3','G4','G5','G7','GL','GO','GA','GJ','GQ','GM','G@',  # M365 Exchange
                'J0','J3','J4','J5','J7','JL','JO','JA','JJ','JQ','JM',  # M365 OneDrive
                'D5F0','D5F3','D5F4','D5F5','D5F6','D5F9','D5F16','D5F17','D5F18','D5F20','D5F22',  # M365 SharePoint
                'D23F0','D23F3','D23F4','D23F5','D23F6','D23F9','D23F16','D23F17','D23F18','D23F20','D23F23','D23F24','D23F25')  # M365 Teams
            $data.params.query.RecordsCount = $deviceIdsFromTickets.Count  # Request all devices, not default of 1
            
            if ($DebugCDP) { 
                Write-Host "[DEBUG] Query PartnerId: $($Script:PartnerId) (from Script:PartnerName='$Script:PartnerName')" -ForegroundColor Magenta 
            }
            
            $jsondata = $data | ConvertTo-Json -Depth 10
            $params = @{
                Uri = $url
                Method = 'POST'
                Headers = @{ 'Authorization' = "Bearer $Script:visa" }
                Body = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
                ContentType = 'application/json; charset=utf-8'
                WebSession = $websession
                UseBasicParsing = $true
            }
            
            $response = Invoke-RestMethod @params
            
            if ($response.error) {
                throw "API Error: $($response.error.message)"
            }
            
            $ticketDeviceStatuses = @($response.result.result)
            Write-Host "  Retrieved status for $($ticketDeviceStatuses.Count) devices from Cove" -ForegroundColor Green
            
            if ($debugCDP) {
                Write-Host "[DEBUG] Expected $($deviceIdsFromTickets.Count) devices, got $($ticketDeviceStatuses.Count)" -ForegroundColor Magenta
                if ($ticketDeviceStatuses.Count -gt 0) {
                    $retrievedIds = $ticketDeviceStatuses | ForEach-Object { $_.Settings.AU -join '' }
                    Write-Host "[DEBUG] Retrieved device IDs (sample): $($retrievedIds | Select-Object -First 5 | ForEach-Object { $_ } | Join-String -Separator ', ')..." -ForegroundColor Magenta
                }
            }
        }
        catch {
            Write-Warning "Failed to query Cove devices: $($_.Exception.Message)"
        }
    }
    
    # Check each open ticket to see if the device is now successful
    $ticketsEvaluated = 0
    $ticketsClosed = 0
    $ticketsSkipped_ParseFailed = 0
    $ticketsSkipped_NotInCove = 0
    $ticketsSkipped_NotFound = 0  # NEW: Tickets that don't exist in ConnectWise anymore
    $ticketsLeftOpen_StillHasIssues = 0
    $ticketsLeftOpen_InProcess = 0
    $ticketsSkipped_JustCreatedOrUpdated = 0  # NEW: Track tickets we just processed
    
    # Get list of ticket IDs we just created/updated to avoid closing them immediately
    $justProcessedTicketIds = @($Script:TicketActions | Where-Object { $_.Action -in @('Created','Updated') } | ForEach-Object { $_.TicketID })
    
    foreach ($ticket in $openCoveTickets) {
        $ticketsEvaluated++
        
        # Skip tickets we just created/updated in this same run
        # (They were created for a reason - don't immediately close them)
        if ($ticket.id -in $justProcessedTicketIds) {
            $ticketsSkipped_JustCreatedOrUpdated++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id): Just created/updated in this run - skipping closure check" -ForegroundColor Cyan
            }
            continue
        }
        
        # Extract device ID from ticket summary
        $deviceId = $null
        # Try Option D format first: "COVE Type devicename #1234567 - ..."
        if ($ticket.summary -match '#(\d+)\s+-') {
            $deviceId = $matches[1]
        }
        # Fallback to old format: "Cove: [Type] DeviceName (ID: 1234567) - ..."
        elseif ($ticket.summary -match '\(ID:\s*(\d+)\)') {
            $deviceId = $matches[1]
        }
        
        if (-not $deviceId) {
            $ticketsSkipped_ParseFailed++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id): Could not parse device ID from: $($ticket.summary)" -ForegroundColor Yellow
            }
            continue
        }
        
        # Find this device in our Cove query results
        $coveDevice = $ticketDeviceStatuses | Where-Object { ($_.Settings.AU -join '') -eq $deviceId }
        
        if (-not $coveDevice) {
            # Device not found in Cove - may have been deleted or archived
            $ticketsSkipped_NotInCove++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id): Device ID $deviceId not found in Cove query" -ForegroundColor Yellow
            }
            continue
        }
        
        # Get device details
        $deviceName = $coveDevice.Settings.AN -join ''
        $lastSessionStatus = $coveDevice.Settings.T0 -join ''
        $accountType = $coveDevice.Settings.AT -join ''  # 1=Systems, 2=M365
        
        # Check if we should be processing this device type
        if ($accountType -eq '1' -and -not $MonitorSystems) {
            # This is a system device but we're not monitoring systems - skip
            $ticketsLeftOpen_StillHasIssues++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName' - System device but MonitorSystems=\$false - skipping" -ForegroundColor Yellow
            }
            continue
        }
        if ($accountType -eq '2' -and -not $MonitorM365) {
            # This is an M365 device but we're not monitoring M365 - skip
            $ticketsLeftOpen_StillHasIssues++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName' - M365 device but MonitorM365=\$false - skipping" -ForegroundColor Yellow
            }
            continue
        }
        
        # Get last successful backup time (TL = Unix timestamp)
        $lastBackupUnix = $coveDevice.Settings.TL -join ''
        $lastBackupTime = $null
        $isStale = $false
        
        if ($lastBackupUnix -and $lastBackupUnix -ne '0') {
            try {
                $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
                $lastBackupTime = $epoch.ToUniversalTime().AddSeconds([int64]$lastBackupUnix)
                $hoursSinceBackup = ((Get-Date).ToUniversalTime() - $lastBackupTime).TotalHours
                
                # Check if backup is stale (older than threshold) - use proper server/workstation/M365 thresholds
                $osType = $coveDevice.Settings.OT -join ''
                $staleThreshold = if ($accountType -eq '2') { 
                    $StaleHoursM365 
                } elseif ($osType -eq '2') { 
                    $StaleHoursServers 
                } else { 
                    $StaleHoursWorkstations 
                }
                $isStale = ($hoursSinceBackup -gt $staleThreshold)
                
                if ($debugCDP) {
                    Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName': Last backup $([Math]::Round($hoursSinceBackup,1))h ago, Threshold: $staleThreshold h, Stale: $isStale" -ForegroundColor Magenta
                }
            } catch {
                if ($debugCDP) {
                    Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName': Failed to parse timestamp $lastBackupUnix" -ForegroundColor Yellow
                }
            }
        }
        
        # Check if device is now successful (status 5 = Completed) AND not stale
        if ($lastSessionStatus -eq '5' -and -not $isStale) {
            Write-Host "  Closing ticket #$($ticket.id) for '$deviceName' (ID: $deviceId) - backup successful" -ForegroundColor Green
            
            # Build device object for Close-CWMTicketForDevice with full settings for datasource details
            $deviceObj = [PSCustomObject]@{
                DeviceName = $deviceName
                AccountId = $deviceId
                IssueSeverity = "Success"
                DeviceSettings = $coveDevice.Settings  # Full device settings for Get-DatasourceDetails
                DataSources = ($coveDevice.Settings.I78 -join '')  # Datasource codes
                AccountType = [int]$accountType  # 1=Systems, 2=M365
                ErrorMessages = @{}  # No errors on successful devices
            }
            
            $ticketCloseStart = Get-Date
            $closeResult = Close-CWMTicketForDevice -Ticket $ticket -Device $deviceObj
            $Script:PerformanceMetrics.TicketCloseTime += ((Get-Date) - $ticketCloseStart).TotalMilliseconds
            if ($closeResult -eq "NotFound") {
                $ticketsSkipped_NotFound++
            } elseif ($closeResult) {
                $ticketsClosed++
            }
            
            $Script:TicketActions += [PSCustomObject]@{
                Action = "Closed"
                TicketID = $ticket.id
                TicketSummary = $ticket.summary
                DeviceName = $deviceName
                Company = $ticket.company.name
                IssueSeverity = "Resolved"
                IssueDescription = "Issue resolved - backup successful"
            }
        } elseif ($lastSessionStatus -eq '5' -and $isStale) {
            # Status shows completed but backup is stale - leave ticket open
            $ticketsLeftOpen_StillHasIssues++
            if ($debugCDP) {
                $hoursSinceBackup = if ($lastBackupTime) { [Math]::Round(((Get-Date).ToUniversalTime() - $lastBackupTime).TotalHours, 1) } else { "N/A" }
                Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName' - Backup stale ($hoursSinceBackup hours old), leaving open" -ForegroundColor Yellow
            }
        } elseif ($lastSessionStatus -eq '1') {
            # Status 1 = In Process
            $ticketsLeftOpen_InProcess++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName' - Backup in process, leaving open" -ForegroundColor Yellow
            }
        } else {
            # Device still has issues - leave ticket open
            $ticketsLeftOpen_StillHasIssues++
            if ($debugCDP) {
                Write-Host "[DEBUG] Ticket #$($ticket.id) for '$deviceName' - Still has issues (Status: $lastSessionStatus), leaving open" -ForegroundColor Yellow
            }
        }
    }
    
    # Calculate consistent ticket closure count from TicketActions
    $ticketsClosedCount = ($Script:TicketActions | Where-Object {$_.Action -eq 'Closed'}).Count
    
    # Display detailed closure statistics
    Write-Host "`n  Ticket Closure Summary:" -ForegroundColor Cyan
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Gray
    Write-Host ("    Total Open Tickets Found           : {0,5}" -f $ticketsEvaluated) -ForegroundColor White
    Write-Host ("    Tickets Closed (Issue Resolved)    : {0,5}" -f $ticketsClosedCount) -ForegroundColor Green
    Write-Host ("    Tickets Left Open (Backup Running) : {0,5}" -f $ticketsLeftOpen_InProcess) -ForegroundColor Yellow
    Write-Host ("    Tickets Left Open (Still Failing)  : {0,5}" -f $ticketsLeftOpen_StillHasIssues) -ForegroundColor Yellow
    Write-Host ("    Tickets Skipped (Just Created/Upd) : {0,5}" -f $ticketsSkipped_JustCreatedOrUpdated) -ForegroundColor Cyan
    Write-Host ("    Tickets Skipped (Not In Cove)      : {0,5}" -f $ticketsSkipped_NotInCove) -ForegroundColor Gray
    Write-Host ("    Tickets Skipped (Not Found in CWM) : {0,5}" -f $ticketsSkipped_NotFound) -ForegroundColor Gray
    Write-Host ("    Tickets Skipped (Parse Failed)     : {0,5}" -f $ticketsSkipped_ParseFailed) -ForegroundColor Gray
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Gray
}

# Export results
if ($Script:AllIssues.Count -gt 0) {
    $Script:AllIssues | Export-Csv -Path $Script:LogFile -NoTypeInformation
    Write-Output "`n$Script:strLineSeparator"
    Write-Output "  Issues exported to: $Script:LogFile"
}

if ($Script:TicketActions.Count -gt 0) {
    $Script:TicketActions | Export-Csv -Path $Script:TicketLogFile -NoTypeInformation
    Write-Output "  Ticket actions exported to: $Script:TicketLogFile"
}

# Display note about Cove Advanced Filter
if ($Script:AllIssues.Count -gt 0) {
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                  COVE ADVANCED FILTER                         ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host "`n  To view these devices in Cove Data Protection:" -ForegroundColor White
    Write-Host "  1. Navigate to Devices → Advanced Filter" -ForegroundColor Gray
    Write-Host "  2. Use the API filter displayed at the beginning of this script" -ForegroundColor Gray
    Write-Host "  3. The filter checks only configured datasources (I78) to avoid false positives" -ForegroundColor Gray
    Write-Host ""
}

# Calculate script execution time first (needed throughout summary)
$Script:ScriptEndTime = Get-Date
$Script:ElapsedTime = $Script:ScriptEndTime - $Script:ScriptStartTime

# Calculate summary statistics
$ticketsCreated = ($Script:TicketActions | Where-Object {$_.Action -eq 'Created'}).Count
$ticketsUpdated = ($Script:TicketActions | Where-Object {$_.Action -eq 'Updated'}).Count
$ticketsClosed = ($Script:TicketActions | Where-Object {$_.Action -eq 'Closed'}).Count
$companiesCreated = if ($Script:CompaniesCreatedViaHelper) { $Script:CompaniesCreatedViaHelper.Count } else { 0 }
$referencesUpdated = if ($Script:ReferencesUpdated) { $Script:ReferencesUpdated.Count } else { 0 }

#region ----- Summary Output ----
Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                   TICKET ACTIONS SUMMARY                      ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ("  Total Issues Found    : {0,4}" -f $Script:AllIssues.Count) -ForegroundColor White
Write-Host ("  Tickets Created       : {0,4}" -f $ticketsCreated) -ForegroundColor $(if ($ticketsCreated -gt 0) { 'Green' } else { 'Gray' })
Write-Host ("  Tickets Updated       : {0,4}" -f $ticketsUpdated) -ForegroundColor $(if ($ticketsUpdated -gt 0) { 'Yellow' } else { 'Gray' })
Write-Host ("  Tickets Closed        : {0,4}" -f $ticketsClosed) -ForegroundColor $(if ($ticketsClosed -gt 0) { 'Magenta' } else { 'Gray' })
Write-Host ("  Companies Created     : {0,4}" -f $companiesCreated) -ForegroundColor $(if ($companiesCreated -gt 0) { 'Cyan' } else { 'Gray' })
Write-Host ("  References Updated    : {0,4}" -f $referencesUpdated) -ForegroundColor $(if ($referencesUpdated -gt 0) { 'Blue' } else { 'Gray' })

Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                 SCRIPT EXECUTION SUMMARY                      ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "  Script Version        : $($Script:ScriptVersion)" -ForegroundColor White
Write-Host "  Elapsed Time          : $($Script:ElapsedTime.ToString('hh\:mm\:ss'))" -ForegroundColor Green
Write-Host "  Start Time            : $($Script:ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
Write-Host "  End Time              : $($Script:ScriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray

# Performance breakdown (optional detail)
if ($Script:HierarchyLookupStats) {
    Write-Host "`n╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              DETAILED PERFORMANCE BREAKDOWN                   ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    # Calculate totals first for summary
    $totalTicketOperationTime = $Script:PerformanceMetrics.TicketCreateTime + $Script:PerformanceMetrics.TicketUpdateTime + $Script:PerformanceMetrics.TicketCloseTime
    $totalAPITime = $Script:PerformanceMetrics.DeviceQueryTime + $Script:HierarchyLookupStats.TotalTime + `
                    $Script:PerformanceMetrics.CompanyLookupTime + $Script:PerformanceMetrics.TicketSearchTime + `
                    $totalTicketOperationTime
    $totalAPISeconds = [math]::Round($totalAPITime / 1000, 2)
    $otherProcessing = [math]::Round($Script:ElapsedTime.TotalSeconds - $totalAPISeconds, 2)
    $apiPercent = [math]::Round(($totalAPISeconds / $Script:ElapsedTime.TotalSeconds) * 100, 1)
    
    # Summary section FIRST (most important info)
    Write-Host "`n  ━━━━━ Summary ━━━━━" -ForegroundColor Yellow
    Write-Host ("    Total API Time       : {0,8:N2} seconds ({1:N1} minutes)" -f $totalAPISeconds, ($totalAPISeconds / 60)) -ForegroundColor Cyan
    Write-Host ("    Other Processing     : {0,8:N2} seconds" -f $otherProcessing) -ForegroundColor Gray
    Write-Host ("    Percentage API Time  : {0,7:N1}%" -f $apiPercent) -ForegroundColor $(if ($apiPercent -gt 75) { 'Yellow' } else { 'Green' })
    
    # Cove API
    Write-Host "`n  ━━━━━ Cove API Performance ━━━━━" -ForegroundColor Yellow
    Write-Host ("    Device Statistics Query : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.DeviceQueryTime / 1000)) -ForegroundColor White
    
    # Partner Hierarchy Cache
    if ($Script:HierarchyLookupStats.TotalLookups -gt 0) {
        $hierarchyCacheHitRate = [math]::Round(($Script:HierarchyLookupStats.CacheHits / $Script:HierarchyLookupStats.TotalLookups) * 100, 1)
        Write-Host "`n  ━━━━━ Partner Hierarchy Performance ━━━━━" -ForegroundColor Yellow
        Write-Host ("    Total Lookups            : {0,2}" -f $Script:HierarchyLookupStats.TotalLookups) -ForegroundColor White
        Write-Host ("    Cache Hits               : {0,2}" -f $Script:HierarchyLookupStats.CacheHits) -ForegroundColor Green
        Write-Host ("    Cache Misses (API calls) : {0,2}" -f $Script:HierarchyLookupStats.CacheMisses) -ForegroundColor Yellow
        Write-Host ("    Unique Partners          : {0,2}" -f $Script:HierarchyLookupStats.UniquePartnerIDs.Count) -ForegroundColor White
        Write-Host ("    Total Time               : {0,6:N2} seconds" -f ($Script:HierarchyLookupStats.TotalTime / 1000)) -ForegroundColor White
        Write-Host ("    Average per Lookup       : {0,6:N2} ms" -f ($Script:HierarchyLookupStats.TotalTime / $Script:HierarchyLookupStats.TotalLookups)) -ForegroundColor Gray
        Write-Host ("    Cache Hit Rate           : {0,5:N1}%" -f $hierarchyCacheHitRate) -ForegroundColor $(if ($hierarchyCacheHitRate -gt 50) { "Green" } else { "Yellow" })
    }
    
    # EndCustomer → CWM Company Cache
    if ($Script:EndCustomerCacheStats.TotalLookups -gt 0) {
        $endCustomerCacheHitRate = [math]::Round(($Script:EndCustomerCacheStats.CacheHits / $Script:EndCustomerCacheStats.TotalLookups) * 100, 1)
        Write-Host "`n  ━━━━━ EndCustomer → CWM Company Cache Performance ━━━━━" -ForegroundColor Yellow
        Write-Host ("    Total Lookups            : {0,2}" -f $Script:EndCustomerCacheStats.TotalLookups) -ForegroundColor White
        Write-Host ("    Cache Hits               : {0,2}" -f $Script:EndCustomerCacheStats.CacheHits) -ForegroundColor Green
        Write-Host ("    Cache Misses (CWM calls) : {0,2}" -f $Script:EndCustomerCacheStats.CacheMisses) -ForegroundColor Yellow
        Write-Host ("    Unique Customers Cached  : {0,2}" -f $Script:EndCustomerToCWMCompanyCache.Count) -ForegroundColor White
        Write-Host ("    Cache Hit Rate           : {0,5:N1}%" -f $endCustomerCacheHitRate) -ForegroundColor $(if ($endCustomerCacheHitRate -gt 50) { "Green" } else { "Yellow" })
    }
    
    # ConnectWise API
    Write-Host "`n  ━━━━━ ConnectWise API Performance ━━━━━" -ForegroundColor Yellow
    $totalCompanyLookups = ($DevicesWithIssues | Measure-Object).Count
    $totalTicketSearches = ($DevicesWithIssues | Measure-Object).Count
    
    Write-Host ("    Company Lookups      : {0,2}" -f $totalCompanyLookups) -ForegroundColor White
    Write-Host ("    Total Lookup Time    : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.CompanyLookupTime / 1000)) -ForegroundColor White
    if ($totalCompanyLookups -gt 0) {
        Write-Host ("    Avg per Lookup       : {0,6:N2} ms" -f ($Script:PerformanceMetrics.CompanyLookupTime / $totalCompanyLookups)) -ForegroundColor Gray
    }
    
    Write-Host ("    Ticket Searches      : {0,2}" -f $totalTicketSearches) -ForegroundColor White
    Write-Host ("    Total Search Time    : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.TicketSearchTime / 1000)) -ForegroundColor White
    if ($totalTicketSearches -gt 0) {
        Write-Host ("    Avg per Search       : {0,6:N2} ms" -f ($Script:PerformanceMetrics.TicketSearchTime / $totalTicketSearches)) -ForegroundColor Gray
    }
    
    # Ticket Operations
    Write-Host "`n  ━━━━━ Ticket Operations Performance ━━━━━" -ForegroundColor Yellow
    Write-Host ("    Ticket Creation Time     : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.TicketCreateTime / 1000)) -ForegroundColor White
    Write-Host ("    Ticket Update Time       : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.TicketUpdateTime / 1000)) -ForegroundColor White
    Write-Host ("    Ticket Closure Time      : {0,6:N2} seconds" -f ($Script:PerformanceMetrics.TicketCloseTime / 1000)) -ForegroundColor White
    Write-Host ("    Total Ticket Operations  : {0,6:N2} seconds" -f ($totalTicketOperationTime / 1000)) -ForegroundColor Cyan
}
#endregion ----- Summary Output ----

# Display ConnectWise URL to view tickets
if ($Script:CWMServerConnection) {
    $cwmUrl = "https://$($Script:CWMServerConnection.Server)/v4_6_release/services/system_io/router/openrecord.rails?locale=en_US&recordType=ServiceFv&recid="
    Write-Host "`n  View Tickets in ConnectWise:" -ForegroundColor Cyan
    Write-Host "  https://$($Script:CWMServerConnection.Server)/v4_6_release/services/system_io/Service/index.html" -ForegroundColor Green
    
    if ($Script:TicketActions.Count -gt 0) {
        Write-Host "`n  Recent Ticket Links:" -ForegroundColor Cyan
        foreach ($action in ($Script:TicketActions | Select-Object -First 5)) {
            if ($action.TicketId) {
                Write-Host "  $cwmUrl$($action.TicketId) - $($action.Summary)" -ForegroundColor Gray
            }
        }
    }
    Write-Host ""
}

#endregion ----- Main Script Execution ----
























