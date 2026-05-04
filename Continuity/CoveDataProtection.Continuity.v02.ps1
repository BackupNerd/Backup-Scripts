<# ----- About: ----
    # Cove Data Protection - Business Continuity Report
    # Revision v02.0 - 2026-04-20
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able Cove Data Protection
    # Requires PowerShell 7+  (auto-relaunches in pwsh if run under PS 5)
    # Requires ImportExcel module for XLSX export
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Check / Get / Store secure API credentials
    # Authenticate to https://backup.management console
    # Retrieve all Business Continuity data for the specified partner, including:
    #   - Recovery Testing sessions
    #   - Standby Image status
    #   - DRaaS (Disaster Recovery as a Service) status
    #   - One-Time Restore sessions
    #   - Recovery Locations
    # Generate a self-contained HTML report for the Continuity section
    # Optionally export raw data to CSV or XLSX
    #
    # Use -PartnerName     to specify the partner to query
    # Use -ExportPath      to specify an alternate output folder
    # Use -ExportCSV       to export data as CSV
    # Use -ExportXLSX      to export data as XLSX (requires ImportExcel module)
    # Use -ClearCredentials to remove stored API credentials at start of script
    # Use -OfflineMode     to generate report from previously cached data
    #
    # Source reference: CoveDataProtection.HealthCheck.v53.5.ps1
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
# -----------------------------------------------------------#>

param (
    [Parameter(Mandatory=$false)][string]$PartnerName,
    [Parameter(Mandatory=$false)][string]$ExportPath      = "",
    [Parameter(Mandatory=$false)][switch]$ExportCSV,
    [Parameter(Mandatory=$false)][switch]$ExportXLSX,
    [Parameter(Mandatory=$false)][string]$delimiter        = ",",
    [Parameter(Mandatory=$false)][switch]$ClearCredentials,
    [Parameter(Mandatory=$false)][switch]$DebugScreenshots,
    [Parameter(Mandatory=$false)][switch]$OfflineMode
)

Set-Location $PSScriptRoot
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

## Version / meta
$ScriptVersion    = "Continuity.v02.0"
$Script:CurrentDate = Get-Date -Format "yyyyMMdd_HHmmss"

## ── PowerShell 7 auto-relaunch ──────────────────────────────────────────────
if ($PSVersionTable.PSVersion.Major -lt 7) {
    if (Get-Command pwsh -ErrorAction SilentlyContinue) {
        Write-Host "  Relaunching in PowerShell 7..." -ForegroundColor Yellow
        $argList = "-File `"$PSCommandPath`""
        if ($PartnerName)        { $argList += " -PartnerName `"$PartnerName`"" }
        if ($ExportPath)         { $argList += " -ExportPath `"$ExportPath`"" }
        if ($ExportCSV)          { $argList += " -ExportCSV" }
        if ($ExportXLSX)         { $argList += " -ExportXLSX" }
        if ($ClearCredentials)   { $argList += " -ClearCredentials" }
        if ($DebugScreenshots)   { $argList += " -DebugScreenshots" }
        if ($OfflineMode)        { $argList += " -OfflineMode" }
        Start-Process pwsh.exe -ArgumentList $argList -Wait
        exit
    } else {
        Write-Warning "  PowerShell 7 not found. Some features may not work correctly."
    }
}

## ── Script-scope defaults ────────────────────────────────────────────────────
$Script:strLineSeparator    = "─" * 80
$Script:DeviceDetail        = @()          ## No full statistics API needed
$Script:DashVersion         = 'v2'
$Script:cred0               = $null
$Script:visa                = $null
$Script:websession          = $null
$Script:partnername         = $null
$Script:partnerid           = $null
$Script:PartnerId           = $null
$Script:ExportPath          = if ($ExportPath) { $ExportPath } else { Join-Path $PSScriptRoot "ContinuityReports" }

Write-Output ""
Write-Output "  Cove Data Protection - Business Continuity Report  |  Script $ScriptVersion"
Write-Output $Script:strLineSeparator

#region ----- Functions ----
#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator
    Write-Output "  Setting Backup API Credentials"
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential Path Present"
    } else {
        New-Item -ItemType Directory -Path $APIcredpath -Force | Out-Null
    }

    Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO { $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove | Backup.Management API'

    # Create credential object in XML format (DPAPI encrypted)
    $CDPCredentials = [PSCustomObject]@{
        PartnerName = $Script:PartnerName
        Username    = $BackupCred.UserName
        Password    = ($BackupCred.Password | ConvertFrom-SecureString)
    }
    $CDPCredentials | Export-Clixml -Path $APIcredfile -Force

    Write-Output "  ✓ Credentials saved to: $APIcredfile"

    Start-Sleep -Milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    if (-not (Test-Path -Path $Script:True_path)) {
        New-Item -ItemType Directory -Path $Script:True_path -Force | Out-Null
    }
    $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "${env:computername}_${env:username}_API_Credentials.Secure.xml"
    $Script:APIcredpath = Split-Path -Path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) {
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential File Cleared"
        Send-APICredentialsCookie  ## Retry Authentication
    } else {
        Write-Output $Script:strLineSeparator
        Write-Output "  Getting Backup API Credentials"

        if (Test-Path $APIcredfile) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Present"
            $APIcredentials = Import-Clixml -Path $APIcredfile

            $Script:cred0 = [string]$APIcredentials.PartnerName
            $Script:cred1 = [string]$APIcredentials.Username
            $Script:cred2 = $APIcredentials.Password | ConvertTo-SecureString

            Write-Output $Script:strLineSeparator
            Write-Output "  Stored Backup API Partner  = $Script:cred0"
            Write-Output "  Stored Backup API User     = $Script:cred1"
            Write-Output "  Stored Backup API Password = Encrypted`n"
        } else {
            Write-Output $Script:strLineSeparator
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
    $data.params.partner  = $Script:cred0
    $data.params.username = $Script:cred1
    $plainPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))
    $data.params.password = $plainPwd

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $plainPwd = $null  ## Clear plaintext password immediately after use
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Authenticate = $webrequest | convertfrom-json

    #Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($authenticate.visa) { 

        $Script:visa = $authenticate.visa
        $script:UserId = $authenticate.result.result.id
           
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-Warning "`n  $($authenticate.error.message)"
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"

        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Test-CDPVisa{
     if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){   
            Send-APICredentialsCookie
        }
    }
}  ## Recheck remaining Visa time and reauthenticate

#endregion ----- Authentication ----
#region ----- Data Conversion ----

Function Convert-BackupFrequency {
    param([string]$FrequencyCode)
    
    # Translate N-able Cove API frequency codes to human-readable format
    $frequencyMap = @{
        'C' = 'Continuous'
        'D' = 'Daily'
        'E' = 'Every 24 Hours'
        'W' = 'Weekly'
        'M' = 'Monthly'
        'Every1Hour' = 'Every 1 Hour'
        'Every2Hours' = 'Every 2 Hours'
        'Every3Hours' = 'Every 3 Hours'
        'Every4Hours' = 'Every 4 Hours'
        'Every6Hours' = 'Every 6 Hours'
        'Every8Hours' = 'Every 8 Hours'
        'Every12Hours' = 'Every 12 Hours'
        'Every24Hours' = 'Every 24 Hours'
    }
    
    if ($frequencyMap.ContainsKey($FrequencyCode)) {
        return $frequencyMap[$FrequencyCode]
    } else {
        return $FrequencyCode  # Return original if no mapping found
    }
}

Function Measure-ElapsedTime {
    if (-not $script:StartTime) {
        $script:StartTime = Get-Date
    }else{
        $script:EndTime = Get-Date
        $elapsedTime = $script:EndTime - $script:StartTime
        $elapsedTime = "{0:hh\:mm\:ss}" -f $elapsedTime
        Write-Output "  Elapsed time: $elapsedTime`n"
    }
}

Function Convert-UnixTimeToDateTime($UnixToConvert) {
    if ($UnixToConvert -gt 0 ) { 
        $Epoch2Date = ((Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()).AddSeconds($UnixToConvert)
        return $Epoch2Date.ToLocalTime()
    }else{ return ""}
}  ## Convert epochtime to datetime and return as local time #Rev.04

Function Convert-DateTimeToUnixTime($DateToConvert) {
    $Date2Epoch = (New-TimeSpan -Start (Get-Date -Date "1970-01-01 00:00:00Z")-End (Get-Date -Date $DateToConvert)).TotalSeconds
    Return $Date2Epoch
}  ## Convert datetime to epochtime #Rev.03

Function Shorten-OSVersion($OSString) {
    if ([string]::IsNullOrWhiteSpace($OSString)) {
        return ""
    }
    
    # Shorten OS version for display in tables: Standard -> Std, Edition -> Edt
    $OSString = $OSString -replace 'Standard', 'Std'
    $OSString = $OSString -replace 'Edition', 'Edt'
    
    return $OSString
}  ## Shorten OS version strings for compact display

Function Convert-DayNamesToAbbreviations($DayString) {
    if ([string]::IsNullOrWhiteSpace($DayString)) {
        return ""
    }
    
    $DayString = $DayString -replace 'Sunday', 'Su'
    $DayString = $DayString -replace 'Monday', 'M'
    $DayString = $DayString -replace 'Tuesday', 'Tu'
    $DayString = $DayString -replace 'Wednesday', 'W'
    $DayString = $DayString -replace 'Thursday', 'Th'
    $DayString = $DayString -replace 'Friday', 'F'
    $DayString = $DayString -replace 'Saturday', 'Sa'
    
    return $DayString
}  ## Convert day names to abbreviations

Function Format-StorageSize {
    <#
    .SYNOPSIS
        Converts storage size in bytes to standardized display format
    
    .DESCRIPTION
        Converts storage values from bytes to the lowest unit where value has max 3 digits left of decimal:
        - Maximum 3 digits to the left of decimal
        - Exactly 2 digits to the right of decimal
        - Automatic unit selection (GB, TB, PB)
    
    .PARAMETER SizeBytes
        Storage size in bytes (can be decimal)
    
    .EXAMPLE
        Format-StorageSize -SizeBytes 1326445977600
        Returns: "1.23 TB" (1326445977600 bytes / 1GB / 1024 = 1.23 TB)
    
    .EXAMPLE
        Format-StorageSize -SizeBytes 537109504000
        Returns: "500.00 GB" (stays in GB, less than 1000)
    
    .EXAMPLE
        Format-StorageSize -SizeBytes 4298339516416
        Returns: "3.91 TB" (4298339516416 bytes / 1GB / 1024 = 3.91 TB)
    
    .EXAMPLE
        Format-StorageSize -SizeBytes 1125899906842624
        Returns: "1.00 PB" (1125899906842624 bytes / 1GB / 1024 / 1024 = 1.00 PB)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [decimal]$SizeBytes
    )
    
    # Convert bytes to GB first
    $sizeGB = $SizeBytes / 1GB
    
    # If less than 1000 GB, display in GB
    if ($sizeGB -lt 1000) {
        return "{0:N2} GB" -f $sizeGB
    }
    
    # Convert to TB
    $sizeTB = $sizeGB / 1024
    
    # If less than 1000 TB, display in TB
    if ($sizeTB -lt 1000) {
        return "{0:N2} TB" -f $sizeTB
    }
    
    # Convert to PB
    $sizePB = $sizeTB / 1024
    
    # Display in PB
    return "{0:N2} PB" -f $sizePB
}  ## Convert GB to lowest unit with max 3 digits left of decimal (###.## GB/TB/PB)

#endregion ----- Data Conversion ----
#region ----- Backup.Management JSON Calls ----

Function CallJSON($url,$object) {
    # Use Invoke-RestMethod for reliable SSL/TLS handling in PowerShell 7+
    # This automatically uses TLS 1.2/1.3 and handles certificate validation properly
    
    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $object -ContentType "application/json; charset=utf-8" -SkipCertificateCheck
        return $response
    }
    catch {
        # If SkipCertificateCheck fails (PS 5.1), fall back to legacy method with SSL settings
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
        [Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
        
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $object -ContentType "application/json; charset=utf-8"
        return $response
    }
}
Function Send-GetPartnerInfo ($PartnerName) { 
                    
    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfo'
    $data.params = @{}
    $data.params.name = [String]$PartnerName

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json

    $RestrictedPartnerLevel = @("Root","SubRoot","Distributor","EndCustomer","Site")
    
    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        Set-Variable -Name PartnerId -Value ([int]$Partner.result.result.Id) -Scope Script
        Set-Variable -Name partnerid -Value ([int]$Partner.result.result.Id) -Scope Script  ## Lowercase version for file naming
        Set-Variable -Name Level -Value $Partner.result.result.Level -Scope Script
        [String]$Script:PartnerName = $Partner.result.result.Name
        [String]$Script:TopLevelParentId = $Partner.result.result.ParentId

        Write-Output $Script:strLineSeparator
        Write-output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
        }else{
        Write-Output $Script:strLineSeparator
        Write-Host "  ⚠️  ERROR: Partner Level '$($Partner.result.result.Level)' Not Allowed" -ForegroundColor Red
        Write-Host "  This script can only be run on Partner/MSP accounts, not End Customers or Sites" -ForegroundColor Yellow
        Write-Host "  Allowed levels: Company, ServiceOrganization" -ForegroundColor Yellow
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Partner/MSP name (not End Customer)"
        Send-GetPartnerInfo $Script:partnername
        }

    if ($partner.error) {
        write-warning "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername

    }else{$script:visa = $Partner.visa}

} ## get PartnerID and Partner Level    
Function Get-ContinuityDevicesByType {
    <#
    .SYNOPSIS
        Retrieves continuity devices for a specific category (RT, SBI, OTR, or DRaaS)
    .DESCRIPTION
        Makes a separate API call for each device type category to avoid type overlap confusion.
        This ensures devices like SELF_HOSTED_ON_DEMAND are clearly categorized as One-Time Restore,
        not incorrectly mixed with Standby Image devices due to pattern matching issues.
    .PARAMETER DeviceTypes
        Comma-separated device type codes for filter (e.g., "RECOVERY_TESTING" or "SELF_HOSTED,AZURE_SELF_HOSTED")
    .PARAMETER CategoryName
        Human-readable category name for logging (e.g., "Recovery Testing", "Standby Image")
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DeviceTypes,
        
        [Parameter(Mandatory=$true)]
        [string]$CategoryName
    )
    
    try {
        # Pagination settings
        $limit = 100
        $offset = 0
        $allDevices = @()
        $continueLoop = $true
        
        while ($continueLoop) {
            $filterParams = @(
                "offset=$offset"
                "limit=$limit"
                "fields=$Script:urlFields"
                "sort=-last_recovery_selected_size_user"
                "filter%5Btype.in%5D=$DeviceTypes"
                "filter%5Bpartner_materialized_path.contains%5D=/$($Script:PartnerId)/"
            )
            
            $url = "https://api.backup.management/draas/actual-statistics/v1/dashboard/?" + ($filterParams -join "&")
            
            if ($offset -eq 0) {
                Write-Host "    API filter: type.in=[$DeviceTypes]" -ForegroundColor Gray
            }
            
            $method = 'GET'
            
            $params = @{
                Uri         = $url
                Method      = $method
                Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
                WebSession  = $Script:websession
                ContentType = 'application/json; charset=utf-8'
            }
            
            $response = Invoke-RestMethod @params -ErrorAction Stop
            
            if ($response.data -and $response.data.Count -gt 0) {
                $allDevices += $response.data.attributes
                
                if ($offset -eq 0) {
                    Write-Host "    Retrieved $($response.data.Count) devices" -ForegroundColor Gray
                }
                
                # Check if there are more devices
                if ($response.data.Count -lt $limit) {
                    $continueLoop = $false
                } else {
                    $offset += $limit
                }
            } else {
                $continueLoop = $false
            }
        }
        
        # Process data if devices found
        if ($allDevices.Count -gt 0) {
            # Build Select-Object with calculated properties for array fields
            $selectProperties = @()
            
            foreach ($field in $Script:continuityFields) {
                if ($field -eq 'colorbar') {
                    # Colorbar: array of session objects - extract status from each
                    $selectProperties += @{
                        Name = 'colorbar'
                        Expression = { 
                            if ($_.colorbar -is [Array]) {
                                $statusValues = @($_.colorbar | ForEach-Object { 
                                    if ($_ -and $_.status) { $_.status } 
                                } | Where-Object { $_ })
                                $statusValues -join ','
                            } else { 
                                $_.colorbar 
                            }
                        }
                    }
                    # Detailed colorbar: preserve date and status for richer tooltips
                    $selectProperties += @{
                        Name = 'colorbar_detail'
                        Expression = { 
                            if ($_.colorbar -is [Array]) {
                                $detailEntries = @($_.colorbar | ForEach-Object { 
                                    if ($_ -and $_.status -and $_.backup_session_timestamp) {
                                        # Format: "timestamp|status" for easy parsing
                                        "$($_.backup_session_timestamp)|$($_.status)"
                                    }
                                } | Where-Object { $_ })
                                $detailEntries -join ','
                            } else { 
                                '' 
                            }
                        }
                    }
                } elseif ($field -eq 'data_sources') {
                    # Data sources: array of datasource codes - join with commas
                    $selectProperties += @{
                        Name = 'data_sources'
                        Expression = { 
                            if ($_.data_sources -is [Array]) { 
                                $_.data_sources -join ',' 
                            } else { 
                                $_.data_sources 
                            }
                        }
                    }
                } elseif ($field -eq 'recovery_session_progress') {
                    # Recovery session progress: extract percentage or status from nested object
                    $selectProperties += @{
                        Name = 'recovery_session_progress'
                        Expression = { 
                            if ($_.recovery_session_progress -is [PSCustomObject]) {
                                if ($_.recovery_session_progress.percentage) {
                                    "$($_.recovery_session_progress.percentage)%"
                                } elseif ($_.recovery_session_progress.status) {
                                    $_.recovery_session_progress.status
                                } else {
                                    ($_.recovery_session_progress | ConvertTo-Json -Compress)
                                }
                            } elseif ($_.recovery_session_progress -is [Array]) {
                                $_.recovery_session_progress -join ','
                            } else {
                                $_.recovery_session_progress
                            }
                        }
                    }
                } else {
                    # All other fields: pass through as-is
                    $selectProperties += $field
                }
            }
            
            # Select columns with calculated properties
            $processedDevices = $allDevices | Select-Object -Property $selectProperties
            
            # Convert data sizes from bytes to GB
            $processedDevices | ForEach-Object {
                if ($_.last_recovery_selected_size_user) {
                    $_.last_recovery_selected_size_user = [Math]::Round([Decimal]($_.last_recovery_selected_size_user / 1GB), 2)
                }
                if ($_.last_recovery_restored_size_user) {
                    $_.last_recovery_restored_size_user = [Math]::Round([Decimal]($_.last_recovery_restored_size_user / 1GB), 2)
                }
            }
            
            # Preserve Unix timestamps BEFORE conversion (needed for filename construction)
            $processedDevices | ForEach-Object {
                if ($_.last_backup_session_timestamp) {
                    $_ | Add-Member -NotePropertyName 'backup_unix' -NotePropertyValue $_.last_backup_session_timestamp -Force
                }
                if ($_.last_boot_test_recovery_session_timestamp) {
                    $_ | Add-Member -NotePropertyName 'recovery_unix' -NotePropertyValue $_.last_boot_test_recovery_session_timestamp -Force
                }
            }
            
            # Convert Unix timestamps to DateTime
            $processedDevices | ForEach-Object {
                if ($_.last_recovery_timestamp) {
                    $_.last_recovery_timestamp = Convert-UnixTimeToDateTime $_.last_recovery_timestamp
                }
                if ($_.last_backup_session_timestamp) {
                    $_.last_backup_session_timestamp = Convert-UnixTimeToDateTime $_.last_backup_session_timestamp
                }
                if ($_.last_boot_test_backup_session_timestamp) {
                    $_.last_boot_test_backup_session_timestamp = Convert-UnixTimeToDateTime $_.last_boot_test_backup_session_timestamp
                }
                if ($_.last_boot_test_recovery_session_timestamp) {
                    $_.last_boot_test_recovery_session_timestamp = Convert-UnixTimeToDateTime $_.last_boot_test_recovery_session_timestamp
                }
            }
            
            # Format recovery duration (preserve numeric value for GB/HR calculations)
            $processedDevices | ForEach-Object {
                if ($_.last_recovery_duration_user) {
                    $totalSeconds = [int]$_.last_recovery_duration_user
                    # Preserve numeric seconds for calculations
                    $_ | Add-Member -NotePropertyName 'last_recovery_duration_seconds' -NotePropertyValue $totalSeconds -Force
                    # Format display string
                    $hours = [math]::Floor($totalSeconds / 3600)
                    $minutes = [math]::Floor(($totalSeconds % 3600) / 60)
                    $seconds = $totalSeconds % 60
                    $_.last_recovery_duration_user = "{0}h {1}m {2}s" -f $hours, $minutes, $seconds
                }
            }
            
            # Enrich with OS, TimeZone, and ConsoleLink from DeviceDetail (Statistics API)
            $processedDevices | ForEach-Object {
                $deviceId = $_.backup_cloud_device_id
                $deviceType = $_.type
                
                # Get OS, timezone, and other data from DeviceDetail (Statistics API with TZ column)
                $deviceDetail = $Script:DeviceDetail | Where-Object { $_.DeviceId -eq $deviceId } | Select-Object -First 1
                
                # Add OS (Operating System)
                if ($deviceDetail -and $deviceDetail.OS) {
                    $_ | Add-Member -NotePropertyName 'OS' -NotePropertyValue $deviceDetail.OS -Force
                } else {
                    $_ | Add-Member -NotePropertyName 'OS' -NotePropertyValue 'Unknown' -Force
                }
                
                # Add TimeZone - DeviceDetail.TimeZone is already formatted as "UTC +X" from TZ column
                $tzOffset = 'UTC +0'
                if ($deviceDetail -and $deviceDetail.TimeZone) {
                    # TimeZone is already formatted (e.g., "UTC +11", "UTC +1", "UTC")
                    $tzOffset = $deviceDetail.TimeZone
                }
                $_ | Add-Member -NotePropertyName 'TimeZone' -NotePropertyValue $tzOffset -Force
                $_ | Add-Member -NotePropertyName 'TimeZoneOffset' -NotePropertyValue $tzOffset -Force
                
                # Build console link based on device type
                $consoleTypeParam = switch -Wildcard ($deviceType) {
                    'RECOVERY_TESTING' { 'RECOVERY_TESTING' }
                    '*SELF_HOSTED*' { 'STANDBY_IMAGE' }
                    '*DRAAS*' { 'DRAAS' }
                    '*ON_DEMAND*' { 'ONE_TIME_RESTORE' }
                    default { 'RECOVERY_TESTING' }
                }
                $consoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/$consoleTypeParam)"
                $_ | Add-Member -NotePropertyName 'ConsoleLink' -NotePropertyValue $consoleLink -Force
            }
            
            Write-Host "    --> ${CategoryName}: $($processedDevices.Count) devices retrieved and processed" -ForegroundColor Green
            return $processedDevices
        } else {
            Write-Host "    --> ${CategoryName}: No devices found" -ForegroundColor Yellow
            return @()
        }
        
    } catch {
        Write-Warning "    Failed to retrieve ${CategoryName} devices: $($_.Exception.Message)"
        return @()
    }
    
} ## Get-ContinuityDevicesByType
Function Get-ContinuityStatistics {
    <#
    .SYNOPSIS
        Retrieves continuity/recovery testing statistics from DRaaS API
    .DESCRIPTION
        Queries the Backup.Management DRaaS actual-statistics endpoint for recovery testing,
        standby image, and DRaaS device information. Includes all 40+ fields from the API.
        
        NOTE: This endpoint may return 400 Bad Request if:
        - Partner doesn't have DRaaS/Continuity features enabled
        - No continuity devices are configured
        - API requires additional permissions not available in standard visa token
        
        The script will continue normally if this fails - continuity data is optional.
    #>
    
    Write-Output $Script:strLineSeparator
    Write-Output "  Retrieving Continuity Statistics"
    
    # Shared field list and URL field string for all continuity API calls
    $Script:continuityFields = @(
        "backup_cloud_device_id",
        "plan_device_id",
        "backup_cloud_partner_id",
        "last_recovery_session_id",
        "current_recovery_status",
        "type",
        "region_name",
        "agent_id",
        "backup_cloud_device_name",
        "backup_cloud_partner_name",
        "recovery_target_type",
        "recovery_agent_state",
        "recovery_agent_name",
        "backup_cloud_device_status",
        "manual_rerun_available",
        "last_backup_session_timestamp",
        "last_boot_test_backup_session_timestamp",
        "last_boot_test_recovery_session_timestamp",
        "last_boot_test_session_id",
        "plan_name",
        "colorbar",
        "recovery_session_progress",
        "last_recovery_errors_count",
        "last_recovery_timestamp",
        "last_recovery_duration_user",
        "last_boot_test_status",
        "last_boot_test_screenshot_presented",
        "last_recovery_restored_files_count",
        "last_recovery_selected_files_count",
        "data_sources",
        "last_recovery_status",
        "last_recovery_restored_size_user",
        "last_recovery_selected_size_user",
        "backup_cloud_device_machine_os_type",
        "recovery_target_vm_virtual_switch",
        "recovery_target_vhd_path",
        "recovery_target_local_speed_vault",
        "recovery_target_lsv_path",
        "recovery_target_enable_replication_service",
        "recovery_target_vm_address",
        "recovery_target_subnet_mask",
        "recovery_target_gateway",
        "recovery_target_dns_server",
        "backup_cloud_device_alias"
    )
    
    $Script:urlFields = @(
        "backup_cloud_device_id,plan_device_id,backup_cloud_partner_id,last_recovery_session_id",
        "current_recovery_status,type,region_name,agent_id,backup_cloud_device_name",
        "backup_cloud_partner_name,recovery_target_type,recovery_agent_state,recovery_agent_name",
        "backup_cloud_device_status,manual_rerun_available,last_backup_session_timestamp",
        "last_boot_test_backup_session_timestamp,last_boot_test_recovery_session_timestamp",
        "last_boot_test_session_id,backup_cloud_device_name,backup_cloud_partner_name,plan_name,colorbar",
        "current_recovery_status,recovery_session_progress,recovery_agent_state",
        "last_recovery_errors_count,last_recovery_timestamp,last_recovery_duration_user",
        "last_boot_test_status,last_boot_test_screenshot_presented",
        "last_recovery_restored_files_count,last_recovery_selected_files_count,region_name",
        "data_sources,last_recovery_status,last_recovery_restored_size_user",
        "last_recovery_selected_size_user,backup_cloud_device_machine_os_type,recovery_target_type",
        "recovery_target_vm_virtual_switch,recovery_target_vhd_path",
        "recovery_target_local_speed_vault,recovery_target_lsv_path",
        "recovery_target_enable_replication_service,recovery_target_vm_address",
        "recovery_target_subnet_mask,recovery_target_gateway,recovery_target_dns_server",
        "backup_cloud_device_alias"
    ) -join ","
    
    # Initialize combined statistics array
    $Script:ContinuityStatistics = @()
    
    try {
        # ARCHITECTURE CHANGE: Separate API calls for each device type category
        # Benefits:
        #   - No type overlap confusion (SELF_HOSTED_ON_DEMAND is clearly OTR, not SBI)
        #   - No badge color mismatches (badge assigned per category, not per type pattern)
        #   - Simpler code (no complex if/elseif categorization logic)
        #   - Each category self-contained
        
        # Category 1: Recovery Testing devices
        Write-Output "  Retrieving Recovery Testing devices..."
        $Script:RecoveryTestingDevices = Get-ContinuityDevicesByType -DeviceTypes "RECOVERY_TESTING" -CategoryName "Recovery Testing"
        
        # Category 2: Standby Image devices
        Write-Output "  Retrieving Standby Image devices..."
        $Script:SBIDevices = Get-ContinuityDevicesByType -DeviceTypes "SELF_HOSTED,AZURE_SELF_HOSTED,ESXI_SELF_HOSTED" -CategoryName "Standby Image"
        
        # Category 3: One-Time Restore devices
        Write-Output "  Retrieving One-Time Restore devices..."
        $Script:OneTimeRestoreDevices = Get-ContinuityDevicesByType -DeviceTypes "AZURE,ESXI_ON_DEMAND,SELF_HOSTED_ON_DEMAND" -CategoryName "One-Time Restore"
        
        # Category 4: DRaaS devices
        Write-Output "  Retrieving DRaaS devices..."
        $Script:DRaaSDevices = Get-ContinuityDevicesByType -DeviceTypes "AZURE_DRAAS" -CategoryName "DRaaS"
        
        # Combine all devices for Excel export
        $Script:ContinuityStatistics = @($Script:RecoveryTestingDevices) + @($Script:SBIDevices) + @($Script:OneTimeRestoreDevices) + @($Script:DRaaSDevices)
        
        Write-Output "`nContinuity device retrieval complete:"
        Write-Output "  Category 1 - Recovery Testing: $($Script:RecoveryTestingDevices.Count) devices"
        Write-Output "  Category 2 - Standby Image (SBI): $($Script:SBIDevices.Count) devices"
        Write-Output "  Category 3 - One-Time Restore: $($Script:OneTimeRestoreDevices.Count) devices"
        Write-Output "  Category 4 - DRaaS: $($Script:DRaaSDevices.Count) devices"
        Write-Output "  --> Total: $($Script:ContinuityStatistics.Count) devices"
        
    } catch {
        Write-Warning "  Failed to retrieve continuity statistics: $($_.Exception.Message)"
        
        # Try to extract error details from response
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
                Write-Output "  API Error Response: $responseBody"
            } catch {
                Write-Output "  Status Code: $($_.Exception.Response.StatusCode.value__)"
                Write-Output "  Status Description: $($_.Exception.Response.StatusDescription)"
            }
        }
        
        Write-Output "  This may indicate: No continuity devices configured, or API endpoint requires special permissions"
        $Script:ContinuityStatistics = @()
        $Script:RecoveryTestingDevices = @()
        $Script:SBIDevices = @()
        $Script:OneTimeRestoreDevices = @()
        $Script:DRaaSDevices = @()
    }
    
    # Export to CSV if enabled
    if ($ExportCSV -and $Script:ContinuityStatistics.Count -gt 0) {
        Write-Output "  Exporting Continuity Data to CSV"
        $Script:ContinuityStatistics | Export-Csv -Path "$Script:ExportPath\$($Script:cleanpartnername)_$($Script:partnerid)_$($Script:CurrentDate)_Continuity.csv" -NoTypeInformation -Delimiter $delimiter
    }
    
    # Export to Excel if enabled - Create separate worksheets for each continuity type
    if ($ExportXLSX -and $Script:ContinuityStatistics.Count -gt 0) {
        Write-Host "  Exporting Continuity Data to Excel (separate worksheets per type)" -ForegroundColor Cyan
        
        $excelPath = "$Script:ExportPath\$($Script:cleanpartnername)_$($Script:partnerid)_$($Script:CurrentDate)_Combined.xlsx"
        
        # Export Recovery Testing devices
        if ($Script:RecoveryTestingDevices.Count -gt 0) {
            Write-Host "    Creating Recovery Testing worksheet ($($Script:RecoveryTestingDevices.Count) devices)" -ForegroundColor Gray
            $excelPackage = $Script:RecoveryTestingDevices | Export-Excel -Path $excelPath -WorksheetName "Recovery Testing" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $ws = $excelPackage.Workbook.Worksheets['Recovery Testing']
            if ($ws.Dimension.End.Row -gt 1) {
                $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $ws.Dimension.End.Column).Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $ws -Range "A2:${lastColLetter}$($ws.Dimension.End.Row)" -HorizontalAlignment Left
            }
            Close-ExcelPackage $excelPackage
        } else {
            Write-Host "    Recovery Testing: No devices (worksheet not created)" -ForegroundColor Yellow
        }
        
        # Export Standby Image (SBI) devices
        if ($Script:SBIDevices.Count -gt 0) {
            Write-Host "    Creating Standby Image worksheet ($($Script:SBIDevices.Count) devices)" -ForegroundColor Gray
            $excelPackage = $Script:SBIDevices | Export-Excel -Path $excelPath -WorksheetName "Standby Image" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $ws = $excelPackage.Workbook.Worksheets['Standby Image']
            if ($ws.Dimension.End.Row -gt 1) {
                $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $ws.Dimension.End.Column).Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $ws -Range "A2:${lastColLetter}$($ws.Dimension.End.Row)" -HorizontalAlignment Left
            }
            Close-ExcelPackage $excelPackage
        } else {
            Write-Host "    Standby Image: No devices (worksheet not created)" -ForegroundColor Yellow
        }
        
        # Export One-Time Restore devices
        if ($Script:OneTimeRestoreDevices.Count -gt 0) {
            Write-Host "    Creating One-Time Restore worksheet ($($Script:OneTimeRestoreDevices.Count) devices)" -ForegroundColor Gray
            $excelPackage = $Script:OneTimeRestoreDevices | Export-Excel -Path $excelPath -WorksheetName "One-Time Restore" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $ws = $excelPackage.Workbook.Worksheets['One-Time Restore']
            if ($ws.Dimension.End.Row -gt 1) {
                $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $ws.Dimension.End.Column).Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $ws -Range "A2:${lastColLetter}$($ws.Dimension.End.Row)" -HorizontalAlignment Left
            }
            Close-ExcelPackage $excelPackage
        } else {
            Write-Host "    One-Time Restore: No devices (worksheet not created)" -ForegroundColor Yellow
        }
        
        # Export DRaaS devices
        if ($Script:DRaaSDevices.Count -gt 0) {
            Write-Host "    Creating DRaaS worksheet ($($Script:DRaaSDevices.Count) devices)" -ForegroundColor Gray
            $excelPackage = $Script:DRaaSDevices | Export-Excel -Path $excelPath -WorksheetName "DRaaS" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $ws = $excelPackage.Workbook.Worksheets['DRaaS']
            if ($ws.Dimension.End.Row -gt 1) {
                $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $ws.Dimension.End.Column).Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $ws -Range "A2:${lastColLetter}$($ws.Dimension.End.Row)" -HorizontalAlignment Left
            }
            Close-ExcelPackage $excelPackage
        } else {
            Write-Host "    DRaaS: No devices (worksheet not created)" -ForegroundColor Yellow
        }
        
        $totalExported = $Script:RecoveryTestingDevices.Count + $Script:SBIDevices.Count + $Script:OneTimeRestoreDevices.Count + $Script:DRaaSDevices.Count
        Write-Host "  --> Continuity worksheets created successfully ($totalExported total devices across all types)" -ForegroundColor Green
    }
    
    # Download screenshots for devices that have them (online mode only)
    if (-not $OfflineMode -and $Script:ContinuityStatistics.Count -gt 0) {
        # Try to get screenshots for any device that has completed a recovery session
        # Check multiple indicators: session ID, screenshot flag, or recovery timestamp
        
        if ($DebugScreenshots) {
            Write-Host "\n[DEBUG] Screenshot Detection Analysis" -ForegroundColor Magenta
            Write-Host "  Total continuity devices: $($Script:ContinuityStatistics.Count)" -ForegroundColor Gray
            
            # Debug: Show breakdown by device type
            $typeBreakdown = $Script:ContinuityStatistics | Group-Object -Property type
            foreach ($group in $typeBreakdown) {
                Write-Host "    $($group.Name): $($group.Count) devices" -ForegroundColor Gray
            }
            
            # Debug: Show devices with each detection criteria
            $withSessionId = @($Script:ContinuityStatistics | Where-Object { $_.last_recovery_session_id -and $_.last_recovery_session_id -gt 0 })
            $withScreenshotFlag = @($Script:ContinuityStatistics | Where-Object { $_.last_boot_test_screenshot_presented -eq $true })
            $withRecoveryTimestamp = @($Script:ContinuityStatistics | Where-Object { $_.last_boot_test_recovery_session_timestamp -and $_.last_boot_test_recovery_session_timestamp -ne $null })
            
            Write-Host "\n  Detection Criteria Results:" -ForegroundColor Cyan
            Write-Host "    Has last_recovery_session_id: $($withSessionId.Count)" -ForegroundColor Gray
            Write-Host "    Has screenshot_presented flag: $($withScreenshotFlag.Count)" -ForegroundColor Gray
            Write-Host "    Has recovery_session_timestamp: $($withRecoveryTimestamp.Count)" -ForegroundColor Gray
            
            # Debug: Show DRaaS devices specifically
            $draasDevices = @($Script:ContinuityStatistics | Where-Object { $_.type -eq 'AZURE_DRAAS' })
            if ($draasDevices.Count -gt 0) {
                Write-Host "\n  DRaaS Device Details ($($draasDevices.Count) total):" -ForegroundColor Cyan
                foreach ($d in $draasDevices) {
                    Write-Host "    Device: $($d.backup_cloud_device_name)" -ForegroundColor White
                    Write-Host "      - last_recovery_session_id: $($d.last_recovery_session_id)" -ForegroundColor Gray
                    Write-Host "      - screenshot_presented: $($d.last_boot_test_screenshot_presented)" -ForegroundColor Gray
                    Write-Host "      - recovery_timestamp: $($d.last_boot_test_recovery_session_timestamp)" -ForegroundColor Gray
                    Write-Host "      - boot_test_status: $($d.last_boot_test_status)" -ForegroundColor Gray
                }
            }
            Write-Host "" # Blank line
        }
        
        $screenshotDevices = @($Script:ContinuityStatistics | Where-Object { 
            # Exclude OTR devices - they don't have boot test screenshots (on-demand recovery only)
            $isOTR = ($_.type -like "*ON_DEMAND*" -or $_.type -eq "AZURE" -or $_.type -eq "ESXI_ON_DEMAND")
            
            if ($isOTR) { return $false }
            
            # Include devices with screenshot data: RT, SBI, DRaaS
            # Prioritize last_boot_test_screenshot_presented for SBI/DRaaS (most reliable)
            ($_.last_boot_test_screenshot_presented -eq $true) -or
            ($_.last_boot_test_recovery_session_timestamp -and $_.last_boot_test_recovery_session_timestamp -ne $null) -or
            ($_.last_recovery_session_id -and $_.last_recovery_session_id -gt 0)
        })
        
        if ($screenshotDevices.Count -gt 0) {
            Write-Output "  Processing screenshots for $($screenshotDevices.Count) devices..."
            
            # CRITICAL: Re-authenticate before screenshot downloads to ensure fresh visa token
            # Screenshot processing happens near end of script, token may have expired
            Write-Host "  Refreshing authentication for screenshot downloads..." -ForegroundColor Yellow
            Send-APICredentialsCookie
            
            if ($DebugScreenshots) {
                Write-Host "[DEBUG] Devices selected for screenshot processing:" -ForegroundColor Magenta
                $screenshotDevices | Group-Object -Property type | ForEach-Object {
                    Write-Host "  $($_.Name): $($_.Count) devices" -ForegroundColor Gray
                }
            }
            
            # Create device lookup hashtable for OS and other metadata from EnumerateAccountStatistics
            # This allows us to match continuity devices to their full device data
            $deviceLookup = @{}
            foreach ($dev in $Script:DeviceDetail) {
                $deviceLookup[$dev.DeviceId] = $dev
            }
            
        # Create screenshots subdirectories by type (RecoveryTesting, StandbyImage, OneTimeRestore, DRaaS)
        $screenshotBaseDir = Join-Path $Script:ExportPath "screenshots"
        $screenshotDirs = @{
            'RT' = Join-Path $screenshotBaseDir "RecoveryTesting"
            'SBI' = Join-Path $screenshotBaseDir "StandbyImage"
            'OTR' = Join-Path $screenshotBaseDir "OneTimeRestore"
            'DRaaS' = Join-Path $screenshotBaseDir "DRaaS"
        }
        
        # Map abbreviations to full folder names for HTML paths
        $typeFolderNames = @{
            'RT' = 'RecoveryTesting'
            'SBI' = 'StandbyImage'
            'OTR' = 'OneTimeRestore'
            'DRaaS' = 'DRaaS'
        }
        
        foreach ($dir in $screenshotDirs.Values) {
            if (-not (Test-Path $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
        }
        
        # Set screenshot path based on device type
        $screenshotPath = $screenshotBaseDir
        
        $Script:ScreenshotFiles = @{}
        $downloadCount = 0
            $cachedCount = 0
            $processedCount = 0
            
            foreach ($device in $screenshotDevices) {
                $deviceId = $device.backup_cloud_device_id
                
                # Determine device type abbreviation for filename
                # Note: OTR devices should not reach here (excluded in filter above)
                $deviceType = $device.type
                $typeAbbrev = if ($deviceType -like "*DRAAS*") { "DRaaS" }
                              elseif ($deviceType -like "*SELF_HOSTED*") { "SBI" }
                              elseif ($deviceType -eq "RECOVERY_TESTING") { "RT" }
                              else { "UNKNOWN" }
                
                # Set screenshot path to type-specific subdirectory
                $screenshotPath = $screenshotDirs[$typeAbbrev]
                
                # DEBUG: Show every device we're processing
                Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
                Write-Host "PROCESSING DEVICE [$($processedCount + 1)/$($screenshotDevices.Count)]" -ForegroundColor Cyan
                Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
                Write-Host "  Name: $($device.backup_cloud_device_name)" -ForegroundColor White
                Write-Host "  Device ID: $deviceId" -ForegroundColor White
                Write-Host "  Type: $deviceType ($typeAbbrev)" -ForegroundColor White
                Write-Host "  Boot Test Session ID: $($device.last_boot_test_session_id)" -ForegroundColor White
                Write-Host "  Screenshot Presented: $($device.last_boot_test_screenshot_presented)" -ForegroundColor White
                Write-Host "  Last Recovery Session Time: $($device.last_boot_test_recovery_session_timestamp)" -ForegroundColor White
                
                # EFFICIENT SESSION-BASED CACHING: Use session ID to instantly determine if screenshot exists
                # Filename format: {DeviceId}_{Type}_{SessionId}_{BackupUnix}_{RecoveryUnix}_{DurationSec}.png
                # This allows instant Excel cross-reference and historical analysis from filename
                $cachedFile = $null
                $useCache = $false
                
                # Look for cached screenshot matching exact session ID
                $sessionId = $device.last_boot_test_session_id
                if ($sessionId) {
                    # We have a session ID - look for exact match
                    $potentialCachedFiles = Get-ChildItem -Path $screenshotPath -Filter "${deviceId}_${typeAbbrev}_${sessionId}_*.png" -ErrorAction SilentlyContinue
                    
                    if ($potentialCachedFiles.Count -gt 0) {
                        # Found exact session ID match - use it (most efficient path)
                        $cachedFile = $potentialCachedFiles | Select-Object -First 1
                        $useCache = $true
                        Write-Host "  >> CACHE HIT: Found screenshot for session $sessionId" -ForegroundColor Green
                    } else {
                        # Session ID exists but no cached file - download fresh
                        Write-Host "  >> NO CACHE: Session $sessionId not found, will download" -ForegroundColor Yellow
                    }
                } else {
                    # No session ID available - can't verify cache validity, download fresh
                    Write-Host "  >> NO SESSION ID: Cannot verify cache, will download fresh" -ForegroundColor Yellow
                }
                
                if ($useCache -and $cachedFile) {
                    # Found valid cached screenshot - reuse it without API call
                    Write-Host "  >> USING CACHED SCREENSHOT: $($cachedFile.Name)" -ForegroundColor Green
                    $cachedCount++
                    $processedCount++
                    Write-Output "    [$processedCount/$($screenshotDevices.Count)] Cached: $($device.backup_cloud_device_name)"
                    
                    # Parse Unix timestamps from filename if available
                    # Format: {DeviceId}_{Type}_{SessionId}_{BackupUnix}_{RecoveryUnix}_{DurationSec}.png
                    $fileName = $cachedFile.Name
                    if ($fileName -match '_([^_]+)_(\d+)_(\d+)_(\d+)\.png$') {
                        # New format with Unix timestamps
                        $parsedBackupUnix = [long]$matches[2]
                        $parsedRecoveryUnix = [long]$matches[3]
                        $parsedDurationSec = [int]$matches[4]
                        $parsedTimestamp = if ($parsedRecoveryUnix -gt 0) {
                            Convert-UnixTimeToDateTime $parsedRecoveryUnix
                        } elseif ($parsedBackupUnix -gt 0) {
                            Convert-UnixTimeToDateTime $parsedBackupUnix
                        } else {
                            $device.last_boot_test_recovery_session_timestamp
                        }
                    } else {
                        # Old format or fallback
                        $parsedTimestamp = $device.last_boot_test_recovery_session_timestamp
                    }
                    
                    # Look up device metadata from DeviceDetail (includes full OS version and timezone)
                    $deviceMetadata = $deviceLookup[[string]$deviceId]
                    $osVersionRaw = if ($deviceMetadata -and $deviceMetadata.OS) { $deviceMetadata.OS } else { $device.backup_cloud_device_machine_os_type }
                    $osVersion = Shorten-OSVersion $osVersionRaw
                    $timezone = if ($deviceMetadata -and $deviceMetadata.TimeZone) { $deviceMetadata.TimeZone } else { "UTC" }
                    
                    # Use composite key: DeviceId_Type_SessionId to prevent duplicate device IDs with same session ID but different types
                    $compositeKey = if ($device.last_recovery_session_id) {
                        "$deviceId`_$($device.type)`_$($device.last_recovery_session_id)"
                    } else {
                        "$deviceId`_$($device.type)"  # Fallback to DeviceId_Type
                    }
                    
                    # Handle InProgress status - look at colorbar for previous session
                    $displayStatus = $device.last_boot_test_status
                    $isPriorSession = $false
                    
                    if ($device.last_boot_test_status -eq 'InProgress') {
                        # Parse colorbar to find most recent completed session status
                        if ($device.colorbar) {
                            $colorbarStatuses = @($device.colorbar -split ',' | Where-Object { $_ -and $_ -ne 'InProgress' })
                            if ($colorbarStatuses.Count -gt 0) {
                                # Get the most recent non-InProgress status
                                $priorStatus = $colorbarStatuses[0]
                                $displayStatus = "$priorStatus (InProgress)"
                                $isPriorSession = $true
                            }
                        }
                    }
                    
                    # Store cached screenshot metadata
                    $Script:ScreenshotFiles[$compositeKey] = @{
                        Path = "screenshots/$($typeFolderNames[$typeAbbrev])/$fileName"
                        DeviceName = $device.backup_cloud_device_name
                        DeviceType = $device.type
                        RecoveryTargetType = $device.recovery_target_type
                        Status = $displayStatus
                        IsPriorSession = $isPriorSession
                        Timestamp = $parsedTimestamp
                        PartnerName = $device.backup_cloud_partner_name
                        OS = $osVersion
                        TimeZone = $timezone  # Use DeviceDetail timezone from Excel/Statistics API
                        BackupSessionTime = $device.last_boot_test_backup_session_timestamp
                        RecoverySessionTime = $device.last_boot_test_recovery_session_timestamp
                        RecoveryDuration = $device.last_recovery_duration_user
                        BootSchedule = $device.plan_name
                    }
                    continue  # Skip API call, move to next device
                }
                
                # No cached screenshot - try to download from API
                Write-Host "  >> NO CACHE - CALLING API..." -ForegroundColor Yellow
                $retryCount = 0
                $maxRetries = 2
                $screenshotSuccess = $false
                
                while (-not $screenshotSuccess -and $retryCount -le $maxRetries) {
                    try {
                        if ($retryCount -gt 0) {
                            Write-Host "    Retry $retryCount/$maxRetries for $($device.backup_cloud_device_name)..." -ForegroundColor Yellow
                            Start-Sleep -Seconds 2
                        }
                        
                        Write-Host "  >> CALLING Get-BootScreenshot for device ID $deviceId" -ForegroundColor Cyan
                        # CRITICAL: Use last_boot_test_session_id NOT last_recovery_session_id
                        $screenshot = Get-BootScreenshot -DeviceId $deviceId -SessionId $device.last_boot_test_session_id -DeviceData $device -DebugScreenshots:$DebugScreenshots
                        
                        # Fix: Don't use nested if() inside $() in Write-Host - causes syntax error
                        $apiResponseMessage = if ($screenshot -and $screenshot.Url) { 'SUCCESS - Got URL' } else { 'NO SCREENSHOT RETURNED' }
                        $apiResponseColor = if ($screenshot -and $screenshot.Url) { 'Green' } else { 'Red' }
                        Write-Host "  >> API RESPONSE: $apiResponseMessage" -ForegroundColor $apiResponseColor
                        
                        if ($screenshot -and $screenshot.Url) {
                            # Build filename with session ID and Unix timestamps for efficient lookup and historical analysis
                            # Format: {DeviceId}_{Type}_{SessionId}_{BackupUnix}_{RecoveryUnix}_{DurationSec}.png
                            # Example: 3386098_SBI_01KFXYZ123ABC_1737900000_1737901234_3600.png
                            
                            $sessionIdPart = if ($device.last_boot_test_session_id) { $device.last_boot_test_session_id } else { "NOSESSION" }
                            $backupUnix = if ($device.backup_unix) { $device.backup_unix } else { "0" }
                            $recoveryUnix = if ($device.recovery_unix) { $device.recovery_unix } else { "0" }
                            $durationSec = if ($device.last_recovery_duration_seconds) { $device.last_recovery_duration_seconds } else { "0" }
                            
                            $fileName = "${deviceId}_${typeAbbrev}_${sessionIdPart}_${backupUnix}_${recoveryUnix}_${durationSec}.png"
                            $filePath = Join-Path $screenshotPath $fileName
                            
                            # Download with retry
                            $downloadRetry = 0
                            $downloadSuccess = $false
                            
                            while (-not $downloadSuccess -and $downloadRetry -le 2) {
                                try {
                                    Invoke-WebRequest -Uri $screenshot.Url -OutFile $filePath -ErrorAction Stop
                                    
                                    # Scale screenshot to 50% (1024x768 -> 512x384) for optimized HTML display
                                    $resized = Resize-Screenshot -FilePath $filePath
                                    
                                    $downloadSuccess = $true
                                    $downloadCount++
                                    $processedCount++
                                    $sizeIndicator = if ($resized) { " (scaled 50%)" } else { "" }
                                    Write-Output "    [$processedCount/$($screenshotDevices.Count)] Downloaded: $($device.backup_cloud_device_name)$sizeIndicator"
                                } catch {
                                    $downloadRetry++
                                    if ($_.Exception.Message -like "*401*") {
                                        # URL expired, get fresh one
                                        Write-Host "    URL expired, getting fresh URL..." -ForegroundColor Yellow
                                        break  # Exit download loop, retry outer loop
                                    } elseif ($downloadRetry -le 2) {
                                        Write-Host "    Download failed, retry $downloadRetry/2..." -ForegroundColor Yellow
                                        Start-Sleep -Seconds 1
                                    } else {
                                        throw
                                    }
                                }
                            }
                            
                            if (-not $downloadSuccess) {
                                # Download failed, retry getting screenshot URL
                                $retryCount++
                                if ($retryCount -le $maxRetries) {
                                    continue  # Retry outer loop
                                } else {
                                    throw "Download failed after retries"
                                }
                            }
                            
                            # Use composite key: DeviceId_Type_SessionId to prevent duplicate device IDs with same session ID but different types
                            # CRITICAL: Use last_boot_test_session_id (matches filename) NOT last_recovery_session_id
                            # This ensures metadata key matches the filename pattern for proper gallery lookups
                            $compositeKey = if ($device.last_boot_test_session_id) {
                                "$deviceId`_$($device.type)`_$($device.last_boot_test_session_id)"
                            } else {
                                # No session ID - use timestamps to differentiate multiple recovery types with no session
                                "$deviceId`_$($device.type)`_NOSESSION_${backupUnix}_${recoveryUnix}"
                            }
                            
                            # Store metadata
                            $Script:ScreenshotFiles[$compositeKey] = @{
                                Path = "screenshots/$($typeFolderNames[$typeAbbrev])/$fileName"
                                DeviceName = $screenshot.DeviceName
                                DeviceType = $screenshot.Type
                                RecoveryTargetType = $screenshot.RecoveryTargetType
                                Status = $screenshot.Status
                                Timestamp = $screenshot.Timestamp
                                PartnerName = $screenshot.PartnerName
                                OS = $screenshot.OS
                                TimeZone = $screenshot.TimeZone
                                BackupSessionTime = $screenshot.BackupSessionTime
                                RecoverySessionTime = $screenshot.RecoverySessionTime
                                RecoveryDuration = $screenshot.RecoveryDuration
                                BootSchedule = $screenshot.BootSchedule
                            }
                            
                            # CLEANUP: Remove old session screenshots (keep last 3 sessions per device)
                            # This prevents storage accumulation from devices with frequent screenshots (every 15 mins vs 30+ days)
                            try {
                                $allDeviceFiles = Get-ChildItem -Path $screenshotPath -Filter "${deviceId}_${typeAbbrev}_*.png" -ErrorAction SilentlyContinue
                                
                                if ($allDeviceFiles.Count -gt 3) {
                                    # Extract unique session IDs from filenames
                                    $sessionFiles = $allDeviceFiles | ForEach-Object {
                                        if ($_.Name -match "^${deviceId}_${typeAbbrev}_([^_]+)_") {
                                            [PSCustomObject]@{
                                                File = $_
                                                SessionId = $matches[1]
                                            }
                                        }
                                    } | Where-Object { $_ }
                                    
                                    # Group by session, keep last 3 unique sessions
                                    $sessionsToKeep = $sessionFiles | 
                                        Group-Object SessionId | 
                                        Sort-Object { $_.Group[0].File.LastWriteTime } -Descending | 
                                        Select-Object -First 3 | 
                                        ForEach-Object { $_.Name }
                                    
                                    # Delete files from old sessions
                                    $filesToDelete = $sessionFiles | Where-Object { $sessionsToKeep -notcontains $_.SessionId }
                                    
                                    if ($filesToDelete) {
                                        $deletedCount = 0
                                        foreach ($fileObj in $filesToDelete) {
                                            try {
                                                Remove-Item -Path $fileObj.File.FullName -Force -ErrorAction Stop
                                                $deletedCount++
                                            } catch {
                                                # Ignore errors (file may already be deleted by another script instance)
                                            }
                                        }
                                        
                                        if ($deletedCount -gt 0) {
                                            Write-Host "  >> CLEANUP: Removed $deletedCount old session file(s)" -ForegroundColor Gray
                                        }
                                    }
                                }
                            } catch {
                                # Non-critical error - continue even if cleanup fails
                                Write-Host "  >> Warning: Screenshot cleanup failed: $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                            
                            $screenshotSuccess = $true
                        } else {
                            # No screenshot available
                            Write-Host "  >> NO SCREENSHOT URL RETURNED FROM API" -ForegroundColor Red
                            Write-Host "     This device has no screenshot file available in the API" -ForegroundColor Red
                            $processedCount++
                            $screenshotSuccess = $true
                        }
                        
                    } catch {
                        $retryCount++
                        
                        if ($retryCount -gt $maxRetries) {
                            # Final failure - log and continue to next device
                            $processedCount++
                            Write-Warning "    [$processedCount/$($screenshotDevices.Count)] Failed: $($device.backup_cloud_device_name) - $($_.Exception.Message)"
                            $screenshotSuccess = $true  # Exit loop, continue to next device
                        }
                    }
                }
            }
            
            # Report download vs cached statistics
            if ($cachedCount -gt 0) {
                Write-Output "  --> Downloaded $downloadCount new screenshots, reused $cachedCount cached screenshots"
            } else {
                Write-Output "  --> Successfully downloaded $downloadCount screenshots"
            }
            
            # Report size optimization
            if ($downloadCount -gt 0) {
                Write-Host "  --> Screenshots scaled to 50% (1024x768 → 512x384) for optimized display" -ForegroundColor Cyan
                $estimatedOriginalSizeMB = [math]::Round(($downloadCount * 279) / 1024, 1)
                $estimatedScaledSizeMB = [math]::Round(($downloadCount * 70) / 1024, 1)
                $savingsMB = $estimatedOriginalSizeMB - $estimatedScaledSizeMB
                Write-Host "  --> Estimated savings: ~$savingsMB MB (from ~$estimatedOriginalSizeMB MB to ~$estimatedScaledSizeMB MB)" -ForegroundColor Green
            }
        }
    }
    
} ## Get-ContinuityStatistics
Function Resize-Screenshot {
    <#
    .SYNOPSIS
        Resizes a PNG screenshot to 50% (512x384) for optimized HTML display
    .DESCRIPTION
        Scales downloaded screenshots from 1024x768 to 512x384, reducing file size by ~75%
        Original dimensions: 1024x768 (~279 KB avg)
        Scaled dimensions: 512x384 (~70 KB avg)
    .PARAMETER FilePath
        Path to the PNG file to resize
    .RETURNS
        $true if successful, $false if failed
    #>
    param(
        [Parameter(Mandatory=$true)][string]$FilePath
    )
    
    try {
        # Load the image
        $originalImage = [System.Drawing.Image]::FromFile($FilePath)
        
        # Calculate 50% scale (1024x768 -> 512x384)
        $newWidth = [int]($originalImage.Width * 0.5)
        $newHeight = [int]($originalImage.Height * 0.5)
        
        # Create new bitmap with scaled dimensions
        $scaledBitmap = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($scaledBitmap)
        
        # Use high-quality scaling
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        
        # Draw scaled image
        $graphics.DrawImage($originalImage, 0, 0, $newWidth, $newHeight)
        
        # Dispose original before overwriting
        $originalImage.Dispose()
        $graphics.Dispose()
        
        # Save scaled image (overwrite original)
        $scaledBitmap.Save($FilePath, [System.Drawing.Imaging.ImageFormat]::Png)
        $scaledBitmap.Dispose()
        
        return $true
    }
    catch {
        Write-Warning "Failed to resize screenshot $FilePath`: $($_.Exception.Message)"
        # Cleanup if objects exist
        if ($originalImage) { $originalImage.Dispose() }
        if ($graphics) { $graphics.Dispose() }
        if ($scaledBitmap) { $scaledBitmap.Dispose() }
        return $false
    }
} ## Resize-Screenshot

Function Get-BootScreenshot {
    <#
    .SYNOPSIS
        Retrieves boot screenshot for a specific device from DRaaS API
    .DESCRIPTION
        Queries for last boot test screenshot and returns the temporary download URL
    .PARAMETER DeviceId
        The backup_cloud_device_id to retrieve screenshot for
    .RETURNS
        Hashtable with screenshot URL and metadata, or $null if not available
    #>
    param(
        [Parameter(Mandatory=$true)][int]$DeviceId,
        [Parameter(Mandatory=$false)][string]$SessionId,
        [Parameter(Mandatory=$false)][object]$DeviceData,
        [Parameter(Mandatory=$false)][switch]$DebugScreenshots
    )
    
    try {
        # CRITICAL FIX: Query dashboard with THIS specific device ID to get ALL recovery plan entries
        # A single device can have multiple entries (e.g., SELF_HOSTED + AZURE_DRAAS)
        # Each entry has its own last_boot_test_session_id where screenshots are stored
        
        $deviceTypes = 'RECOVERY_TESTING,SELF_HOSTED,AZURE_SELF_HOSTED,ESXI_SELF_HOSTED,AZURE_DRAAS'
        $url = "https://api.backup.management/draas/actual-statistics/v1/dashboard/" +
               "?filter%5Bbackup_cloud_device_id.eq%5D=$DeviceId" +
               "&filter%5Btype.in%5D=$deviceTypes"
        
        $params = @{
            Uri         = $url
            Method      = 'GET'
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $Script:websession
            ContentType = 'application/json; charset=utf-8'
        }
        
        $dashboardResponse = Invoke-RestMethod @params -ErrorAction Stop
        
        # Get all entries for this device (may be multiple recovery plans)
        $deviceEntries = @($dashboardResponse.data.attributes)
        
        if ($deviceEntries.Count -eq 0) {
            Write-Warning "  Device ID $DeviceId not found in dashboard query"
            return $null
        }
        
        if ($DebugScreenshots) {
            Write-Host "  Found $($deviceEntries.Count) recovery plan entries for device $DeviceId" -ForegroundColor Cyan
            foreach ($entry in $deviceEntries) {
                Write-Host "    - Type: $($entry.type), Boot Test Session: $($entry.last_boot_test_session_id)" -ForegroundColor Gray
            }
        }
        
        # Find the first entry with a valid boot test session ID and screenshot_presented = true
        $deviceData = $deviceEntries | Where-Object { 
            $_.last_boot_test_session_id -and 
            $_.last_boot_test_screenshot_presented -eq $true 
        } | Select-Object -First 1
        
        if (-not $deviceData) {
            # Fallback: Use any entry with a boot test session ID (even if screenshot_presented is false/null)
            $deviceData = $deviceEntries | Where-Object { 
                $_.last_boot_test_session_id 
            } | Select-Object -First 1
        }
        
        if (-not $deviceData) {
            if ($DebugScreenshots) {
                Write-Host "  No recovery plan entry has a boot test session ID" -ForegroundColor Yellow
            }
            return $null
        }
        
        # CRITICAL: Use boot test session ID, NOT regular recovery session ID
        # This is where screenshots are actually stored
        if ($SessionId) {
            $sessionId = $SessionId
        } else {
            $sessionId = $deviceData.last_boot_test_session_id
        }
        
        # ALWAYS DEBUG FOR SBI
        $forceDebug = ($deviceData.type -like "*SELF_HOSTED*")
        
        if ($DebugScreenshots -or $forceDebug) {
            Write-Host "`n  [DEBUG] Get-BootScreenshot for device: $($deviceData.backup_cloud_device_name)" -ForegroundColor Magenta
            Write-Host "    Device ID: $DeviceId" -ForegroundColor Gray
            Write-Host "    Device Type: $($deviceData.type)" -ForegroundColor Gray
            Write-Host "    Boot test session ID: $sessionId" -ForegroundColor Cyan
        }
        
        # If no direct session ID, try querying sessions by device ID (for DRaaS devices)
        if (-not $sessionId -or $sessionId -le 0) {
            # Try to get the most recent session for this device
            if ($DebugScreenshots) {
                Write-Host "    --> Session ID not populated, attempting fallback query..." -ForegroundColor Yellow
            }
            
            try {
                $deviceSessionParams = @{
                    Uri         = "https://api.backup.management/draas/actual-statistics/v1/sessions/?filter[backup_cloud_device_id]=$DeviceId&sort=-id&page[limit]=1"
                    Method      = 'GET'
                    Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
                    WebSession  = $Script:websession
                    ContentType = 'application/json'
                }
                
                $sessionData = Invoke-RestMethod @deviceSessionParams -ErrorAction Stop
                
                if ($sessionData.data -and $sessionData.data[0].id) {
                    $sessionId = $sessionData.data[0].id
                    if ($DebugScreenshots) {
                        Write-Host "    --> Fallback query successful: Found session ID $sessionId" -ForegroundColor Green
                    } else {
                        Write-Host "    Found session ID $sessionId for device $DeviceId via device query" -ForegroundColor Gray
                    }
                } else {
                    if ($DebugScreenshots) {
                        Write-Host "    --> Fallback query returned no sessions" -ForegroundColor Red
                    }
                }
            } catch {
                if ($DebugScreenshots) {
                    Write-Host "    --> Fallback query failed: $($_.Exception.Message)" -ForegroundColor Red
                } else {
                    Write-Warning "  Could not query sessions for device $DeviceId`: $($_.Exception.Message)"
                }
            }
        } else {
            if ($DebugScreenshots) {
                Write-Host "    --> Using direct session ID from API field" -ForegroundColor Green
            }
        }
        
        if ($sessionId -and $sessionId -gt 0) {
            
            if ($DebugScreenshots -or $forceDebug) {
                Write-Host "    --> Querying screenshot files for session ID $sessionId" -ForegroundColor Cyan
            }
            
            # Get screenshot file metadata
            $params2 = @{
                Uri         = "https://api.backup.management/draas/actual-statistics/v1/sessions/$sessionId/files/?filter[file_type.in]=screenshot"
                Method      = 'GET'
                Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
                WebSession  = $Script:websession
                ContentType = 'application/json'
            }
            
            $filesData = Invoke-RestMethod @params2 -ErrorAction Stop
            
            if ($DebugScreenshots -or $forceDebug) {
                if ($filesData.data.id) {
                    Write-Host "    --> Found screenshot file ID: $($filesData.data.id)" -ForegroundColor Green
                } else {
                    Write-Host "    --> NO SCREENSHOT FILE FOUND in API response" -ForegroundColor Red
                    Write-Host "        API returned: $($filesData | ConvertTo-Json -Depth 2 -Compress)" -ForegroundColor Gray
                }
            }
            
            if ($filesData.data.id -and $sessionId) {
                # Get temporary download URL
                $params3 = @{
                    Uri         = "https://api.backup.management/draas/actual-statistics/v1/sessions/$sessionId/files/$($filesData.data.id)/get-temporary-url/"
                    Method      = 'POST'
                    Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
                    WebSession  = $Script:websession
                    ContentType = 'application/json'
                }
                
                $urlData = Invoke-RestMethod @params3 -ErrorAction Stop
                
                if ($urlData.data.attributes.url) {
                    # Enrich with OS version and timezone from DeviceDetail (match by DeviceId)
                    $deviceDetail = $Script:DeviceDetail | Where-Object { $_.DeviceId -eq $deviceData.backup_cloud_device_id } | Select-Object -First 1
                    
                    # DEBUG: Check if lookup is working
                    if ($DebugScreenshots -or $forceDebug) {
                        Write-Host "    --> DeviceDetail Lookup:" -ForegroundColor Cyan
                        Write-Host "        API DeviceId: $($deviceData.backup_cloud_device_id)" -ForegroundColor Gray
                        # Fix: Don't use nested if() inside $() - causes syntax error  
                        $deviceDetailFound = if ($deviceDetail) { 'YES' } else { 'NO' }
                        $deviceDetailColor = if ($deviceDetail) { 'Green' } else { 'Red' }
                        Write-Host "        Found DeviceDetail: $deviceDetailFound" -ForegroundColor $deviceDetailColor
                        if ($deviceDetail) {
                            Write-Host "        DeviceDetail.DeviceId: $($deviceDetail.DeviceId)" -ForegroundColor Gray
                            Write-Host "        DeviceDetail.TimeZone: $($deviceDetail.TimeZone)" -ForegroundColor Gray
                        }
                    }
                    
                    # Get timezone from device's Statistics API data (TZ column) which has actual device timezone
                    # Prefer Statistics API timezone (actual timezone name) over draas-dashboard offset
                    $tzOffset = if ($deviceDetail -and $deviceDetail.TimeZone) {
                        # Use timezone from Statistics API (TZ column - actual timezone name like 'America/New_York')
                        $deviceDetail.TimeZone
                    } elseif ($deviceData.backup_cloud_device_timezone) {
                        # Use timezone from DRaaS dashboard API (numeric offset)
                        $tz = $deviceData.backup_cloud_device_timezone
                        # If numeric offset (e.g., "1", "-5"), format as UTC±X
                        if ($tz -match '^-?\d+(\.\d+)?$') {
                            $offsetHours = [decimal]$tz
                            if ($offsetHours -eq 0) { "UTC" }
                            elseif ($offsetHours -gt 0) { "UTC+$offsetHours" }
                            else { "UTC$offsetHours" }
                        } else {
                            # Already formatted (e.g., "America/New_York")
                            $tz
                        }
                    } else { "UTC" }
                    
                    # Convert timestamps from Unix to DateTime if needed (handles both already-converted and raw Unix timestamps)
                    # CRITICAL FIX: Use boot-test-specific backup timestamp, not generic device backup timestamp
                    # Each recovery type (Hyper-V, ESXi) has its own boot test schedule with separate backup session times
                    $backupTime = if ($deviceData.last_boot_test_backup_session_timestamp -is [DateTime]) {
                        $deviceData.last_boot_test_backup_session_timestamp
                    } elseif ($deviceData.last_boot_test_backup_session_timestamp -and $deviceData.last_boot_test_backup_session_timestamp -gt 0) {
                        Convert-UnixTimeToDateTime $deviceData.last_boot_test_backup_session_timestamp
                    } else { $null }
                    
                    $recoveryTime = if ($deviceData.last_boot_test_recovery_session_timestamp -is [DateTime]) {
                        $deviceData.last_boot_test_recovery_session_timestamp
                    } elseif ($deviceData.last_boot_test_recovery_session_timestamp -and $deviceData.last_boot_test_recovery_session_timestamp -gt 0) {
                        Convert-UnixTimeToDateTime $deviceData.last_boot_test_recovery_session_timestamp
                    } else { $null }
                    
                    $bootTime = if ($deviceData.last_boot_test_backup_session_timestamp -is [DateTime]) {
                        $deviceData.last_boot_test_backup_session_timestamp
                    } elseif ($deviceData.last_boot_test_backup_session_timestamp -and $deviceData.last_boot_test_backup_session_timestamp -gt 0) {
                        Convert-UnixTimeToDateTime $deviceData.last_boot_test_backup_session_timestamp
                    } else { $null }
                    
                    # Format duration as human-readable string if it's raw seconds
                    $formattedDuration = if ($deviceData.last_recovery_duration_user) {
                        if ($deviceData.last_recovery_duration_user -match '^\d+$') {
                            # Raw seconds - format it
                            $totalSeconds = [int]$deviceData.last_recovery_duration_user
                            $hours = [math]::Floor($totalSeconds / 3600)
                            $minutes = [math]::Floor(($totalSeconds % 3600) / 60)
                            $seconds = $totalSeconds % 60
                            "{0}h {1}m {2}s" -f $hours, $minutes, $seconds
                        } else {
                            # Already formatted
                            $deviceData.last_recovery_duration_user
                        }
                    } else { $null }
                    
                    # Fix: Extract OS value before hashtable to avoid PowerShell parsing the if as a command
                    $osValue = if ($deviceDetail -and $deviceDetail.OS) { $deviceDetail.OS } else { $deviceData.backup_cloud_device_machine_os_type }
                    
                    return @{
                        Url = $urlData.data.attributes.url
                        DeviceName = $deviceData.backup_cloud_device_name
                        PartnerName = $deviceData.backup_cloud_partner_name
                        Status = $deviceData.last_boot_test_status
                        Timestamp = $bootTime
                        Type = $deviceData.type
                        RecoveryTargetType = $deviceData.recovery_target_type  # Hyper-V, VMware, Azure, Local VHD, etc.
                        OS = Shorten-OSVersion $osValue
                        TimeZone = $tzOffset
                        BackupSessionTime = $backupTime
                        RecoverySessionTime = $recoveryTime
                        RecoveryDuration = $formattedDuration
                        BootSchedule = $deviceData.plan_name
                    }
                }
            }
        }
        
        return $null
        
    } catch {
        if ($DebugScreenshots) {
            Write-Host "  [ERROR] Failed to retrieve screenshot for device $DeviceId" -ForegroundColor Red
            Write-Host "    Exception: $($_.Exception.Message)" -ForegroundColor Gray
            Write-Host "    StackTrace: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        } else {
            Write-Warning "  Failed to retrieve screenshot for device $DeviceId`: $($_.Exception.Message)"
        }
        return $null
    }
} ## Get-BootScreenshot

Function Get-RecoveryLocations {
    <#
    .SYNOPSIS
        Retrieves recovery location infrastructure agents (Hyper-V, VMware, Azure hosts)
    .DESCRIPTION
        Queries DRaaS API for recovery agents showing infrastructure, concurrency limits, and storage
    .PARAMETER PartnerId
        Partner ID to filter by (default: Script:PartnerId)
    .RETURNS
        Array of recovery location objects with infrastructure details
    #>
    param(
        [Parameter(Mandatory=$false)][int]$PartnerId = $Script:PartnerId
    )
    
    Write-Output "`nRetrieving Recovery Locations..."
    
    try {
        # Build URL-encoded filter parameters
        $filterParams = @(
            "filter%5Bagent_state.in%5D=ONLINE,OFFLINE,STORAGE_NOT_CONFIGURED"
            "filter%5Bmaterialized_path.contains%5D=/$PartnerId/"
            "sort=name"
        )
        
        $url = "https://api.backup.management/draas/actual-statistics/v1/dashboard/recovery-agents/?" + ($filterParams -join "&")
        
        Write-Output "  API URL: $url"
        
        $params = @{
            Uri         = $url
            Method      = 'GET'
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $Script:websession
            ContentType = 'application/json; charset=utf-8'
        }
        
        $response = Invoke-RestMethod @params -ErrorAction Stop
        
        if ($response.data) {
            $Script:RecoveryLocations = @()
            
            foreach ($item in $response.data) {
                $attrs = $item.attributes
                
                # Extract partition storage details
                $partitionTotal = if ($attrs.partition_total_size_bytes) { 
                    [math]::Round($attrs.partition_total_size_bytes / 1GB, 2) 
                } else { 0 }
                
                $partitionUsed = if ($attrs.partition_used_size_bytes) { 
                    [math]::Round($attrs.partition_used_size_bytes / 1GB, 2) 
                } else { 0 }
                
                $partitionFree = if ($partitionTotal -gt 0 -and $partitionUsed -gt 0) {
                    [math]::Round($partitionTotal - $partitionUsed, 2)
                } else { 0 }
                
                $partitionPct = if ($partitionTotal -gt 0) {
                    [math]::Round(($partitionUsed / $partitionTotal) * 100, 1)
                } else { 0 }
                
                $location = [PSCustomObject]@{
                    recovery_agent_id = $item.id
                    name = $attrs.name
                    agent_state = $attrs.agent_state
                    type = $attrs.type
                    concurrency_limit = $attrs.concurrency_limit
                    assigned_devices_number = $attrs.assigned_devices_number
                    partition_total_gb = $partitionTotal
                    partition_used_gb = $partitionUsed
                    partition_free_gb = $partitionFree
                    partition_used_percent = $partitionPct
                    partner_name = $attrs.partner_name
                    partner_id = $attrs.partner_id
                }
                
                $Script:RecoveryLocations += $location
            }
            
            Write-Output "  --> Retrieved $($Script:RecoveryLocations.Count) recovery locations"
            Write-Output "      ONLINE: $(($Script:RecoveryLocations | Where-Object {$_.agent_state -eq 'ONLINE'}).Count)"
            Write-Output "      OFFLINE: $(($Script:RecoveryLocations | Where-Object {$_.agent_state -eq 'OFFLINE'}).Count)"
            Write-Output "      STORAGE_NOT_CONFIGURED: $(($Script:RecoveryLocations | Where-Object {$_.agent_state -eq 'STORAGE_NOT_CONFIGURED'}).Count)"
        }
        
    } catch {
        Write-Warning "  Failed to retrieve recovery locations: $($_.Exception.Message)"
        
        # Try to extract error details
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
                Write-Output "  API Error Response: $responseBody"
            } catch {
                Write-Output "  Status Code: $($_.Exception.Response.StatusCode.value__)"
                Write-Output "  Status Description: $($_.Exception.Response.StatusDescription)"
            }
        }
        
        Write-Output "  This may indicate: No recovery infrastructure configured, or API endpoint requires special permissions"
        $Script:RecoveryLocations = @()
    }
    
    # Export to CSV if enabled
    if ($ExportCSV -and $Script:RecoveryLocations.Count -gt 0) {
        Write-Output "  Exporting Recovery Locations to CSV"
        $Script:RecoveryLocations | Export-Csv -Path "$Script:ExportPath\$($Script:cleanpartnername)_$($Script:partnerid)_$($Script:CurrentDate)_RecoveryLocations.csv" -NoTypeInformation -Delimiter $delimiter
    }
    
    # Export to Excel if enabled
    if ($ExportXLSX -and $Script:RecoveryLocations.Count -gt 0) {
        Write-Output "  Exporting Recovery Locations to Excel"
        
        $excelPackage = $Script:RecoveryLocations | Export-Excel -Path "$Script:ExportPath\$($Script:cleanpartnername)_$($Script:partnerid)_$($Script:CurrentDate)_Combined.xlsx" -WorksheetName "Recovery Locations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
        $ws = $excelPackage.Workbook.Worksheets['Recovery Locations']
        
        # Set left alignment for all data rows
        if ($ws.Dimension.End.Row -gt 1) {
            $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::new(1, $ws.Dimension.End.Column).Address -replace '\d+', ''
            Set-ExcelRange -Worksheet $ws -Range "A2:${lastColLetter}$($ws.Dimension.End.Row)" -HorizontalAlignment Left
        }
        
        Close-ExcelPackage $excelPackage
        Write-Output "  --> Recovery Locations worksheet created successfully ($($Script:RecoveryLocations.Count) locations)"
    }
    
} ## Get-RecoveryLocations

Function Install-RequiredModules {
    param (
        [string[]]$Modules
    )
    foreach ($Module in $Modules) {
        if (Get-Module -ListAvailable -Name $Module) { ## Check if the module is already installed
            Write-Host "  Module $Module Already Installed" ## Inform the user that the module is already installed
        } 
        else {
            try {
                Install-Module -Name $Module -Confirm:$True -Force ## Attempt to install the module
            }
            catch [Exception] {
                $_.message ## Output the exception message
                Write-Warning "  Module '$Module' not found, run with Administrator rights to install" ## Warn the user if the module is not found and suggest running with admin rights
                exit ## Exit the script
            }
        }
    }
}

Function ExitRoutine {
    Write-Output $Script:strLineSeparator
    Write-Output "  Secure credential file found here:"
    Write-Output $Script:strLineSeparator
    Write-Output "  & $APIcredfile"
    Write-Output ""
    Write-Output $Script:strLineSeparator
    Start-Sleep -Seconds 3
}

#region DASHBOARD SECTION BUILDERS
#region ========== V2 DASHBOARD ARCHITECTURE ==========
# Independent Section Builder Functions
# Each section builds independently with error handling to prevent cascade failures

Function Get-DashSectionInfo {
    <#
    .SYNOPSIS
        Returns section metadata (ID, name, description) for dashboard sections
    #>
    param([Parameter(Mandatory=$true)][string]$SectionKey)
    
    $sectionMap = @{
        'users'       = @{ ID = 'DASH-001'; Name = 'Users & Permissions'; Description = 'User accounts and access control' }
        'devices'     = @{ ID = 'DASH-002'; Name = 'Device Statistics'; Description = 'Device inventory and types' }
        'errors'      = @{ ID = 'DASH-003'; Name = 'Errors & Issues'; Description = 'Backup failures and troubleshooting' }
        'm365'        = @{ ID = 'DASH-004'; Name = 'Microsoft 365'; Description = 'M365 licensing and offboarding' }
        'restores'    = @{ ID = 'DASH-005'; Name = 'Restore Operations'; Description = 'Restore activity and success rates' }
        'policies'    = @{ ID = 'DASH-006'; Name = 'Backup Policies'; Description = 'Retention and backup configs' }
        'profiles'    = @{ ID = 'DASH-007'; Name = 'Backup Profiles'; Description = 'Profile configurations and datasources' }
        'partners'    = @{ ID = 'DASH-008'; Name = 'Partner Hierarchy'; Description = 'Partner structure and storage' }
        'continuity'  = @{ ID = 'DASH-009'; Name = 'Business Continuity'; Description = 'Standby Image and DR readiness' }
        'compliance'  = @{ ID = 'DASH-010'; Name = 'Compliance'; Description = 'Compliance status and auditing' }
        'advisories'  = @{ ID = 'DASH-011'; Name = 'Health Advisories'; Description = 'Recommendations and optimizations' }
        'security'    = @{ ID = 'DASH-012'; Name = 'Security'; Description = 'Encryption and security settings' }
    }
    
    if ($sectionMap.ContainsKey($SectionKey)) { return $sectionMap[$SectionKey] }
    else { return @{ ID = 'DASH-000'; Name = 'Unknown'; Description = 'Section not found' } }
}

Function New-DashSectionError {
    <#
    .SYNOPSIS
        Creates error placeholder for failed dashboard section
    #>
    param(
        [Parameter(Mandatory=$true)][hashtable]$SectionInfo,
        [Parameter(Mandatory=$true)][string]$ErrorMessage,
        [string]$ErrorDetails = ""
    )
    
    return @"
        <div class="content-section" id="$($SectionInfo.ID.ToLower() -replace 'dash-','section-')">
            <div class="section-header">
                <div class="section-title">
                    <span style="background: #1e90ff; color: white; padding: 3px 10px; border-radius: 12px; margin-right: 8px; font-size: 11px; font-weight: 600; letter-spacing: 0.5px;">$($sectionInfo.ID)</span>
                    $($SectionInfo.Name)
                </div>
                <div class="section-subtitle">⚠️ Section Error</div>
            </div>
            <div style="background: #fef2f2; border: 2px solid #ef4444; border-radius: 8px; padding: 20px; margin: 20px 0;">
                <div style="display: flex; gap: 16px;">
                    <div style="font-size: 40px;">⚠️</div>
                    <div>
                        <h3 style="margin: 0 0 8px 0; color: #991b1b;">Section Failed to Render</h3>
                        <p style="margin: 0; color: #7f1d1d;"><strong>Error:</strong> $ErrorMessage</p>
                        $(if ($ErrorDetails) { "<p style='margin: 8px 0 0 0; color: #7f1d1d; font-size: 0.9em;'>$ErrorDetails</p>" })
                        <div style="margin-top: 12px; padding: 10px; background: white; border-radius: 4px; border-left: 3px solid #3b82f6;">
                            <p style="margin: 0; font-size: 0.9em; color: #1e40af;">
                                ✓ Other sections continue working normally (independent architecture)
                            </p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

"@
}
Function Build-DashSectionContinuity {
    <#
    .SYNOPSIS
        Builds the Business Continuity section (DASH-009) for v2 dashboard architecture
    .DESCRIPTION
        Independent section builder for disaster recovery and continuity planning with detailed tables
        Now supports 5 categories: Recovery Testing, SBI, One-Time Restore, DRaaS, Recovery Locations
    #>
    param(
        [Parameter(Mandatory=$true)][int]$RecoveryTestedCount,
        [Parameter(Mandatory=$true)][int]$RecoveryOldTests,
        [Parameter(Mandatory=$true)][int]$RecoveryNeverTested,
        [Parameter(Mandatory=$true)][int]$StandbyImageCount,
        [Parameter(Mandatory=$false)]$RecoveryTestingDevices,
        [Parameter(Mandatory=$false)]$SBIDevices,
        [Parameter(Mandatory=$false)]$OneTimeRestoreDevices,
        [Parameter(Mandatory=$false)]$DRaaSDevices,
        [Parameter(Mandatory=$false)]$RecoveryLocations,
        [Parameter(Mandatory=$false)]$DevicesData
    )
    
    try {
        $sectionInfo = Get-DashSectionInfo -SectionKey 'continuity'
        
        # Build Recovery Testing table rows
        $recoveryTestingRows = ""
        if ($RecoveryTestingDevices -and @($RecoveryTestingDevices).Count -gt 0) {
            $recoveryTestingData = @($RecoveryTestingDevices | Sort-Object last_recovery_timestamp -Descending | Select-Object -First 20)
            
            foreach ($device in $recoveryTestingData) {
                $deviceName = "$($device.backup_cloud_device_name)"
                $partnerName = $device.backup_cloud_partner_name
                $status = $device.current_recovery_status
                $statusColor = switch ($status) {
                    'Completed' { '#10b981' }
                    'Failed' { '#ef4444' }
                    'InProgress' { '#f59e0b' }
                    default { '#6b7280' }
                }
                $lastRecovery = if ($device.last_recovery_timestamp) { 
                    ([datetime]$device.last_recovery_timestamp).ToString("MM/dd/yy HH:mm")
                } else { 
                    "N/A" 
                }
                $duration = if ($device.last_recovery_duration_user) { $device.last_recovery_duration_user } else { "N/A" }
                
                # Use last_recovery_restored_size_user from Excel (already in GB)
                $restoredSize = if ($device.last_recovery_restored_size_user) { "$($device.last_recovery_restored_size_user) GB" } else { "N/A" }
                
            # Calculate GB/hr metric (data already in GB from Excel)
            $gbPerHour = "N/A"
            if ($device.last_recovery_restored_size_user -and $device.last_recovery_duration_seconds) {
                try {
                    $sizeGB = [decimal]$device.last_recovery_restored_size_user  # Already in GB
                    $durationHours = [decimal]$device.last_recovery_duration_seconds / 3600
                    if ($durationHours -gt 0) {
                        $gbPerHourValue = [Math]::Round($sizeGB / $durationHours, 2)
                        $gbPerHour = "$gbPerHourValue"
                    }
                } catch {
                    $gbPerHour = "N/A"
                }
            }
                
                # Region and OS Type
                $region = if ($device.region_name) { $device.region_name } else { "N/A" }
                $osTypeRaw = if ($device.backup_cloud_device_machine_os_type) { $device.backup_cloud_device_machine_os_type } else { "N/A" }
                $osType = Shorten-OSVersion $osTypeRaw
                
                # ConsoleLink and clickable device name (opens in same named tab)
                $consoleLink = if ($device.ConsoleLink) { $device.ConsoleLink } else { "#" }
                $deviceNameLink = "<a href='$consoleLink' target='cove-console' style='color: #0ea5e9; text-decoration: none; font-weight: 500;'>$deviceName</a>"
                
                $planName = if ($device.plan_name) { $device.plan_name } else { "N/A" }
                
                # Backup session timestamp
                $backupSession = if ($device.last_backup_session_timestamp) {
                    ([datetime]$device.last_backup_session_timestamp).ToString("MM/dd/yy HH:mm")
                } else { "N/A" }
                
                # Type display
                $typeDisplay = "🔄 Recovery Testing"
                
                # Recovery location (cross-reference agent_id)
                $recoveryLocationName = "N/A"
                if ($device.agent_id -and $Script:RecoveryLocations) {
                    $rlMatch = $Script:RecoveryLocations | Where-Object { $_.recovery_agent_id -eq $device.agent_id } | Select-Object -First 1
                    if ($rlMatch) { $recoveryLocationName = $rlMatch.name }
                }
                
                # 28-day colorbar
                $colorbarHtml = ""
                if ($device.colorbar_detail) {
                    $cbEntries = @($device.colorbar_detail -split ',' | Where-Object { $_ })
                    if ($cbEntries.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbEntry in $cbEntries[0..([Math]::Min($cbEntries.Count,28)-1)]) {
                            $cbParts = $cbEntry -split '\|'
                            if ($cbParts.Count -eq 2) {
                                $cbColor = switch ($cbParts[1]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                                $cbTz = if ($device.backup_cloud_device_timezone) { $device.backup_cloud_device_timezone } else { 'UTC' }
                                $colorbarHtml += "<div class='colorbar-block' data-timestamp='$($cbParts[0])' data-status='$($cbParts[1])' data-device-timezone='$cbTz' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                            }
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } elseif ($device.colorbar) {
                    $cbStatuses = @($device.colorbar -split ',' | Where-Object { $_ })
                    if ($cbStatuses.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbIdx in 0..([Math]::Min($cbStatuses.Count,28)-1)) {
                            $cbColor = switch ($cbStatuses[$cbIdx]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                            $colorbarHtml += "<div title='Day -${cbIdx}: $($cbStatuses[$cbIdx])' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                
                $recoveryTestingRows += "<tr><td>$deviceNameLink</td><td>$partnerName</td><td>$typeDisplay</td><td>$region</td><td><small>$osType</small></td><td><span style='color: $statusColor; font-weight: 600;'>$status</span></td><td>$backupSession</td><td>$lastRecovery</td><td>$duration</td><td>$restoredSize</td><td><strong style='color: #0097D6;'>$gbPerHour</strong></td><td><small>$planName</small></td><td>$recoveryLocationName</td><td>$colorbarHtml</td></tr>"
            }
            
            if ([string]::IsNullOrEmpty($recoveryTestingRows)) {
                $recoveryTestingRows = "<tr><td colspan='14' style='text-align: center; padding: 20px; color: #6b7280; font-style: italic;'>No recovery testing devices found</td></tr>"
            }
        } else {
            $recoveryTestingRows = "<tr><td colspan='14' style='text-align: center; padding: 20px; color: #6b7280; font-style: italic;'>No continuity data available</td></tr>"
        }
        
        # Build Standby Image (SBI only - SELF_HOSTED, AZURE_SELF_HOSTED, ESXI_SELF_HOSTED) table rows with screenshots
        $standbyImageRows = ""
        $Script:screenshotGallerySBI = ""
        $Script:screenshotGallerySBICompact = ""
        $Script:screenshotGalleryRT = ""
        $Script:screenshotGalleryRTCompact = ""
        $Script:screenshotGallery = ""
        $Script:screenshotGalleryCompact = ""
        
        # DEDUPLICATION: Track device IDs already added to galleries (prevents duplicate screenshots)
        # Same device can appear with multiple types (ESXI_SELF_HOSTED, SELF_HOSTED, AZURE_SELF_HOSTED)
        # Priority: Platform-specific types (ESXI_, AZURE_) over generic SELF_HOSTED
        $Script:addedToSBIGallery = @{}
        $Script:addedToRTGallery = @{}
        $Script:addedToOTRGallery = @{}
        $Script:addedToDRaaSGallery = @{}  # Added for DRaaS gallery deduplication
        
        # Use ONLY Continuity API data for Standby Image section (from 'Standby Image' Excel tab)
        # NO merging with Statistics API - show only devices with actual continuity data
        
        if ($SBIDevices -and @($SBIDevices).Count -gt 0) {
            $standbyImageData = @($SBIDevices | Sort-Object last_recovery_timestamp -Descending | Select-Object -First 50)
                
                Write-Host "  [DEBUG] standbyImageData count after limit: $($standbyImageData.Count)" -ForegroundColor Magenta
                Write-Host "  [DEBUG] About to enter foreach loop..." -ForegroundColor Magenta
                
                foreach ($device in $standbyImageData) {
                    $deviceName = "$($device.backup_cloud_device_name)"
                    $partnerName = $device.backup_cloud_partner_name
                    $deviceId = $device.backup_cloud_device_id
                    $type = $device.type
                    $typeDisplay = switch ($type) {
                        'SELF_HOSTED' { '💾 Self-Hosted' }
                        'AZURE_SELF_HOSTED' { '☁️ Azure SBI' }
                        'ESXI_SELF_HOSTED' { '🖥️ ESXi SBI' }
                        default { $type }
                    }
                    $targetType = $device.recovery_target_type
                    $status = $device.current_recovery_status
                    
                    # Add recovery progress to InProgress status
                    $statusDisplay = if ($status -eq 'InProgress' -and $device.recovery_session_progress) {
                        # Parse percentage from recovery_session_progress object or string
                        $progressPct = if ($device.recovery_session_progress -is [PSCustomObject] -and $device.recovery_session_progress.percentage) {
                            "$($device.recovery_session_progress.percentage)%"
                        } elseif ($device.recovery_session_progress -match '(\d+)%') {
                            "$($matches[1])%"
                        } else {
                            $device.recovery_session_progress
                        }
                        "InProgress ($progressPct)"
                    } else {
                        $status
                    }
                    
                    $statusColor = switch ($status) {
                        'Completed' { '#10b981' }
                        'Failed' { '#ef4444' }
                        'InProgress' { '#f59e0b' }
                        'NotStarted' { '#6b7280' }
                        default { '#6b7280' }
                    }
                    $lastRecovery = if ($device.last_recovery_timestamp) { 
                        ([datetime]$device.last_recovery_timestamp).ToString("MM/dd/yy HH:mm")
                    } else { 
                        "Never" 
                    }
                    
                    # Add backup session timestamp
                    $backupSession = if ($device.last_backup_session_timestamp) {
                        ([datetime]$device.last_backup_session_timestamp).ToString("MM/dd/yy HH:mm")
                    } else {
                        "N/A"
                    }
                    
                    # Calculate GB/HR throughput from last recovery
                    $gbPerHour = "N/A"
                    if ($device.last_recovery_restored_size_user -and $device.last_recovery_duration_user) {
                        # Parse duration string (e.g., "2h 30m", "45m", "1h 15m 30s")
                        $durationStr = $device.last_recovery_duration_user
                        $totalMinutes = 0
                        
                        if ($durationStr -match '(\d+)h') { $totalMinutes += [int]$matches[1] * 60 }
                        if ($durationStr -match '(\d+)m') { $totalMinutes += [int]$matches[1] }
                        if ($durationStr -match '(\d+)s' -and $totalMinutes -eq 0) { $totalMinutes = 1 }  # Round up seconds to 1 min minimum
                        
                        if ($totalMinutes -gt 0) {
                            $totalHours = $totalMinutes / 60.0
                            
                            # Parse size string (e.g., "15.23 GB", "1.5 TB", "500 MB")
                            $sizeStr = $device.last_recovery_restored_size_user
                            $sizeGB = 0
                            
                            if ($sizeStr -match '([\d.]+)\s*TB') {
                                $sizeGB = [decimal]$matches[1] * 1024
                            } elseif ($sizeStr -match '([\d.]+)\s*GB') {
                                $sizeGB = [decimal]$matches[1]
                            } elseif ($sizeStr -match '([\d.]+)\s*MB') {
                                $sizeGB = [decimal]$matches[1] / 1024
                            }
                            
                            if ($sizeGB -gt 0 -and $totalHours -gt 0) {
                                $throughput = $sizeGB / $totalHours
                                $gbPerHour = "{0:N2} GB/hr" -f $throughput
                            }
                        }
                    }
                    
                    # Cross-reference agent_id with recovery_agent_id to get recovery location name
                    $recoveryLocationName = "N/A"
                    if ($device.agent_id -and $Script:RecoveryLocations) {
                        $matchingLocation = $Script:RecoveryLocations | Where-Object { $_.recovery_agent_id -eq $device.agent_id } | Select-Object -First 1
                        if ($matchingLocation) {
                            $recoveryLocationName = $matchingLocation.name
                        }
                    }
                    
                    # Build 28-day colorbar (mimics Cove console visualization)
                    $colorbarHtml = ""
                    
                    # Try to use detailed colorbar data first (includes dates), fallback to status-only
                    if ($device.colorbar_detail) {
                        $colorbarEntries = @($device.colorbar_detail -split ',' | Where-Object { $_ })
                        if ($colorbarEntries.Count -gt 0) {
                            $colorbarHtml = "<div class='colorbar-container' style='display: flex; gap: 2px; align-items: center; min-width: 140px;'>"
                            $dayCount = [Math]::Min($colorbarEntries.Count, 28)
                            for ($i = 0; $i -lt $dayCount; $i++) {
                                $entry = $colorbarEntries[$i]
                                $parts = $entry -split '\|'
                                if ($parts.Count -eq 2) {
                                    $unixTimestamp = $parts[0]
                                    $sessionStatus = $parts[1]
                                    
                                    $blockColor = switch ($sessionStatus) {
                                        'Completed' { '#10b981' }
                                        'CompletedWithErrors' { '#f59e0b' }
                                        'Failed' { '#ef4444' }
                                        'InProgress' { '#3b82f6' }
                                        'Interrupted' { '#f97316' }
                                        'NotStarted' { '#d1d5db' }
                                        'Aborted' { '#9ca3af' }
                                        default { '#e5e7eb' }
                                    }
                                    
                                    # Store Unix timestamp and device timezone in data attributes for JavaScript conversion
                                    # This ensures correct date display in all timezone views (Local, UTC, Device)
                                    $deviceTimezone = if ($device.backup_cloud_device_timezone) { $device.backup_cloud_device_timezone } else { 'UTC' }
                                    $colorbarHtml += "<div class='colorbar-block' data-timestamp='$unixTimestamp' data-status='$sessionStatus' data-device-timezone='$deviceTimezone' style='width: 4px; height: 24px; background: $blockColor; border-radius: 1px; cursor: help;'></div>"
                                }
                            }
                            $colorbarHtml += "</div>"
                        } else {
                            $colorbarHtml = "<span style='color: #9ca3af; font-size: 11px;'>No data</span>"
                        }
                    } elseif ($device.colorbar) {
                        # Fallback to status-only colorbar
                        $colorbarStatuses = @($device.colorbar -split ',' | Where-Object { $_ })
                        if ($colorbarStatuses.Count -gt 0) {
                            $colorbarHtml = "<div style='display: flex; gap: 2px; align-items: center; min-width: 140px;'>"
                            $dayCount = [Math]::Min($colorbarStatuses.Count, 28)
                            for ($i = 0; $i -lt $dayCount; $i++) {
                                $sessionStatus = $colorbarStatuses[$i]
                                $blockColor = switch ($sessionStatus) {
                                    'Completed' { '#10b981' }
                                    'CompletedWithErrors' { '#f59e0b' }
                                    'Failed' { '#ef4444' }
                                    'InProgress' { '#3b82f6' }
                                    'Interrupted' { '#f97316' }
                                    'NotStarted' { '#d1d5db' }
                                    'Aborted' { '#9ca3af' }
                                    default { '#e5e7eb' }
                                }
                                $daysAgo = $i
                                $tooltipText = "Day -${daysAgo}: $sessionStatus"
                                $colorbarHtml += "<div title='$tooltipText' style='width: 4px; height: 24px; background: $blockColor; border-radius: 1px; cursor: help;'></div>"
                            }
                            $colorbarHtml += "</div>"
                        } else {
                            $colorbarHtml = "<span style='color: #9ca3af; font-size: 11px;'>No data</span>"
                        }
                    } else {
                        $colorbarHtml = "<span style='color: #9ca3af; font-size: 11px;'>No data</span>"
                    }
                    
                    # Region, OS Type, Duration, Restored Size, Plan
                    $region = if ($device.region_name) { $device.region_name } else { "N/A" }
                    $osTypeRaw = if ($device.OS) { $device.OS } elseif ($device.backup_cloud_device_machine_os_type) { $device.backup_cloud_device_machine_os_type } else { "N/A" }
                    $osType = Shorten-OSVersion $osTypeRaw
                    $duration = if ($device.last_recovery_duration_user) { $device.last_recovery_duration_user } else { "N/A" }
                    $restoredSize = if ($device.last_recovery_restored_size_user) { "$($device.last_recovery_restored_size_user) GB" } else { "N/A" }
                    $planName = if ($device.plan_name) { $device.plan_name } else { "N/A" }
                    
                    # Build console link for Standby Image device
                    $sbiConsoleLink = "https://backup.management/#/continuity/view/standby-image(panel:device-properties/$deviceId/summary)"
                    $deviceNameLink = "<a href='$sbiConsoleLink' target='_blank' style='color: #3b82f6; text-decoration: none; font-weight: 500;' onmouseover='this.style.textDecoration=``underline``' onmouseout='this.style.textDecoration=``none``'>$deviceName</a>"
                    
                    # Note: Screenshot galleries are now built separately after table building (see dedicated gallery builder below)
                    
                    $standbyImageRows += "<tr><td>$deviceNameLink</td><td>$partnerName</td><td>$typeDisplay</td><td>$region</td><td><small>$osType</small></td><td><span style='color: $statusColor; font-weight: 600;'>$statusDisplay</span></td><td>$backupSession</td><td>$lastRecovery</td><td>$duration</td><td>$restoredSize</td><td>$gbPerHour</td><td><small>$planName</small></td><td>$recoveryLocationName</td><td>$colorbarHtml</td></tr>"
                }
                
                Write-Host "  [DEBUG] After foreach - standbyImageRows length: $($standbyImageRows.Length)" -ForegroundColor Magenta
                Write-Host "  [DEBUG] standbyImageRows IsNullOrEmpty: $([string]::IsNullOrEmpty($standbyImageRows))" -ForegroundColor Magenta
                
            if ([string]::IsNullOrEmpty($standbyImageRows)) {
                $standbyImageRows = "<tr><td colspan='14' style='text-align: center; padding: 20px; color: #6b7280; font-style: italic;'>No standby image devices found</td></tr>"
            }
        } else {
            $standbyImageRows = "<tr><td colspan='14' style='text-align: center; padding: 20px; color: #6b7280; font-style: italic;'>No continuity data available</td></tr>"
        }
        
        # Build screenshot galleries from ALL available screenshots (not just first 20 table rows)
        # This ensures all downloaded/cached screenshots appear in the gallery sections
        # IMPORTANT: Match screenshots to device type to prevent using RT screenshots for SBI devices, etc.
        if ($Script:ScreenshotFiles -and $Script:ScreenshotFiles.Count -gt 0) {
            Write-Host "  Building screenshot galleries for $($Script:ScreenshotFiles.Count) devices..." -ForegroundColor Gray
            
            # Also load any cached screenshots not in $Script:ScreenshotFiles (from previous runs)
            # Match them to device records to get correct type and metadata
            $screenshotBaseDir = Join-Path $Script:ExportPath "screenshots"
            if (Test-Path $screenshotBaseDir) {
                $allScreenshotFiles = Get-ChildItem -Path $screenshotBaseDir -Filter "*.png" -Recurse
                $processedDeviceIds = @($Script:ScreenshotFiles.Keys)
                
                foreach ($screenshotFile in $allScreenshotFiles) {
                    # Parse filename: {deviceId}_{typeAbbrev}_{timestamp}.png
                    # Valid types: RT (Recovery Testing), SBI (Standby Image), DRaaS
                    if ($screenshotFile.Name -match '^(\d+)_(RT|SBI|DRaaS)_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\.png$') {
                        $fileDeviceId = [int]$matches[1]
                        $fileTypeAbbrev = $matches[2]
                        $fileTimestamp = $matches[3]
                        
                        # Skip if already processed in this run
                        if ($processedDeviceIds -contains $fileDeviceId) { continue }
                        
                        # Find matching device in continuity statistics with matching type
                        # Note: OTR screenshots should not exist (excluded from processing)
                        $matchingDevice = $Script:ContinuityStatistics | Where-Object { 
                            $_.backup_cloud_device_id -eq $fileDeviceId -and
                            (
                                ($fileTypeAbbrev -eq "RT" -and $_.type -eq "RECOVERY_TESTING") -or
                                ($fileTypeAbbrev -eq "SBI" -and $_.type -like "*SELF_HOSTED*") -or
                                ($fileTypeAbbrev -eq "DRaaS" -and $_.type -like "*DRAAS*")
                            )
                        } | Select-Object -First 1
                        
                        if ($matchingDevice) {
                            # Add to ScreenshotFiles hashtable
                            # Try to parse timestamp from filename, fall back to device's timestamp
                            $parsedTimestamp = $matchingDevice.last_boot_test_backup_session_timestamp
                            try {
                                $parsedTimestamp = [DateTime]::ParseExact($fileTimestamp, "yyyy-MM-dd_HH-mm-ss", $null)
                            } catch {
                                # Use device timestamp as fallback
                            }
                            
                            # Try to get full OS version and TimeZone from DeviceDetail, fall back to device type
                            $deviceOSRaw = $matchingDevice.backup_cloud_device_machine_os_type
                            $deviceTimeZone = "UTC"  # Default fallback
                            if ($Script:DeviceDetail) {
                                $fullDeviceData = $Script:DeviceDetail | Where-Object { $_.DeviceId -eq $fileDeviceId } | Select-Object -First 1
                                if ($fullDeviceData -and $fullDeviceData.OS) {
                                    $deviceOSRaw = $fullDeviceData.OS
                                }
                                if ($fullDeviceData -and $fullDeviceData.TimeZone) {
                                    $deviceTimeZone = $fullDeviceData.TimeZone
                                }
                            }
                            
                            $deviceOS = Shorten-OSVersion $deviceOSRaw
                            
                            # Use composite key: DeviceId_Type_SessionId to prevent duplicate device IDs with same session ID but different types
                            $compositeKey = if ($matchingDevice.last_recovery_session_id) {
                                "$fileDeviceId`_$($matchingDevice.type)`_$($matchingDevice.last_recovery_session_id)"
                            } else {
                                "$fileDeviceId`_$($matchingDevice.type)"  # Fallback to DeviceId_Type
                            }
                            
                            # Get type subdirectory from file's parent folder name
                            $typeSubdir = $screenshotFile.Directory.Name
                            
                            $Script:ScreenshotFiles[$compositeKey] = @{
                                Path = "screenshots/$typeSubdir/$($screenshotFile.Name)"
                                DeviceName = $matchingDevice.backup_cloud_device_name
                                DeviceType = $matchingDevice.type
                                RecoveryTargetType = $matchingDevice.recovery_target_type
                                Status = $matchingDevice.last_boot_test_status
                                Timestamp = $parsedTimestamp
                                PartnerName = $matchingDevice.backup_cloud_partner_name
                                OS = $deviceOS
                                TimeZone = $deviceTimeZone
                                BackupSessionTime = $matchingDevice.last_boot_test_backup_session_timestamp
                                RecoverySessionTime = $matchingDevice.last_boot_test_recovery_session_timestamp
                                RecoveryDuration = $matchingDevice.last_recovery_duration_user
                                BootSchedule = $matchingDevice.plan_name
                            }
                        }
                    }
                }
            }
            
            Write-Host "  Building screenshot galleries for $($Script:ScreenshotFiles.Count) devices (including cached)..." -ForegroundColor Gray
            
            # Sort screenshot gallery to match device table order (by BackupSessionTime descending)
            # Convert hashtable to array of key-value pairs and sort by BackupSessionTime (most recent first)
            $sortedScreenshots = $Script:ScreenshotFiles.GetEnumerator() | Sort-Object { 
                if ($_.Value.BackupSessionTime) { 
                    [datetime]$_.Value.BackupSessionTime 
                } else { 
                    [datetime]::MinValue 
                }
            } -Descending
            
            foreach ($entry in $sortedScreenshots) {
                $deviceId = $entry.Key
                $screenshotData = $entry.Value
                
                # Parse device timezone offset for calculations
                # Handles formats: "UTC +11", "UTC -5", "UTC+11", "UTC-5", "UTC"
                $deviceTzOffset = 0
                if ($screenshotData.TimeZone -match 'UTC\s*([+-]?\d+)') {
                    $deviceTzOffset = [int]$matches[1]
                }
                
                # Format timestamps in all three views: Local, UTC, Device
                # Backup Session Time
                $backupTimeLocal = $backupTimeUTC = $backupTimeDevice = "N/A"
                if ($screenshotData.BackupSessionTime) {
                    $backupDT = [datetime]$screenshotData.BackupSessionTime
                    $backupTimeLocal = $backupDT.ToString("MMM dd HH:mm")
                    $backupTimeUTC = $backupDT.ToUniversalTime().ToString("MMM dd HH:mm")
                    $backupTimeDevice = $backupDT.ToUniversalTime().AddHours($deviceTzOffset).ToString("MMM dd HH:mm")
                }
                
                # Recovery Test Time
                $recoveryTimeLocal = $recoveryTimeUTC = $recoveryTimeDevice = "N/A"
                if ($screenshotData.RecoverySessionTime) {
                    $recoveryDT = [datetime]$screenshotData.RecoverySessionTime
                    $recoveryTimeLocal = $recoveryDT.ToString("MMM dd HH:mm")
                    $recoveryTimeUTC = $recoveryDT.ToUniversalTime().ToString("MMM dd HH:mm")
                    $recoveryTimeDevice = $recoveryDT.ToUniversalTime().AddHours($deviceTzOffset).ToString("MMM dd HH:mm")
                }
                
                # Boot Time (matches Recovery Time)
                $bootTimeLocal = $recoveryTimeLocal
                $bootTimeUTC = $recoveryTimeUTC
                $bootTimeDevice = $recoveryTimeDevice
                
                $duration = if ($screenshotData.RecoveryDuration) { $screenshotData.RecoveryDuration } else { "N/A" }
                # Extract OS version from screenshot metadata (already shortened by Shorten-OSVersion in cache building)
                $osVersion = if ($screenshotData.OS) { $screenshotData.OS } else { "Unknown" }
                $partner = if ($screenshotData.PartnerName) { $screenshotData.PartnerName } else { "N/A" }
                
                # Timezone is already formatted as "UTC +X" from DeviceDetail.TimeZone
                # We only have the offset, not the timezone name, so display offset only
                $timezoneOffset = if ($screenshotData.TimeZone) { $screenshotData.TimeZone } else { "UTC" }
                
                # Device type display with icon and color - categorize based on type field
                $deviceType = if ($screenshotData.DeviceType) { $screenshotData.DeviceType } else { "UNKNOWN" }
                
                # Categorize based on type field content (prioritize SELF_HOSTED over ON_DEMAND for dual-type devices like SELF_HOSTED_ON_DEMAND)
                if ($deviceType -like "*DRAAS*") {
                    $typeCategory = "DRaaS"
                    $typeIcon = "🛡️"
                    $typeColor = "#8b5cf6"
                } elseif ($deviceType -like "*SELF_HOSTED*") {
                    $typeCategory = "Standby Image"
                    $typeIcon = "💾"
                    $typeColor = "#3b82f6"
                } elseif ($deviceType -like "*ON_DEMAND*") {
                    $typeCategory = "One-Time Restore"
                    $typeIcon = "⚡"
                    $typeColor = "#f59e0b"
                } elseif ($deviceType -eq "RECOVERY_TESTING") {
                    $typeCategory = "Recovery Testing"
                    $typeIcon = @"
<svg width="14" height="14" viewBox="0 0 16 16" fill="none" style="vertical-align: middle; margin-right: 2px;">
  <circle cx="8" cy="8" r="7" fill="white" stroke="currentColor" stroke-width="1.5"/>
  <path d="M5 8.5L7 10.5L11 6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
</svg>
"@
                    $typeColor = "#10b981"
                } else {
                    $typeCategory = $deviceType
                    $typeIcon = "❓"
                    $typeColor = "#6b7280"
                }
                
                # Platform-specific logo badges - prefer RecoveryTargetType (Hyper-V, VMware, etc.) over generic DeviceType
                $recoveryTarget = if ($screenshotData.RecoveryTargetType) { $screenshotData.RecoveryTargetType } else { "" }
                if ($recoveryTarget -like "*Hyper-V*") {
                    $typePlatform = "Hyper-V"
                    $platformLogo = @"
<svg width="20" height="20" viewBox="0 0 20 20" fill="none" style="vertical-align: middle; margin-right: 4px;">
  <rect width="20" height="20" rx="4" fill="white"/>
  <rect x="5" y="5" width="4" height="10" fill="`#0078D4"/>
  <rect x="11" y="5" width="4" height="10" fill="`#50E6FF"/>
  <rect x="8" y="8" width="4" height="4" fill="`#0078D4" opacity="0.6"/>
</svg>
"@
                } elseif ($recoveryTarget -like "*VMware*" -or $recoveryTarget -like "*ESXi*") {
                    $typePlatform = "VMware"
                    $platformLogo = '<img src="https://tse1.mm.bing.net/th/id/OIP.gmFCC6T3aCniqCDLOPmykwHaCk?pid=ImgDet&w=60&h=60&c=7&rs=1&o=7&rm=3" width="20" height="20" style="vertical-align: middle; margin-right: 4px; border-radius: 4px;">'
                } elseif ($recoveryTarget -like "*Azure*" -or $deviceType -like "*AZURE*") {
                    $typePlatform = "Azure"
                    $platformLogo = @"
<svg width="20" height="20" viewBox="0 0 20 20" fill="none" style="vertical-align: middle; margin-right: 4px;">
  <rect width="20" height="20" rx="4" fill="white"/>
  <path d="M8.5 4L6 10h4l-2.5 6L14 8h-4l2.5-4z" fill="`#0078D4"/>
</svg>
"@
                } elseif ($recoveryTarget -like "*Local VHD*" -or $recoveryTarget -like "*.VHD*" -or $recoveryTarget -like "*.VHDX*") {
                    $typePlatform = "Local VHD"
                    $platformLogo = @"
<svg width="20" height="20" viewBox="0 0 20 20" fill="none" style="vertical-align: middle; margin-right: 4px;">
  <rect width="20" height="20" rx="4" fill="white"/>
  <rect x="4" y="6" width="12" height="8" rx="1" fill="`#10b981" stroke="`#059669" stroke-width="1"/>
  <circle cx="7" cy="10" r="0.8" fill="white"/>
  <rect x="9" y="9.2" width="4" height="1.6" rx="0.3" fill="white"/>
</svg>
"@
                } else {
                    $typePlatform = "N-able Cove"
                    $platformLogo = @"
<svg width="20" height="20" viewBox="0 0 20 20" fill="none" style="vertical-align: middle; margin-right: 4px;">
  <rect width="20" height="20" rx="4" fill="white"/>
  <path d="M6 6h3v8H6V6z" fill="`#0066CC"/>
  <path d="M11 6h3v8h-3V6z" fill="`#00AEEF"/>
</svg>
"@
                }
                
                $typeText = if ($typePlatform -ne "N-able Cove") { "$typeCategory" } else { $typeCategory }
                $platformBadge = "$platformLogo<span style='font-weight: 600;'>$typePlatform</span>"
                
                # Build Cove console link - use device type from screenshot data
                $deviceType = if ($screenshotData.DeviceType) { $screenshotData.DeviceType } else { "RECOVERY_TESTING" }
                $consoleLink = if ($deviceType -like "*SELF_HOSTED*" -or $deviceType -like "*AZURE*" -or $deviceType -like "*ESXI*") {
                    "https://backup.management/#/continuity/view/standby-image(panel:device-properties/$deviceId/recovery-verification/$deviceType)"
                } else {
                    "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/RECOVERY_TESTING)"
                }
                
                # Determine border color based on boot test status
                $borderColor = switch ($screenshotData.Status) {
                    'SUCCESS' { '#10b981' }  # Green
                    'FAILED' { '#ef4444' }   # Red
                    default { '#f59e0b' }    # Orange/Yellow
                }
                
                # Extract actual numeric device ID from composite key (format: DeviceId_Type_SessionId)
                $actualDeviceId = ($deviceId -split '_')[0]
                
                # CRITICAL FIX: Create unique key using device ID + device type to allow multiple recovery types per device
                # Example: device 5033463 with BOTH Hyper-V and ESXi should show 2 separate cards
                # Old key: "5033463" (only shows one card - first type found)
                # New key: "5033463_ESXI_SELF_HOSTED" and "5033463_SELF_HOSTED" (shows both cards)
                $uniqueGalleryKey = "${actualDeviceId}_${deviceType}"
                
                # Determine console link based on device type (URL format: device-properties/{deviceId}/recovery-verification/{TYPE})
                $consoleTypeParam = if ($deviceType -eq "RECOVERY_TESTING") { "RECOVERY_TESTING" } elseif ($deviceType -like "*SELF_HOSTED*") { "STANDBY_IMAGE" } else { "RECOVERY_TESTING" }
                $consoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$actualDeviceId/recovery-verification/$consoleTypeParam)"
                
                # Add to type-specific galleries based on category
                $compactCard = @"
<div class='screenshot-compact-card' data-device-id='$deviceId' data-device-name='$($screenshotData.DeviceName)' data-partner='$partner' data-os='$osVersion' data-timezone-offset='$timezoneOffset' data-backup-local='$backupTimeLocal' data-backup-utc='$backupTimeUTC' data-backup-device='$backupTimeDevice' data-recovery-local='$recoveryTimeLocal' data-recovery-utc='$recoveryTimeUTC' data-recovery-device='$recoveryTimeDevice' data-boot-local='$bootTimeLocal' data-boot-utc='$bootTimeUTC' data-boot-device='$bootTimeDevice' data-duration='$duration' data-status='$($screenshotData.Status)' data-image-path='$($screenshotData.Path)' style='background: #1f2937; border-radius: 8px; overflow: hidden; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s;' onclick="openScreenshotModal('$deviceId')" onmouseenter='this.style.transform="translateY(-4px)"'; this.style.boxShadow="0 8px 16px rgba(0,0,0,0.4)"' onmouseleave='this.style.transform=""; this.style.boxShadow=""'>
    <div style='position: relative;'>
        <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; border: 5px solid $borderColor;' alt='Boot Screenshot' loading='lazy'>
        <div style='position: absolute; top: 8px; left: 8px; background: rgba(31,41,55,0.95); color: white; padding: 6px 10px; border-radius: 12px; font-size: 11px; box-shadow: 0 2px 8px rgba(0,0,0,0.4); backdrop-filter: blur(4px); display: flex; align-items: center;'>$platformBadge</div>
        <div style='position: absolute; bottom: 8px; right: 8px; background: $typeColor; color: white; padding: 4px 8px; border-radius: 8px; font-size: 9px; font-weight: 700; box-shadow: 0 2px 4px rgba(0,0,0,0.3);'>$typeIcon $typeCategory</div>
    </div>
    <div style='padding: 8px; text-align: center;'>
        <a href='$consoleLink' target='_blank' style='color: #60a5fa; text-decoration: none; font-size: 11px; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display: inline-flex; align-items: center; gap: 4px; justify-content: center; width: 100%;' onclick='event.stopPropagation();'>
            $($screenshotData.DeviceName)
            <svg width='12' height='12' viewBox='0 0 16 16' fill='currentColor' style='flex-shrink: 0;'>
                <path d='M6.22 8.72a.75.75 0 0 0 1.06 1.06l5.22-5.22v1.69a.75.75 0 0 0 1.5 0v-3.5a.75.75 0 0 0-.75-.75h-3.5a.75.75 0 0 0 0 1.5h1.69L6.22 8.72z'/>
                <path d='M3.5 6.75c0-.69.56-1.25 1.25-1.25H7A.75.75 0 0 0 7 4H4.75A2.75 2.75 0 0 0 2 6.75v4.5A2.75 2.75 0 0 0 4.75 14h4.5A2.75 2.75 0 0 0 12 11.25V9a.75.75 0 0 0-1.5 0v2.25c0 .69-.56 1.25-1.25 1.25h-4.5c-.69 0-1.25-.56-1.25-1.25v-4.5z'/>
            </svg>
        </a>
    </div>
</div>
"@
                
                # CRITICAL FIX: Match screenshots to galleries by FILENAME type abbreviation (RT/SBI/DRaaS)
                # NOT by current API device type - prevents SBI screenshots appearing in DRaaS gallery
                # when a device changes types between sessions
                # Extract type abbreviation from screenshot path: screenshots/{type}/{deviceId}_{typeAbbrev}_{timestamp}.png
                $screenshotTypeAbbrev = ""
                if ($screenshotData.Path -match '_(RT|SBI|DRaaS)_') {
                    $screenshotTypeAbbrev = $matches[1]
                }
                
                # Filter by screenshot filename type (RT/SBI/DRaaS), not current device type
                # DEDUPLICATION: Use unique key (device ID + type) to allow multiple recovery types per device
                if ($screenshotTypeAbbrev -eq "SBI") {
                    # Standby Image screenshot - add to SBI gallery only
                    if (-not $Script:addedToSBIGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGallerySBICompact += $compactCard
                        $Script:addedToSBIGallery[$uniqueGalleryKey] = $deviceType  # Track device type used
                    }
                    # No else block needed - each unique key (device+type) is independent
                } elseif ($screenshotTypeAbbrev -eq "RT") {
                    # Recovery Testing screenshot - add to RT gallery only
                    if (-not $Script:addedToRTGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryRTCompact += $compactCard
                        $Script:addedToRTGallery[$uniqueGalleryKey] = $deviceType
                    }
                } elseif ($screenshotTypeAbbrev -eq "DRaaS") {
                    # DRaaS screenshot - add to DRaaS gallery only
                    if (-not $Script:addedToDRaaSGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryDRaaSCompact += $compactCard
                        $Script:addedToDRaaSGallery[$uniqueGalleryKey] = $deviceType
                    }
                } elseif ($deviceType -like "*ON_DEMAND*" -or $deviceType -eq "AZURE" -or $deviceType -eq "ESXI_ON_DEMAND" -or $deviceType -eq "SELF_HOSTED_ON_DEMAND") {
                    # OTR uses device type (no screenshots expected for OTR)
                    if (-not $Script:addedToOTRGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryOTRCompact += $compactCard
                        $Script:addedToOTRGallery[$uniqueGalleryKey] = $deviceType
                    }
                }
                $Script:screenshotGalleryCompact += $compactCard
                
                # Now build the detailed card
                $detailedCard = @"
<div id='screenshot-$deviceId' class='screenshot-tile' data-device-id='$deviceId' data-device-name='$($screenshotData.DeviceName)' data-partner='$partner' data-os='$osVersion' data-timezone-offset='$timezoneOffset' data-backup-local='$backupTimeLocal' data-backup-utc='$backupTimeUTC' data-backup-device='$backupTimeDevice' data-recovery-local='$recoveryTimeLocal' data-recovery-utc='$recoveryTimeUTC' data-recovery-device='$recoveryTimeDevice' data-boot-local='$bootTimeLocal' data-boot-utc='$bootTimeUTC' data-boot-device='$bootTimeDevice' data-duration='$duration' data-status='$($screenshotData.Status)' data-image-path='$($screenshotData.Path)'>
    <div style='background: #1f2937; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);'>
        <div style='position: relative;'>
            <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; cursor: pointer; border: 5px solid $borderColor;' alt='Boot Screenshot' loading='lazy' onclick="openScreenshotModal('$deviceId')">
            <div style='position: absolute; top: 10px; left: 10px; background: rgba(31,41,55,0.95); color: white; padding: 8px 12px; border-radius: 16px; font-size: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.5); backdrop-filter: blur(4px); display: flex; align-items: center;'>$platformBadge</div>
            <div style='position: absolute; bottom: 10px; right: 10px; background: $typeColor; color: white; padding: 6px 12px; border-radius: 12px; font-size: 11px; font-weight: 700; box-shadow: 0 4px 8px rgba(0,0,0,0.4);'>$typeIcon $typeCategory</div>
        </div>
        <div style='padding: 14px; color: #e5e7eb; font-size: 12px; background: #111827; line-height: 1.6;'>
            <div style='font-weight: 600; font-size: 13px; margin-bottom: 10px; text-align: center; border-bottom: 1px solid #374151; padding-bottom: 8px;'>
                <a href='$consoleLink' target='_blank' style='color: #60a5fa; text-decoration: none; display: inline-flex; align-items: center; gap: 4px;'>
                    $($screenshotData.DeviceName)
                    <svg width='14' height='14' viewBox='0 0 16 16' fill='currentColor' style='flex-shrink: 0;'>
                        <path d='M6.22 8.72a.75.75 0 0 0 1.06 1.06l5.22-5.22v1.69a.75.75 0 0 0 1.5 0v-3.5a.75.75 0 0 0-.75-.75h-3.5a.75.75 0 0 0 0 1.5h1.69L6.22 8.72z'/>
                        <path d='M3.5 6.75c0-.69.56-1.25 1.25-1.25H7A.75.75 0 0 0 7 4H4.75A2.75 2.75 0 0 0 2 6.75v4.5A2.75 2.75 0 0 0 4.75 14h4.5A2.75 2.75 0 0 0 12 11.25V9a.75.75 0 0 0-1.5 0v2.25c0 .69-.56 1.25-1.25 1.25h-4.5c-.69 0-1.25-.56-1.25-1.25v-4.5z'/>
                    </svg>
                </a>
            </div>
            <div style='display: grid; grid-template-columns: auto 1fr; gap: 6px 10px; font-size: 11px;'>
                <span style='color: #9ca3af; font-weight: 500;'>Partner:</span>
                <span style='color: #d1d5db; overflow: hidden; text-overflow: ellipsis;'>$partner</span>
                <span style='color: #9ca3af; font-weight: 500;'>OS:</span>
                <span style='color: #d1d5db; overflow: hidden; text-overflow: ellipsis;'>$osVersion</span>
                <span style='color: #9ca3af; font-weight: 500;'>Timezone:</span>
                <span style='color: #a78bfa; font-size: 11px;'>$timezoneOffset</span>
            </div>
            <div style='border-top: 1px solid #374151; margin: 4px 0;'></div>
            <div style='display: flex; justify-content: space-between;'>
                <span style='color: #9ca3af; font-weight: 500;'>Backup Session:</span>
                <span style='color: #60a5fa;' class='time-local-$deviceId'>$backupTimeLocal</span>
                <span style='color: #60a5fa; display: none;' class='time-utc-$deviceId'>$backupTimeUTC</span>
                <span style='color: #60a5fa; display: none;' class='time-device-$deviceId'>$backupTimeDevice</span>
            </div>
            <div style='display: flex; justify-content: space-between;'>
                <span style='color: #9ca3af; font-weight: 500;'>Recovery Test:</span>
                <span style='color: #10b981;' class='time-local-$deviceId'>$recoveryTimeLocal</span>
                <span style='color: #10b981; display: none;' class='time-utc-$deviceId'>$recoveryTimeUTC</span>
                <span style='color: #10b981; display: none;' class='time-device-$deviceId'>$recoveryTimeDevice</span>
            </div>
            <div style='display: flex; justify-content: space-between;'>
                <span style='color: #9ca3af; font-weight: 500;'>Boot Time:</span>
                <span style='color: #f59e0b;' class='time-local-$deviceId'>$bootTimeLocal</span>
                <span style='color: #f59e0b; display: none;' class='time-utc-$deviceId'>$bootTimeUTC</span>
                <span style='color: #f59e0b; display: none;' class='time-device-$deviceId'>$bootTimeDevice</span>
            </div>
            <div style='display: flex; justify-content: space-between;'>
                <span style='color: #9ca3af; font-weight: 500;'>Duration:</span>
                <span style='color: #e5e7eb;'>$duration</span>
            </div>
            <div style='display: flex; justify-content: space-between;'>
                <span style='color: #9ca3af; font-weight: 500;'>Boot Schedule:</span>
                <span style='color: #fbbf24; font-size: 10px; font-weight: 500;'>$(if ($screenshotData.BootSchedule) { $screenshotData.BootSchedule } else { 'Every Session' })</span>
            </div>
            <div style='border-top: 1px solid #374151; margin: 4px 0;'></div>
            <div style='text-align: center; $(if ($screenshotData.IsPriorSession) { "color: #f59e0b;" } else { "color: #10b981;" }) font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px;'>$($screenshotData.Status)</div>
        </div>
    </div>
</div>
"@
                
                # CRITICAL FIX: Match detailed cards to galleries by FILENAME type abbreviation (RT/SBI/DRaaS)
                # Must match the logic used for compact cards above
                # Extract type abbreviation from screenshot path
                $screenshotTypeAbbrev = ""
                if ($screenshotData.Path -match '_(RT|SBI|DRaaS)_') {
                    $screenshotTypeAbbrev = $matches[1]
                }
                
                # DEDUPLICATION: Use same tracking as compact cards (prevents duplicates)
                if ($screenshotTypeAbbrev -eq "SBI") {
                    # Only add if compact card was added (checked during compact card processing)
                    # This ensures compact and detailed galleries stay in sync
                    if ($Script:addedToSBIGallery.ContainsKey($uniqueGalleryKey)) {
                        # This device+type was added to compact gallery, add to detailed too
                        $Script:screenshotGallerySBI += $detailedCard
                    }
                } elseif ($screenshotTypeAbbrev -eq "RT") {
                    if ($Script:addedToRTGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryRT += $detailedCard
                    }
                } elseif ($screenshotTypeAbbrev -eq "DRaaS") {
                    if ($Script:addedToDRaaSGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryDRaaS += $detailedCard
                    }
                } elseif ($deviceType -like "*ON_DEMAND*" -or $deviceType -eq "AZURE" -or $deviceType -eq "ESXI_ON_DEMAND" -or $deviceType -eq "SELF_HOSTED_ON_DEMAND") {
                    if ($Script:addedToOTRGallery.ContainsKey($uniqueGalleryKey)) {
                        $Script:screenshotGalleryOTR += $detailedCard
                    }
                }
                $Script:screenshotGallery += $detailedCard
            }
        }
        
        # Build One-Time Restore table rows with screenshots
        $oneTimeRestoreRows = ""
        $Script:screenshotGalleryOTR = ""
        $Script:screenshotGalleryOTRCompact = ""
        if ($OneTimeRestoreDevices -and @($OneTimeRestoreDevices).Count -gt 0) {
            $oneTimeRestoreData = @($OneTimeRestoreDevices | Sort-Object last_recovery_timestamp -Descending)
            
            foreach ($device in $oneTimeRestoreData) {
                $deviceName = "$($device.backup_cloud_device_name)"
                $partnerName = $device.backup_cloud_partner_name
                $deviceId = $device.backup_cloud_device_id
                $type = $device.type
                $typeDisplay = switch ($type) {
                    'AZURE' { '☁️ Azure' }
                    'ESXI_ON_DEMAND' { '🖥️ ESXi On-Demand' }
                    'SELF_HOSTED_ON_DEMAND' { '💾 Self-Hosted On-Demand' }
                    default { $type }
                }
                $status = $device.current_recovery_status
                $statusColor = switch ($status) {
                    'Completed' { '#10b981' }
                    'Failed' { '#ef4444' }
                    'InProgress' { '#f59e0b' }
                    default { '#6b7280' }
                }
                $lastRecovery = if ($device.last_recovery_timestamp) { 
                    ([datetime]$device.last_recovery_timestamp).ToString("MM/dd/yy HH:mm")
                } else { 
                    "Never" 
                }
                
                # Check if device has screenshot available
                if ($device.last_boot_test_screenshot_presented -eq $true) {
                    # Build composite key matching screenshot storage format (DeviceId_Type_SessionId or DeviceId_Type)
                    $compositeKey = if ($device.last_recovery_session_id) {
                        "$deviceId`_$($device.type)`_$($device.last_recovery_session_id)"
                    } else {
                        "$deviceId`_$($device.type)"
                    }
                    
                    if ($Script:ScreenshotFiles -and $Script:ScreenshotFiles.ContainsKey($compositeKey)) {
                        $screenshotData = $Script:ScreenshotFiles[$compositeKey]
                        
                        $deviceTzOffset = 0
                        if ($screenshotData.TimeZone -match 'UTC([+-]\d+)') {
                            $deviceTzOffset = [int]$matches[1]
                        }
                        
                        $recoveryTimeLocal = $recoveryTimeUTC = $recoveryTimeDevice = "N/A"
                        if ($screenshotData.RecoverySessionTime) {
                            $recoveryDT = [datetime]$screenshotData.RecoverySessionTime
                            $recoveryTimeLocal = $recoveryDT.ToString("MMM dd HH:mm")
                            $recoveryTimeUTC = $recoveryDT.ToUniversalTime().ToString("MMM dd HH:mm")
                            $recoveryTimeDevice = $recoveryDT.ToUniversalTime().AddHours($deviceTzOffset).ToString("MMM dd HH:mm")
                        }
                        
                        $borderColor = switch ($screenshotData.Status) {
                            'SUCCESS' { '#10b981' }
                            'FAILED' { '#ef4444' }
                            default { '#f59e0b' }
                        }
                        
                        $consoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/ONE_TIME_RESTORE)"
                        
                        $Script:screenshotGalleryOTRCompact += @"
<div class='screenshot-compact-card' style='background: #1f2937; border-radius: 8px; overflow: hidden; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s;' onclick="openScreenshotModal('$deviceId')" onmouseenter='this.style.transform="translateY(-4px)"'; this.style.boxShadow="0 8px 16px rgba(0,0,0,0.4)"' onmouseleave='this.style.transform=""; this.style.boxShadow=""'>
    <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; border: 3px solid $borderColor;' alt='Recovery Screenshot' loading='lazy'>
    <div style='padding: 10px; text-align: center; font-weight: 600; font-size: 13px; color: #60a5fa; background: #111827;'>$($screenshotData.DeviceName)</div>
</div>
"@

                        $Script:screenshotGalleryOTR += @"
<div id='screenshot-otr-$deviceId' class='screenshot-tile' data-device-id='$deviceId' data-device-name='$($screenshotData.DeviceName)' data-partner='$($screenshotData.PartnerName)' data-image-path='$($screenshotData.Path)'>
    <div style='background: #1f2937; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);'>
        <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; cursor: pointer; border: 3px solid $borderColor;' alt='Recovery Screenshot' loading='lazy' onclick="openScreenshotModal('$deviceId')">
        <div style='padding: 14px; color: #e5e7eb; font-size: 12px; background: #111827; line-height: 1.6;'>
            <div style='font-weight: 700; margin-bottom: 8px; font-size: 14px;'>
                <a href='$consoleLink' target='_blank' style='color: #60a5fa; text-decoration: none;'>$($screenshotData.DeviceName)</a>
            </div>
            <div style='color: #a78bfa; font-size: 11px; margin-bottom: 8px;'>$typeDisplay</div>
            <div style='display: flex; justify-content: space-between; color: #9ca3af;'>
                <span>Recovery:</span>
                <span class='time-local-$deviceId' style='color: #10b981;'>$recoveryTimeLocal</span>
                <span class='time-utc-$deviceId' style='color: #10b981; display: none;'>$recoveryTimeUTC</span>
                <span class='time-device-$deviceId' style='color: #10b981; display: none;'>$recoveryTimeDevice</span>
            </div>
        </div>
    </div>
</div>
"@
                    }
                }
                
                $duration = if ($device.last_recovery_duration_user) { $device.last_recovery_duration_user } else { "N/A" }
                
                # Use last_recovery_restored_size_user from Excel (already in GB)
                $restoredSize = if ($device.last_recovery_restored_size_user) { "$($device.last_recovery_restored_size_user) GB" } else { "N/A" }
                
                # Calculate GB/hr metric (data already in GB from Excel)
                $gbPerHour = "N/A"
                if ($device.last_recovery_restored_size_user -and $device.last_recovery_duration_seconds) {
                    try {
                        $sizeGB = [decimal]$device.last_recovery_restored_size_user  # Already in GB
                        $durationHours = [decimal]$device.last_recovery_duration_seconds / 3600
                        if ($durationHours -gt 0) {
                            $gbPerHourValue = [Math]::Round($sizeGB / $durationHours, 2)
                            $gbPerHour = "$gbPerHourValue"
                        }
                    } catch {
                        $gbPerHour = "N/A"
                    }
                }
                
                # Region, OS Type, Backup Session, Plan, Recovery Location, Colorbar
                $region = if ($device.region_name) { $device.region_name } else { "N/A" }
                $osTypeRaw = if ($device.backup_cloud_device_machine_os_type) { $device.backup_cloud_device_machine_os_type } else { "N/A" }
                $osType = Shorten-OSVersion $osTypeRaw
                $backupSession = if ($device.last_backup_session_timestamp) { ([datetime]$device.last_backup_session_timestamp).ToString("MM/dd/yy HH:mm") } else { "N/A" }
                $planName = if ($device.plan_name) { $device.plan_name } else { "N/A" }
                $recoveryLocationName = "N/A"
                if ($device.agent_id -and $Script:RecoveryLocations) {
                    $rlMatch = $Script:RecoveryLocations | Where-Object { $_.recovery_agent_id -eq $device.agent_id } | Select-Object -First 1
                    if ($rlMatch) { $recoveryLocationName = $rlMatch.name }
                }
                $colorbarHtml = ""
                if ($device.colorbar_detail) {
                    $cbEntries = @($device.colorbar_detail -split ',' | Where-Object { $_ })
                    if ($cbEntries.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbEntry in $cbEntries[0..([Math]::Min($cbEntries.Count,28)-1)]) {
                            $cbParts = $cbEntry -split '\|'
                            if ($cbParts.Count -eq 2) {
                                $cbColor = switch ($cbParts[1]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                                $colorbarHtml += "<div data-timestamp='$($cbParts[0])' data-status='$($cbParts[1])' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                            }
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } elseif ($device.colorbar) {
                    $cbStatuses = @($device.colorbar -split ',' | Where-Object { $_ })
                    if ($cbStatuses.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbIdx in 0..([Math]::Min($cbStatuses.Count,28)-1)) {
                            $cbColor = switch ($cbStatuses[$cbIdx]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                            $colorbarHtml += "<div title='Day -${cbIdx}: $($cbStatuses[$cbIdx])' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                
                # Build console link for One-Time Restore device
                $otrConsoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/ONE_TIME_RESTORE)"
                $deviceNameLink = "<a href='$otrConsoleLink' target='_blank' style='color: #3b82f6; text-decoration: none; font-weight: 500;' onmouseover='this.style.textDecoration=``underline``' onmouseout='this.style.textDecoration=``none``'>$deviceName</a>"
                
                $oneTimeRestoreRows += "<tr><td>$deviceNameLink</td><td>$partnerName</td><td>$typeDisplay</td><td>$region</td><td><small>$osType</small></td><td><span style='color: $statusColor; font-weight: 600;'>$status</span></td><td>$backupSession</td><td>$lastRecovery</td><td>$duration</td><td>$restoredSize</td><td style='color: #3b82f6; font-weight: 600;'>$gbPerHour</td><td><small>$planName</small></td><td>$recoveryLocationName</td><td>$colorbarHtml</td></tr>"
            }
        }
        
        # Build DRaaS table rows with screenshots
        $draasRows = ""
        $Script:screenshotGalleryDRaaS = ""
        $Script:screenshotGalleryDRaaSCompact = ""
        if ($DRaaSDevices -and @($DRaaSDevices).Count -gt 0) {
            $draasData = @($DRaaSDevices | Sort-Object last_recovery_timestamp -Descending)
            
            foreach ($device in $draasData) {
                $deviceName = "$($device.backup_cloud_device_name)"
                $partnerName = $device.backup_cloud_partner_name
                $deviceId = $device.backup_cloud_device_id
                $status = $device.current_recovery_status
                $statusColor = switch ($status) {
                    'Completed' { '#10b981' }
                    'Failed' { '#ef4444' }
                    'InProgress' { '#f59e0b' }
                    default { '#6b7280' }
                }
                $lastRecovery = if ($device.last_recovery_timestamp) { 
                    ([datetime]$device.last_recovery_timestamp).ToString("MM/dd/yy HH:mm")
                } else { 
                    "Never" 
                }
                
                # Check if device has screenshot available
                if ($device.last_boot_test_screenshot_presented -eq $true) {
                    # Build composite key matching screenshot storage format: DeviceId_Type_SessionId
                    $compositeKey = if ($device.last_recovery_session_id) {
                        "$deviceId`_$($device.type)`_$($device.last_recovery_session_id)"
                    } else {
                        "$deviceId`_$($device.type)"  # Fallback to DeviceId_Type
                    }
                    
                    if ($Script:ScreenshotFiles -and $Script:ScreenshotFiles.ContainsKey($compositeKey)) {
                        $screenshotData = $Script:ScreenshotFiles[$compositeKey]
                        
                        $borderColor = switch ($screenshotData.Status) {
                            'SUCCESS' { '#10b981' }
                            'FAILED' { '#ef4444' }
                            default { '#f59e0b' }
                        }
                        
                        $consoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/DRAAS)"
                        
                        $Script:screenshotGalleryDRaaSCompact += @"
<div class='screenshot-compact-card' style='background: #1f2937; border-radius: 8px; overflow: hidden; cursor: pointer;' onclick="openScreenshotModal('$deviceId')">
    <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; border: 3px solid $borderColor;' alt='DRaaS Screenshot' loading='lazy'>
    <div style='padding: 10px; text-align: center; font-weight: 600; font-size: 13px; color: #a78bfa; background: #111827;'>$($screenshotData.DeviceName)</div>
</div>
"@

                        $Script:screenshotGalleryDRaaS += @"
<div id='screenshot-draas-$deviceId' class='screenshot-tile'>
    <div style='background: #1f2937; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);'>
        <img src='$($screenshotData.Path)' style='width: 100%; height: auto; display: block; cursor: pointer; border: 3px solid $borderColor;' alt='DRaaS Screenshot' loading='lazy' onclick="openScreenshotModal('$deviceId')">
        <div style='padding: 14px; color: #e5e7eb; font-size: 12px; background: #111827;'>
            <div style='font-weight: 700; margin-bottom: 8px;'><a href='$consoleLink' target='_blank' style='color: #a78bfa; text-decoration: none;'>$($screenshotData.DeviceName)</a></div>
        </div>
    </div>
</div>
"@
                    }
                }
                
                $duration = if ($device.last_recovery_duration_user) { $device.last_recovery_duration_user } else { "N/A" }
                
                # Use last_recovery_restored_size_user from Excel (already in GB)
                $restoredSize = if ($device.last_recovery_restored_size_user) { "$($device.last_recovery_restored_size_user) GB" } else { "N/A" }
                
                # Type, Region, OS Type, Backup Session, GB/hr, Plan, Recovery Location, Colorbar
                $typeDisplay = "☁️ Azure DRaaS"
                $region = if ($device.region_name) { $device.region_name } else { "N/A" }
                $osTypeRaw = if ($device.backup_cloud_device_machine_os_type) { $device.backup_cloud_device_machine_os_type } else { "N/A" }
                $osType = Shorten-OSVersion $osTypeRaw
                $backupSession = if ($device.last_backup_session_timestamp) { ([datetime]$device.last_backup_session_timestamp).ToString("MM/dd/yy HH:mm") } else { "N/A" }
                $gbPerHour = "N/A"
                if ($device.last_recovery_restored_size_user -and $device.last_recovery_duration_seconds) {
                    try {
                        $sizeGB = [decimal]$device.last_recovery_restored_size_user
                        $durationHours = [decimal]$device.last_recovery_duration_seconds / 3600
                        if ($durationHours -gt 0) { $gbPerHour = [Math]::Round($sizeGB / $durationHours, 2).ToString() }
                    } catch {}
                }
                $planName = if ($device.plan_name) { $device.plan_name } else { "N/A" }
                $recoveryLocationName = "N/A"
                if ($device.agent_id -and $Script:RecoveryLocations) {
                    $rlMatch = $Script:RecoveryLocations | Where-Object { $_.recovery_agent_id -eq $device.agent_id } | Select-Object -First 1
                    if ($rlMatch) { $recoveryLocationName = $rlMatch.name }
                }
                $colorbarHtml = ""
                if ($device.colorbar_detail) {
                    $cbEntries = @($device.colorbar_detail -split ',' | Where-Object { $_ })
                    if ($cbEntries.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbEntry in $cbEntries[0..([Math]::Min($cbEntries.Count,28)-1)]) {
                            $cbParts = $cbEntry -split '\|'
                            if ($cbParts.Count -eq 2) {
                                $cbColor = switch ($cbParts[1]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                                $colorbarHtml += "<div data-timestamp='$($cbParts[0])' data-status='$($cbParts[1])' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                            }
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } elseif ($device.colorbar) {
                    $cbStatuses = @($device.colorbar -split ',' | Where-Object { $_ })
                    if ($cbStatuses.Count -gt 0) {
                        $colorbarHtml = "<div style='display:flex;gap:2px;align-items:center;min-width:112px;'>"
                        foreach ($cbIdx in 0..([Math]::Min($cbStatuses.Count,28)-1)) {
                            $cbColor = switch ($cbStatuses[$cbIdx]) { 'Completed'{'#10b981'} 'CompletedWithErrors'{'#f59e0b'} 'Failed'{'#ef4444'} 'InProgress'{'#3b82f6'} 'Interrupted'{'#f97316'} 'NotStarted'{'#d1d5db'} 'Aborted'{'#9ca3af'} default{'#e5e7eb'} }
                            $colorbarHtml += "<div title='Day -${cbIdx}: $($cbStatuses[$cbIdx])' style='width:4px;height:24px;background:$cbColor;border-radius:1px;cursor:help;'></div>"
                        }
                        $colorbarHtml += "</div>"
                    } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                } else { $colorbarHtml = "<span style='color:#9ca3af;font-size:11px;'>No data</span>" }
                
                # Build console link for DRaaS device
                $draasConsoleLink = "https://backup.management/#/continuity/view/default(panel:device-properties/$deviceId/recovery-verification/DRAAS)"
                $deviceNameLink = "<a href='$draasConsoleLink' target='_blank' style='color: #3b82f6; text-decoration: none; font-weight: 500;' onmouseover='this.style.textDecoration=``underline``' onmouseout='this.style.textDecoration=``none``'>$deviceName</a>"
                
                $draasRows += "<tr><td>$deviceNameLink</td><td>$partnerName</td><td>$typeDisplay</td><td>$region</td><td><small>$osType</small></td><td><span style='color: $statusColor; font-weight: 600;'>$status</span></td><td>$backupSession</td><td>$lastRecovery</td><td>$duration</td><td>$restoredSize</td><td><strong style='color: #a78bfa;'>$gbPerHour</strong></td><td><small>$planName</small></td><td>$recoveryLocationName</td><td>$colorbarHtml</td></tr>"
            }
        }
        
        # Build Recovery Locations table rows
        $recoveryLocationsRows = ""
        if ($RecoveryLocations -and @($RecoveryLocations).Count -gt 0) {
            foreach ($location in $RecoveryLocations) {
                $locationName = $location.name
                $agentType = $location.type
                $status = if ($location.agent_state) { $location.agent_state } else { "UNKNOWN" }
                $statusColor = switch ($status) {
                    'ONLINE' { '#10b981' }
                    'OFFLINE' { '#ef4444' }
                    default { '#f59e0b' }
                }
                $statusDisplay = $status.ToUpper()
                $totalSpace = if ($location.partition_total_gb -and $location.partition_total_gb -gt 0) { "$($location.partition_total_gb) GB" } else { "N/A" }
                $usedSpace = if ($location.partition_used_gb -and $location.partition_used_gb -gt 0) { "$($location.partition_used_gb) GB" } else { "N/A" }
                $freeSpace = if ($location.partition_free_gb -and $location.partition_free_gb -gt 0) { "$($location.partition_free_gb) GB" } else { "N/A" }
                $usedPercent = if ($location.partition_used_percent -and $location.partition_used_percent -gt 0) { "$($location.partition_used_percent)%" } else { "0%" }
                
                $recoveryLocationsRows += "<tr><td>$locationName</td><td>$agentType</td><td><span style='color: $statusColor; font-weight: 600;'>$statusDisplay</span></td><td>$totalSpace</td><td>$usedSpace</td><td>$freeSpace</td><td>$usedPercent</td></tr>"
            }
        }
        
        # Count by type
        $recoveryTestingTotal = if ($RecoveryTestingDevices) { @($RecoveryTestingDevices).Count } else { 0 }
        $standbyImageTotal = if ($SBIDevices) { @($SBIDevices).Count } else { 0 }
        $oneTimeRestoreTotal = if ($OneTimeRestoreDevices) { @($OneTimeRestoreDevices).Count } else { 0 }
        $draasTotal = if ($DRaaSDevices) { @($DRaaSDevices).Count } else { 0 }
        $recoveryLocationsTotal = if ($RecoveryLocations) { @($RecoveryLocations).Count } else { 0 }
        
        return @"
        <div class="content-section" id="continuity">
            <div class="section-header">
                <div class="section-title">$($sectionInfo.Name) <span style="background: #0097D6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px; margin-left: 8px;">$($sectionInfo.ID)</span></div>
                <div class="section-subtitle">$($sectionInfo.Description)</div>
            </div>
            
            <div class="dashboard-grid" id="continuity-overview">
                <div class="stat-card" style="border-left: 4px solid #10b981;">
                    <div class="stat-label">Recovery Testing</div>
                    <div class="stat-value" style="color: #10b981;">$recoveryTestingTotal</div>
                    <div style="font-size: 12px; color: #6b7280; margin-top: 8px;">Active recovery test plans</div>
                    <div class="card-id">DASH-009-01</div>
                </div>
                <div class="stat-card" style="border-left: 4px solid #3b82f6;">
                    <div class="stat-label">Standby Image</div>
                    <div class="stat-value" style="color: #3b82f6;">$standbyImageTotal</div>
                    <div style="font-size: 12px; color: #6b7280; margin-top: 8px;">VDR/Standby devices</div>
                    <div class="card-id">DASH-009-02</div>
                </div>
                <div class="stat-card" style="border-left: 4px solid #f59e0b;">
                    <div class="stat-label">One-Time Restore</div>
                    <div class="stat-value" style="color: #f59e0b;">$oneTimeRestoreTotal</div>
                    <div style="font-size: 12px; color: #6b7280; margin-top: 8px;">On-demand recovery</div>
                    <div class="card-id">DASH-009-03</div>
                </div>
                <div class="stat-card" style="border-left: 4px solid #8b5cf6;">
                    <div class="stat-label">DRaaS</div>
                    <div class="stat-value" style="color: #8b5cf6;">$draasTotal</div>
                    <div style="font-size: 12px; color: #6b7280; margin-top: 8px;">Disaster Recovery as a Service</div>
                    <div class="card-id">DASH-009-04</div>
                </div>
                <div class="stat-card" style="border-left: 4px solid #0097D6;">
                    <div class="stat-label">Recovery Locations</div>
                    <div class="stat-value" style="color: #0097D6;">$recoveryLocationsTotal</div>
                    <div style="font-size: 12px; color: #6b7280; margin-top: 8px;">Infrastructure agents</div>
                    <div class="card-id">DASH-009-05</div>
                </div>
            </div>
            
            <div style="margin-top: 30px;" id="continuity-recovery">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">♻️ Recovery Testing Status (<span id="recoveryTestingVisibleCount">$recoveryTestingTotal</span>)</h3>
                        <span style="background: #10b981; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-A</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <input type="text" id="recoveryTestingSearch" placeholder="🔍 Search devices..." style="padding: 6px 12px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 13px; background: #ffffff; color: #111827; width: 200px;" oninput="filterRecoveryTestingTable()">
                        <span style="color: #9ca3af; font-size: 13px; white-space: nowrap;">Showing: <span id="recoveryTestingCountDisplay" style="color: #10b981; font-weight: 600;">$recoveryTestingTotal</span> of <span id="recoveryTestingTotalCount" style="color: #9ca3af;">$recoveryTestingTotal</span></span>
                        <button onclick="toggleTable('recoveryTestingTable')" class="toggle-btn" style="padding: 8px 16px; background: #10b981; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 600; transition: background 0.2s; min-width: 110px;" onmouseover="this.style.background='#059669'" onmouseout="this.style.background='#10b981'">▼ Collapse</button>
                    </div>
                </div>
                <div style="background: #f0fdf4; border-left: 4px solid #10b981; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #065f46; font-size: 14px; line-height: 1.6;">
                        <strong>✅ Best Practice:</strong> Regular recovery testing validates your backup strategy. Schedule periodic restore tests to ensure data recoverability. 
                        <strong>$recoveryTestingTotal</strong> recovery testing plans are currently active.
                    </p>
                </div>
                
                <div class="data-table" id="recoveryTestingTable">
                    <table id="recoveryTestingTableElement" style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white;">
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(0)">Device Name</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(1)">Partner</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(2)">Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(3)">Region</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(4)">OS Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(5)">Status</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(6)">Backup Session</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(7)">Last Recovery</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(8)">Duration</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(9)">Restored Size</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(10)">GB/hr</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(11)">Plan</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortRecoveryTestingTable(12)">Recovery Location</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Colorbar</th>
                            </tr>
                        </thead>
                        <tbody>
                            $recoveryTestingRows
                        </tbody>
                    </table>
                </div>
            </div>
            
            <div style="margin-top: 30px;" id="continuity-screenshots">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">📸 Recovery Testing Boot Screenshots</h3>
                        <span style="background: #10b981; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-E</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <button onclick="toggleViewMode()" id="view-mode-toggle" class='view-mode-btn' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🔲 Detailed</button>
                        <button onclick="cycleTimeView()" id="time-view-toggle" class='tz-btn-global' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>📍 Device</button>
                    </div>
                </div>
                <div style="background: #d1fae5; border-left: 4px solid #10b981; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #065f46; font-size: 14px; line-height: 1.6;">
                        <strong>🔄 Recovery Testing:</strong> These boot screenshots verify successful recovery testing for Recovery Testing devices only. 
                        Toggle between <strong>Compact</strong> view (image + name) and <strong>Detailed</strong> view (full metadata). 
                        Click any screenshot to view full-screen with device details. Green border = Success, Red border = Failed.
                    </p>
                </div>
                <div id="gallery-compact" style="display: none; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
                    $Script:screenshotGalleryRTCompact
                </div>
                <div id="gallery-detailed" class="screenshot-gallery-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px;">
                    $Script:screenshotGalleryRT
                </div>
            </div>
            
            <div style="margin-top: 30px;" id="continuity-standby">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">💾 Standby Image Devices (<span id="standbyImageVisibleCount">$standbyImageTotal</span>)</h3>
                        <span style="background: #3b82f6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-B</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <input type="text" id="standbyImageSearch" placeholder="🔍 Search devices..." style="padding: 6px 12px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 13px; background: #ffffff; color: #111827; width: 200px;" oninput="filterStandbyImageTable()">
                        <span style="color: #9ca3af; font-size: 13px; white-space: nowrap;">Showing: <span id="standbyImageCountDisplay" style="color: #3b82f6; font-weight: 600;">$standbyImageTotal</span> of <span id="standbyImageTotalCount" style="color: #9ca3af;">$standbyImageTotal</span></span>
                        <button onclick="toggleTable('standbyImageTable')" class="toggle-btn" style="padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 600; transition: background 0.2s; min-width: 110px;" onmouseover="this.style.background='#2563eb'" onmouseout="this.style.background='#3b82f6'">▼ Collapse</button>
                    </div>
                </div>
                <div style="background: #eff6ff; border-left: 4px solid #3b82f6; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #1e3a8a; font-size: 14px; line-height: 1.6;">
                        <strong>💡 Best Practice:</strong> Standby Image (VDR) provides virtualized disaster recovery. 
                        Monitor boot status and ensure LSV paths are accessible. Click any column header to sort the table, or use the search below to filter devices. Boot screenshots sync with table order in the gallery at the end of this section.
                    </p>
                </div>
                
                <div class="data-table" id="standbyImageTable">
                    <table id="standbyImageTableElement" style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white;">
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(0)">Device Name</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(1)">Partner</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(2)">Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(3)">Region</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(4)">OS Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(5)">Status</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(6)">Backup Session</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(7)">Last Recovery</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(8)">Duration</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(9)">Restored Size</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(10)">GB/hr</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(11)">Plan</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortStandbyImageTable(12)">Recovery Location</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Colorbar</th>
                            </tr>
                        </thead>
                        <tbody>
                            $standbyImageRows
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Standby Image Screenshots Section -->
            <div style="margin-top: 30px;" id="continuity-standby-screenshots">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">📸 Standby Image Boot Screenshots</h3>
                        <span style="background: #3b82f6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-D</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <button onclick="toggleViewMode()" id="standby-view-toggle" class='view-mode-btn' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🔲 Detailed</button>
                        <button onclick="cycleTimeView()" id="standby-time-toggle" class='tz-btn-global' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🕐 Local</button>
                    </div>
                </div>
                <div style="background: #dbeafe; border-left: 4px solid #3b82f6; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #1e3a8a; font-size: 14px; line-height: 1.6;">
                        <strong>💾 Standby Image:</strong> These boot screenshots verify successful recovery testing for Standby Image devices only (VDR). 
                        Toggle between <strong>Compact</strong> view (image + name) and <strong>Detailed</strong> view (full metadata). 
                        Click any screenshot to view full-screen with device details. Green border = Success, Red border = Failed.
                    </p>
                </div>
                <div id="standby-gallery-compact" style="display: none; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
                    $Script:screenshotGallerySBICompact
                </div>
                <div id="standby-gallery-detailed" class="screenshot-gallery-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px;">
                    $Script:screenshotGallerySBI
                </div>
            </div>
            
            <!-- DRaaS Section -->
            <div style="margin-top: 30px;" id="continuity-draas-devices">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">🚀 DRaaS Devices (<span id="draasVisibleCount">$draasTotal</span>)</h3>
                        <span style="background: #8b5cf6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-C</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <input type="text" id="draasSearch" placeholder="🔍 Search devices..." style="padding: 6px 12px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 13px; background: #ffffff; color: #111827; width: 200px;" oninput="filterDRaaSTable()">
                        <span style="color: #9ca3af; font-size: 13px; white-space: nowrap;">Showing: <span id="draasCountDisplay" style="color: #8b5cf6; font-weight: 600;">$draasTotal</span> of <span id="draasTotalCount" style="color: #9ca3af;">$draasTotal</span></span>
                        <button onclick="toggleTable('draasTable')" class="toggle-btn" style="padding: 8px 16px; background: #8b5cf6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 600; transition: background 0.2s; min-width: 110px;" onmouseover="this.style.background='#7c3aed'" onmouseout="this.style.background='#8b5cf6'">▼ Collapse</button>
                    </div>
                </div>
                <div style="background: #f5f3ff; border-left: 4px solid #8b5cf6; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #5b21b6; font-size: 14px; line-height: 1.6;">
                        <strong>☁️ Disaster Recovery as a Service:</strong> DRaaS provides automated failover to cloud infrastructure for business continuity. 
                        These devices maintain continuous synchronization for rapid recovery in disaster scenarios. 
                        Click any column header to sort the table, or use the search to filter devices. Boot screenshots sync with table order in the gallery below.
                    </p>
                </div>
                
                <div class="data-table" id="draasTable">
                    <table id="draasTableElement" style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); color: white;">
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(0)">Device Name</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(1)">Partner</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(2)">Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(3)">Region</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(4)">OS Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(5)">Status</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(6)">Backup Session</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(7)">Last Recovery</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(8)">Duration</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(9)">Restored Size</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(10)">GB/hr</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(11)">Plan</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortDRaaSTable(12)">Recovery Location</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Colorbar</th>
                            </tr>
                        </thead>
                        <tbody>
                            $draasRows
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- DRaaS Screenshots Section -->
            <div style="margin-top: 30px;" id="continuity-draas-screenshots">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">📸 DRaaS Boot Screenshots</h3>
                        <span style="background: #8b5cf6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-G</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <button onclick="toggleViewMode()" id="draas-view-toggle" class='view-mode-btn' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🔲 Detailed</button>
                        <button onclick="cycleTimeView()" id="draas-time-toggle" class='tz-btn-global' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🕐 Local</button>
                    </div>
                </div>
                <div style="background: #f5f3ff; border-left: 4px solid #8b5cf6; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #5b21b6; font-size: 14px; line-height: 1.6;">
                        <strong>☁️ DRaaS Recovery:</strong> These boot screenshots verify successful recovery testing for DRaaS devices. 
                        Toggle between <strong>Compact</strong> view (image + name) and <strong>Detailed</strong> view (full metadata). 
                        Click any screenshot to view full-screen with device details. Green border = Success, Red border = Failed.
                    </p>
                </div>
                <div id="draas-gallery-compact" style="display: none; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
                    $Script:screenshotGalleryDRaaSCompact
                </div>
                <div id="draas-gallery-detailed" class="screenshot-gallery-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px;">
                    $Script:screenshotGalleryDRaaS
                </div>
            </div>
            
            <div style="margin-top: 30px;" id="continuity-draas">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 15px;">
                    <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">📦 One-Time Restore</h3>
                    <span style="background: #f59e0b; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-F</span>
                </div>
                <div style="background: #f5f3ff; border-left: 4px solid #8b5cf6; padding: 16px; border-radius: 4px;">
                    <p style="margin: 0; color: #5b21b6; font-size: 14px; line-height: 1.6;">
                        <strong>💡 Tip:</strong> One-Time Restore provides flexible recovery options for individual files, folders, or full systems. 
                        Consider implementing for flexible recovery scenarios. Currently <strong>$oneTimeRestoreTotal</strong> one-time restore configurations available.
                    </p>
                </div>
            </div>
            
            <!-- One-Time Restore Section -->
            <div style="margin-top: 30px;" id="continuity-one-time-restore">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">🔄 One-Time Restore Devices</h3>
                        <span style="background: #f59e0b; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-I</span>
                    </div>
                    <button onclick="toggleTable('oneTimeRestoreTable')" class="toggle-btn" style="padding: 8px 16px; background: #f59e0b; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 600; transition: background 0.2s; min-width: 110px;" onmouseover="this.style.background='#d97706'" onmouseout="this.style.background='#f59e0b'">▼ Collapse</button>
                </div>
                <div style="background: #fffbeb; border-left: 4px solid #f59e0b; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #92400e; font-size: 14px; line-height: 1.6;">
                        <strong>🔄 Flexible Recovery:</strong> One-Time Restore devices provide on-demand recovery capabilities to various targets (Azure, ESXi, or Self-Hosted). 
                        These devices can be restored when needed without maintaining continuous standby images. 
                        <strong>$oneTimeRestoreTotal</strong> one-time restore configurations available.
                    </p>
                </div>
                
                <div class="data-table" id="oneTimeRestoreTable">
                    <table id="oneTimeRestoreTableElement" style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white;">
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(0)">Device Name</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(1)">Partner</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(2)">Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(3)">Region</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(4)">OS Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(5)">Status</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(6)">Backup Session</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(7)">Last Recovery</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(8)">Duration</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(9)">Restored Size</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(10)">GB/hr</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(11)">Plan</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600; cursor: pointer;" onclick="sortOneTimeRestoreTable(12)">Recovery Location</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Colorbar</th>
                            </tr>
                        </thead>
                        <tbody>
                            $oneTimeRestoreRows
                        </tbody>
                    </table>
                </div>
                
                <!-- One-Time Restore Screenshots -->
$(if ($Script:screenshotGalleryOTR) {@"
                <div style="margin-top: 25px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                        <h4 style="color: #92400e; font-size: 16px; font-weight: 600; margin: 0;">📸 One-Time Restore Screenshots</h4>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <button onclick="toggleViewMode()" id="view-mode-toggle" class='view-mode-btn' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🔲 Detailed</button>
                            <button onclick="cycleTimeView()" id="time-view-toggle" class='tz-btn-global' style='padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; min-width: 110px;'>🕐 Local</button>
                        </div>
                    </div>
                    <div id="otr-gallery-compact" style="display: none; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 15px;">
                        $Script:screenshotGalleryOTRCompact
                    </div>
                    <div id="otr-gallery-detailed" class="screenshot-gallery-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px;">
                        $Script:screenshotGalleryOTR
                    </div>
                </div>
"@})
            </div>
            
            <!-- Recovery Locations Section -->
            <div style="margin-top: 30px;" id="continuity-recovery-locations">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <h3 style="color: var(--primary-dark); font-size: 18px; font-weight: 600; margin: 0;">📍 Recovery Locations</h3>
                        <span style="background: #0097D6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px;">$($sectionInfo.ID)-H</span>
                    </div>
                    <button onclick="toggleTable('recoveryLocationsTable')" class="toggle-btn" style="padding: 8px 16px; background: #0097D6; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 600; transition: background 0.2s; min-width: 110px;" onmouseover="this.style.background='#007bb5'" onmouseout="this.style.background='#0097D6'">▼ Collapse</button>
                </div>
                <div style="background: #e6f7ff; border-left: 4px solid #0097D6; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
                    <p style="margin: 0; color: #004d73; font-size: 14px; line-height: 1.6;">
                        <strong>🏗️ Infrastructure Agents:</strong> Recovery Locations are infrastructure agents that host Standby Images and facilitate recovery operations. 
                        Monitor agent status, version, and available storage to ensure recovery capability. 
                        <strong>$recoveryLocationsTotal</strong> recovery locations available.
                    </p>
                </div>
                
                <div class="data-table" id="recoveryLocationsTable">
                    <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <thead>
                            <tr style="background: linear-gradient(135deg, #0097D6 0%, #007bb5 100%); color: white;">
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Location Name</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Agent Type</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Status</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Version</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Total Space</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Used Space</th>
                                <th style="padding: 12px; text-align: left; font-weight: 600;">Free Space</th>
                            </tr>
                        </thead>
                        <tbody>
                            $recoveryLocationsRows
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
"@
    }
    catch {
        return New-DashSectionError -SectionInfo $sectionInfo -ErrorMessage $_.Exception.Message -ErrorDetails "Line: $($_.InvocationInfo.ScriptLineNumber)"
    }
}

#endregion DASHBOARD SECTION BUILDERS

#endregion ----- Functions ----

#region ========== MAIN EXECUTION ==========

## ── Required modules ─────────────────────────────────────────────────────────
Install-RequiredModules -Modules @('ImportExcel')

## ── Authenticate ─────────────────────────────────────────────────────────────
Send-APICredentialsCookie

## ── Partner lookup ───────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($PartnerName)) {
    Send-GetPartnerInfo $Script:cred0
} else {
    $Script:OriginalPartnerName = $PartnerName
    Send-GetPartnerInfo $PartnerName
}

## ── Export path ──────────────────────────────────────────────────────────────
$Script:cleanpartnername = $Script:partnername -replace '[^a-zA-Z0-9]', '_'
$reportFolder = "Continuity_$($Script:cleanpartnername)_$($Script:partnerid)"
if ($ExportPath) {
    $Script:ExportPath = Join-Path $ExportPath $reportFolder
} else {
    $Script:ExportPath = Join-Path $PSScriptRoot $reportFolder
}
New-Item -ItemType Directory -Path $Script:ExportPath -Force | Out-Null
Write-Output "  Export folder: $Script:ExportPath"

## ── Collect data ─────────────────────────────────────────────────────────────
try {
    Get-ContinuityStatistics
} catch {
    Write-Warning "  Get-ContinuityStatistics failed: $($_.Exception.Message)"
}

try {
    Get-RecoveryLocations
} catch {
    Write-Warning "  Get-RecoveryLocations failed: $($_.Exception.Message)"
}

## ── Ensure all arrays initialised ────────────────────────────────────────────
if (-not $Script:RecoveryTestingDevices) { $Script:RecoveryTestingDevices = @() }
if (-not $Script:SBIDevices)             { $Script:SBIDevices             = @() }
if (-not $Script:OneTimeRestoreDevices)  { $Script:OneTimeRestoreDevices  = @() }
if (-not $Script:DRaaSDevices)           { $Script:DRaaSDevices           = @() }
if (-not $Script:RecoveryLocations)      { $Script:RecoveryLocations      = @() }

#endregion ========== MAIN EXECUTION ==========

#region ========== HTML OUTPUT ==========

Write-Output $Script:strLineSeparator
Write-Output "  Generating Business Continuity HTML Report..."

try {
    $partnerDisplay = if ($Script:OriginalPartnerName) { $Script:OriginalPartnerName }
                      elseif ($Script:partnername)     { $Script:partnername }
                      else                             { "Unknown Partner" }
    $generatedDate  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $devicesData    = @()   ## Statistics API data not needed for continuity section

    ## ── Build section HTML ───────────────────────────────────────────────────
    $continuityHtml = Build-DashSectionContinuity `
        -RecoveryTestedCount    $Script:RecoveryTestingDevices.Count `
        -RecoveryOldTests       0 `
        -RecoveryNeverTested    0 `
        -StandbyImageCount      $Script:SBIDevices.Count `
        -RecoveryTestingDevices $Script:RecoveryTestingDevices `
        -SBIDevices             $Script:SBIDevices `
        -OneTimeRestoreDevices  $Script:OneTimeRestoreDevices `
        -DRaaSDevices           $Script:DRaaSDevices `
        -RecoveryLocations      $Script:RecoveryLocations `
        -DevicesData            $devicesData

    ## Activate section (standalone page - no sidebar toggle needed)
    $continuityHtml = $continuityHtml -replace 'class="content-section"', 'class="content-section active"'

    ## ── Assemble full HTML page ───────────────────────────────────────────────
    $htmlPage = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Business Continuity - $partnerDisplay</title>
    <style>
        :root {
            --primary-dark: #00457C;
            --primary-blue: #0097D6;
            --accent-gold:  #FFB81C;
            --bg-main:      #f5f7fa;
            --bg-white:     #ffffff;
            --border-color: #e0e4e8;
            --text-primary:   #1a1a1a;
            --text-secondary: #6b7280;
            --shadow-sm: 0 1px 3px rgba(0,0,0,.08);
            --shadow-md: 0 4px 6px rgba(0,0,0,.10);
            --transition:    all .3s ease;
            --color-success:      #10b981;
            --color-warning:      #f59e0b;
            --color-danger-light: #ea580c;
            --color-danger:       #dc2626;
            --color-info:         #3b82f6;
            --color-muted:        #6b7280;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: var(--bg-main);
            color: var(--text-primary);
            margin: 0; padding: 0;
        }
        body.dark-mode {
            --bg-main:      #0a0a0a;
            --bg-white:     #121212;
            --border-color: rgba(255,255,255,.12);
            --text-primary:   #e8e8e8;
            --text-secondary: rgba(255,255,255,.6);
            --shadow-sm: 0 2px 4px rgba(0,0,0,.5);
            --shadow-md: 0 4px 8px rgba(0,0,0,.7);
            --color-success:      #34d399;
            --color-warning:      #fbbf24;
            --color-danger-light: #fb923c;
            --color-danger:       #f87171;
            --color-info:         #60a5fa;
            --color-muted:        #9ca3af;
        }
        /* ── Page header ─────────────────────────────────────────────────── */
        .page-header {
            background: linear-gradient(135deg, var(--primary-dark) 0%, #0066A1 100%);
            color: white;
            padding: 14px 24px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: var(--shadow-md);
            position: sticky; top: 0; z-index: 100;
        }
        .page-header-left { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
        .page-title  { font-size: 17px; font-weight: 700; letter-spacing: -.3px; }
        .badge {
            padding: 4px 10px; border-radius: 12px;
            font-size: 11px; font-weight: 700; letter-spacing: .3px;
        }
        .badge-continuity { background: #10b981; color: #fff; }
        .badge-partner    { background: rgba(255,255,255,.15); color: #fff; }
        .badge-version    { background: #4b5563; color: #d1d5db; font-weight: 500; }
        .badge-date       { background: rgba(255,255,255,.08); color: rgba(255,255,255,.75); font-weight: 400; }
        .header-controls  { display: flex; gap: 10px; }
        .ctrl-btn {
            background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.2);
            color: #fff; padding: 7px 14px; border-radius: 6px;
            cursor: pointer; font-size: 12px; font-weight: 600;
            transition: var(--transition);
        }
        .ctrl-btn:hover { background: rgba(255,255,255,.22); }
        /* ── Content ─────────────────────────────────────────────────────── */
        .page-content   { padding: 20px 24px; }
        .content-section{ display: none; flex-direction: column; }
        .content-section.active { display: flex; }
        /* ── Stat cards ──────────────────────────────────────────────────── */
        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 16px; margin-bottom: 20px;
        }
        @media (max-width:1400px) { .dashboard-grid { grid-template-columns: repeat(4,1fr); } }
        @media (max-width:1100px) { .dashboard-grid { grid-template-columns: repeat(3,1fr); } }
        @media (max-width: 800px) { .dashboard-grid { grid-template-columns: repeat(2,1fr); } }
        @media (max-width: 500px) { .dashboard-grid { grid-template-columns: 1fr; } }
        .stat-card {
            background: var(--bg-white); padding: 18px 20px;
            border-radius: 8px; border: 1px solid var(--border-color);
            box-shadow: var(--shadow-sm); transition: var(--transition);
            position: relative; overflow: hidden;
        }
        .stat-card::before {
            content:''; position: absolute; top:0; left:0;
            width:4px; height:100%;
            background: linear-gradient(180deg,var(--primary-blue) 0%,var(--primary-dark) 100%);
        }
        .stat-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-md); }
        .stat-label { font-size:11px; font-weight:700; text-transform:uppercase;
            color:var(--text-secondary); letter-spacing:.5px; margin-bottom:8px; }
        .stat-value { font-size:30px; font-weight:700; color:var(--primary-dark); line-height:1; }
        body.dark-mode .stat-value { color: #60a5fa; }
        .card-id { position:absolute; bottom:5px; right:7px; font-size:9px;
            color:var(--text-secondary); opacity:.5; }
        /* ── Data tables ─────────────────────────────────────────────────── */
        .data-table { width:100%; background:var(--bg-white); border-radius:8px;
            overflow:hidden; box-shadow:var(--shadow-sm); border:1px solid var(--border-color);
            display:block; margin-bottom:0; }
        .data-table table { width:100%; border-collapse:collapse; }
        .data-table th {
            background: linear-gradient(135deg,var(--primary-dark) 0%,#0066A1 100%);
            color:#fff; padding:11px 10px; text-align:left;
            font-weight:600; font-size:11px; text-transform:uppercase;
            letter-spacing:.4px; cursor:pointer; white-space:nowrap;
        }
        .data-table th:hover { filter: brightness(1.15); }
        .data-table td {
            padding:9px 10px; border-bottom:1px solid rgba(0,0,0,.06);
            font-size:12.5px; color:var(--text-primary);
        }
        body.dark-mode .data-table td { border-bottom-color:rgba(255,255,255,.05); }
        .data-table tbody tr:hover { background: rgba(0,151,214,.05); }
        body.dark-mode .data-table tbody tr:hover { background:rgba(96,165,250,.08); }
        .data-table tbody tr:last-child td { border-bottom:none; }
        /* ── Section headers ─────────────────────────────────────────────── */
        .section-header { margin-bottom:20px; }
        .section-title  { font-size:22px; font-weight:700; color:var(--primary-dark); margin-bottom:4px; }
        body.dark-mode .section-title { color: #60a5fa; }
        .section-subtitle { font-size:13px; color:var(--text-secondary); }
        /* ── Screenshot galleries ────────────────────────────────────────── */
        .screenshot-gallery-grid {
            display:grid; grid-template-columns:repeat(auto-fill,minmax(300px,1fr));
            gap:20px;
        }
        .screenshot-tile {
            background:#1a1a2e; border-radius:10px; overflow:hidden;
            box-shadow:0 4px 16px rgba(0,0,0,.35); transition:transform .2s,box-shadow .2s;
        }
        .screenshot-tile:hover { transform:translateY(-3px); box-shadow:0 8px 28px rgba(0,0,0,.45); }
        .screenshot-compact-card {
            background:#1a1a2e; border-radius:8px; overflow:hidden;
            box-shadow:0 2px 10px rgba(0,0,0,.3); cursor:pointer; transition:transform .2s;
        }
        .screenshot-compact-card:hover { transform:scale(1.03); }
        /* ── Toggle / view-mode buttons ──────────────────────────────────── */
        .toggle-btn, .view-mode-btn, .tz-btn-global { transition:all .2s; }
        /* ── Sort indicators ─────────────────────────────────────────────── */
        th.sort-asc::after  { content:' ▲'; font-size:9px; opacity:.8; }
        th.sort-desc::after { content:' ▼'; font-size:9px; opacity:.8; }
    </style>
</head>
<body>
<div class="page-header">
    <div class="page-header-left">
        <span class="page-title">COVE DATA PROTECTION</span>
        <span class="badge badge-continuity">CONTINUITY REPORT</span>
        <span class="badge badge-partner">$partnerDisplay</span>
        <span class="badge badge-version">Script $ScriptVersion</span>
        <span class="badge badge-date">$generatedDate</span>
    </div>
    <div class="header-controls">
        <button class="ctrl-btn" onclick="toggleDarkMode()">🌙 Dark</button>
        <button class="ctrl-btn" onclick="window.print()">🖨️ Print</button>
    </div>
</div>
<div class="page-content">
$continuityHtml
</div>
<script>
    /* ── Dark mode ──────────────────────────────────────────────────────── */
    function toggleDarkMode() {
        document.body.classList.toggle('dark-mode');
        event.target.innerHTML = document.body.classList.contains('dark-mode') ? '☀️ Light' : '🌙 Dark';
    }

    /* ── Toggle table collapse ──────────────────────────────────────────── */
    function toggleTable(tableId) {
        const tbl = document.getElementById(tableId);
        const btn = event.target;
        const hidden = tbl.style.display === 'none';
        tbl.style.display = hidden ? 'block' : 'none';
        btn.textContent = hidden ? '▼ Collapse' : '▶ Expand';
    }

    /* ── Screenshot view mode (compact / detailed) ──────────────────────── */
    let _viewMode = 'detailed';
    function toggleViewMode() {
        _viewMode = _viewMode === 'compact' ? 'detailed' : 'compact';
        const isCompact = _viewMode === 'compact';
        document.querySelectorAll('.view-mode-btn').forEach(b => {
            b.innerHTML = isCompact ? '👁️ Compact' : '🔲 Detailed';
        });
        ['standby-gallery-compact','draas-gallery-compact','gallery-compact','otr-gallery-compact']
            .forEach(id => { const el = document.getElementById(id); if (el) el.style.display = isCompact ? 'grid' : 'none'; });
        ['standby-gallery-detailed','draas-gallery-detailed','gallery-detailed','otr-gallery-detailed']
            .forEach(id => { const el = document.getElementById(id); if (el) el.style.display = isCompact ? 'none' : 'grid'; });
    }

    /* ── Timezone cycle ─────────────────────────────────────────────────── */
    let _tz = 'local';
    function cycleTimeView() {
        _tz = _tz === 'device' ? 'local' : _tz === 'local' ? 'utc' : 'device';
        const label = _tz === 'device' ? '📍 Device' : _tz === 'local' ? '🕐 Local' : '🌍 UTC';
        document.querySelectorAll('.tz-btn-global').forEach(b => b.innerHTML = label);
        document.querySelectorAll('[class*="time-"]').forEach(el => el.style.display = 'none');
        document.querySelectorAll('[class*="time-' + _tz + '-"]').forEach(el => el.style.display = 'inline');
    }

    /* ── Screenshot fullscreen overlay ─────────────────────────────────── */
    function viewFullScreenshot(deviceId) {
        const tile = document.getElementById('screenshot-' + deviceId);
        if (!tile) return;
        const img = tile.querySelector('img');
        if (!img) return;
        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.95);' +
            'display:flex;align-items:center;justify-content:center;z-index:9999;cursor:zoom-out;';
        const fi = document.createElement('img');
        fi.src = img.src;
        fi.style.cssText = 'max-width:95%;max-height:95%;object-fit:contain;box-shadow:0 10px 40px rgba(0,0,0,.8);';
        overlay.appendChild(fi);
        document.body.appendChild(overlay);
        overlay.onclick = () => document.body.removeChild(overlay);
        const esc = e => { if (e.key === 'Escape') { document.body.removeChild(overlay); document.removeEventListener('keydown',esc); } };
        document.addEventListener('keydown', esc);
    }
    function openScreenshotModal(id) { viewFullScreenshot(id); }
    function closeScreenshotModal()  {}

    /* ── Generic table sort ──────────────────────────────────────────────── */
    function _sortTable(table, colIdx) {
        if (!table) return;
        const tbody = table.querySelector('tbody');
        const rows  = Array.from(tbody.querySelectorAll('tr'));
        const th    = table.querySelectorAll('thead th')[colIdx];
        const dir   = th?.getAttribute('data-sort-dir') === 'asc' ? 'desc' : 'asc';
        table.querySelectorAll('thead th').forEach(h => { h.removeAttribute('data-sort-dir'); h.classList.remove('sort-asc','sort-desc'); });
        if (th) { th.setAttribute('data-sort-dir', dir); th.classList.add(dir === 'asc' ? 'sort-asc' : 'sort-desc'); }
        rows.sort((a, b) => {
            const av = a.cells[colIdx]?.textContent.trim() || '';
            const bv = b.cells[colIdx]?.textContent.trim() || '';
            const an = parseFloat(av.replace(/[^0-9.-]/g, ''));
            const bn = parseFloat(bv.replace(/[^0-9.-]/g, ''));
            if (!isNaN(an) && !isNaN(bn)) return dir === 'asc' ? an - bn : bn - an;
            return dir === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av);
        });
        rows.forEach(r => tbody.appendChild(r));
    }

    /* ── Table-specific sort/filter wrappers ────────────────────────────── */
    function sortRecoveryTestingTable(c) { _sortTable(document.getElementById('recoveryTestingTableElement'), c); }
    function sortStandbyImageTable(c)    { _sortTable(document.getElementById('standbyImageTableElement'), c); }
    function sortOneTimeRestoreTable(c)  { _sortTable(document.getElementById('oneTimeRestoreTableElement'), c); }
    function sortDRaaSTable(c)           { _sortTable(document.getElementById('draasTableElement'), c); }
    function syncRecoveryTestingGalleryOrder() {}
    function syncStandbyImageGalleryOrder()    {}
    function syncDRaaSGalleryOrder()           {}

    function _filterTable(inputId, tableId, countId) {
        const s   = (document.getElementById(inputId)?.value || '').toLowerCase();
        const tbl = document.getElementById(tableId);
        if (!tbl) return;
        let n = 0;
        tbl.querySelectorAll('tbody tr').forEach(r => {
            const show = !s || r.textContent.toLowerCase().includes(s);
            r.style.display = show ? '' : 'none';
            if (show) n++;
        });
        const el = document.getElementById(countId);
        if (el) el.textContent = n;
    }
    function filterRecoveryTestingTable() { _filterTable('recoveryTestingSearch','recoveryTestingTable','recoveryTestingCountDisplay'); }
    function filterStandbyImageTable()    { _filterTable('standbyImageSearch','standbyImageTableElement','standbyImageCountDisplay'); }
    function filterDRaaSTable()           { _filterTable('draasSearch','draasTableElement','draasCountDisplay'); }

    /* ── Init: click-to-sort on all data tables ─────────────────────────── */
    function initColorbarTooltips() {}
    document.addEventListener('DOMContentLoaded', () => {
        document.querySelectorAll('.data-table table').forEach(tbl => {
            tbl.querySelectorAll('thead th').forEach((th, i) => {
                th.style.cursor = 'pointer';
                th.addEventListener('click', () => _sortTable(tbl, i));
            });
        });
    });
</script>
</body>
</html>
"@

    ## ── Save file ──────────────────────────────────────────────────────────
    $htmlFile = "$Script:ExportPath\$($Script:cleanpartnername)_$($Script:partnerid)_$($Script:CurrentDate)_Continuity.html"
    $htmlPage | Out-File -FilePath $htmlFile -Encoding UTF8 -ErrorAction Stop
    Write-Output "  --> Report saved: $htmlFile"

    if (Test-Path $htmlFile) {
        Start-Process $htmlFile
        Write-Output "  --> Opening in browser"
    }

} catch {
    Write-Warning "  HTML generation failed: $($_.Exception.Message)"
    Write-Warning "  Line: $($_.InvocationInfo.ScriptLineNumber)"
}

Write-Output $Script:strLineSeparator
#endregion ========== HTML OUTPUT ==========

ExitRoutine
