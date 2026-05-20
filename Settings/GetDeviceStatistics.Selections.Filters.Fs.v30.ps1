<# ----- About: ----
    # Get N-able Backup Device Selections Filters Statistics 
    # Revision v30 - 2026-05-20
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
    #
    # v29 Updates (2026-05-13 - AI-Assisted):
    # - ForEach-Object -Parallel support (PS 7+) with configurable -ThrottleLimit (default=10)
    # - Sequential fallback path for PS 5 with warning
    # - Function bodies serialized as strings + [scriptblock]::Create() to satisfy PS7 $using: restrictions
    # - Thread-safe counter (Monitor lock), log bag (ConcurrentBag), and failed-device bag (ConcurrentBag)
    # - Second-pass sequential retry for devices that exhaust all parallel retries
    # - Added retry logic (3x exponential backoff) to Get-AccountInfoById
    # - Added SSL/TLS error pattern to retry conditions in Get-DeviceAuditQuery and Get-AccountInfoById
    # - Added retry logic (3x exponential backoff) to Send-BatchUpdateCustomColumns
    # - Added URL validation guard in Get-DeviceAuditQuery (skips audit if repserver URL malformed)
    # - Fixed Get-Credential -UserName "" (removed empty UserName parameter)
    # - Removed hardcoded __cfduid cookie from Send-UpdateCustomColumn
    # - Fixed unreachable $reader.Close() in CallJSON
    # - Fixed undefined $plugin/$selected/$path variables in Filter PSCustomObject
    # - Fixed undefined $Filters variable in Selection PSCustomObject
    #
    # v28 Updates (2026-02-11 - AI-Assisted):
    # - Applied v25/v26/v27 bug fixes to v27 codebase (merged from v26 fix branch)
    #
    # v27 Updates (2025-11-12 - AI-Assisted):
    # - Batch custom column updates - all 4 columns updated in single API call (75% reduction in API calls)
    # - Reduced delay to 25ms for 60+ devices/min throughput
    # - Skip logging for "All (*)" selections to reduce log file size
    # - Batch log file writing every 50 devices with progress metrics
    # - Added comprehensive performance metrics (elapsed time, devices/min, ETA)
    #
    # v26 Updates (2025-11-05 - AI-Assisted):
    # - Added SSLTest (Selection Simplification Logic Test) to simplify chronological selection data
    # - Added custom column AA3135 output for simplified selection results
    # - Added detailed analysis log file generation
    # - Added retry logic with exponential backoff for connection failures
    # - Added timeout protection (30 seconds) for API calls
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
    # For use with the Standalone edition of N-able Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select devices
    # Get device statistics for selected partner/ devices
    # Get device session statistics for selected partner/ devices
    # Optionally export to XLS/CSV
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -AllDevices switch parameter to skip GUI device selection
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -Export switch parameter to export statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    # Use the -DeviceFilter (default="AT==1 && OT==2 && TL>0") parameter to filter devices by Active, Online and Total Last Session Time
    # Use the -Size (default=50) parameter to set the number of total sessions to query
    # Use the -Unique (default=1) parameter to set the number of unique sessions per data source to return
    # Use the -Order (default=ASC) parameter to set the order of sessions to pull
    # Use the -Query parameter to set the query to filter sessions
    
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [switch]$AllPartners = $true,                    ## Skip partner selection
        [Parameter(Mandatory=$False)] [switch]$AllDevices = $true,                     ## Skip device selection
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 5000,                         ## Change Maximum Number of devices results to return
        [Parameter(Mandatory=$False)] [switch]$Export = $true,                          ## Generate CSV / XLS Output Files
        [Parameter(Mandatory=$False)] [switch]$Launch = $true,                          ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                         ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",                    ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials,                        ## Remove Stored API Credentials at start of script
        [Parameter(Mandatory=$False)] [switch]$debugdata = $false,                          ## Enable debug data output
        [Parameter(Mandatory=$False)] [string]$DeviceFilter = "AT==1 && TL>0",  ## Filter for Active, Online and Total Last Session Time
        [Parameter(Mandatory=$False)] [int]$DeviceId = 0,                                   ## Specific Device ID for troubleshooting individual devices
        [Parameter(Mandatory=$False)] [int]$ThrottleLimit = 20                               ## Max parallel threads (PS 7+ only)

        #= "AU == 5087125 or AU == 5103094 or AU == 5103803 or AU == 5104819 or AU == 5106219 or AU == 5107375 or AU == 5107460 or AU == 5111057 or AU == 5117509 or AU == 5125031 or AU == 5125414 or AU == 5044298"  ## Filter for Devices
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Cove Backup Selections and Filters from Audit"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

    Write-Output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"


    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urljson = "https://api.backup.management/jsonapi"
    
    # Append DeviceId to DeviceFilter if specified
    if ($DeviceId -gt 0) {
        $DeviceFilter = "$DeviceFilter && AU==$DeviceId"
    }
    
    # Initialize SSLTest log
    $Script:SSLTestLog = @"
SSLTest (Selection Simplification Logic Test) Analysis Log
Generated: $CurrentDate
Script: GetDeviceStatistics.Selections.Filters.Fs.v27.ps1

"@

    Write-output "  Current Parameters:"
    Write-output "  -AllPartners = $AllPartners"
    Write-output "  -AllDevices  = $AllDevices"
    Write-output "  -DeviceCount = $DeviceCount"
    Write-output "  -Export      = $Export"
    Write-output "  -Launch      = $Launch"
    Write-output "  -ExportPath  = $ExportPath"
    Write-output "  -Delimiter   = $Delimiter"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 

        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Data Protection | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove Data Protection  API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

    $BackupCred.UserName | Out-file -append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
    
    Start-Sleep -milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    $Script:APIcredfile = join-path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-path -path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential File Cleared"
        Send-APICredentialsCookie  ## Retry Authentication
        
        }else{ 
            Write-Output $Script:strLineSeparator 
            Write-Output "  Getting Backup API Credentials" 
        
            if (Test-Path $APIcredfile) {
                Write-Output    $Script:strLineSeparator        
                "  Backup API Credential File Present"
                $APIcredentials = get-content $APIcredfile
                
                $Script:cred0 = [string]$APIcredentials[0] 
                $Script:cred1 = [string]$APIcredentials[1]
                $Script:cred2 = $APIcredentials[2] | Convertto-SecureString 
                $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))

                Write-Output    $Script:strLineSeparator 
                Write-output "  Stored Backup API Partner  = $Script:cred0"
                Write-output "  Stored Backup API User     = $Script:cred1"
                Write-output "  Stored Backup API Password = Encrypted"
                
            }else{
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

$webrequest = Invoke-WebRequest -Method POST `
    -ContentType 'application/json' `
    -Body (ConvertTo-Json $data) `
    -Uri $url `
    -SessionVariable Script:websession `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:Authenticate = $webrequest | convertfrom-json

#Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

if ($authenticate.visa) { 

    $Script:visa = $authenticate.visa
    $script:UserId = $authenticate.result.result.id
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator 
        
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
}

#endregion ----- Authentication ----

#region ----- Data Conversion ----
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
}  ## Convert epoch time to date time 

Function Save-CSVasExcel {
    param (
        [string]$CSVFile = $(Throw 'No file provided.')
    )
    
    BEGIN {
        function Resolve-FullPath ([string]$Path) {    
            if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) {
                # $Path = Join-Path (Get-Location) $Path
                $Path = "$PWD\$Path"
            }
            [IO.Path]::GetFullPath($Path)
        }

        function Release-Ref ($ref) {
            ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        
        $CSVFile = Resolve-FullPath $CSVFile
        $xl = New-Object -ComObject Excel.Application
    }

    PROCESS {
        $wb = $xl.workbooks.open($CSVFile)
        $xlOut = $CSVFile -replace '\.csv$', '.xlsx'
        
        # can comment out this part if you don't care to have the columns autosized
        $ws = $wb.Worksheets.Item(1)
        $range = $ws.UsedRange
        [void]$range.AutoFilter()
        [void]$range.EntireColumn.Autofit()

        $num = 1
        $dir = Split-Path $xlOut
        $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
        $nextname = $xlOut
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }

        $wb.SaveAs($nextname, 51)
    }

    END {
        $xl.Quit()
    
        $null = $ws, $wb, $xl | ForEach-Object {Release-Ref $_}

        # del $CSVFile
    }
}  ## Save as output XLS Routine

Function Get-DataSourceNames {
    <#
    .SYNOPSIS
        Decodes AP column string into human-readable data source names
    
    .DESCRIPTION
        The AP column contains a string where each character represents an active data source.
        This function converts those character codes into readable data source names.
    
    .PARAMETER apString
        The AP column value (e.g., "FSQ" or "FWNZ")
    
    .OUTPUTS
        Returns a comma-separated string of data source names
    
    .EXAMPLE
        Get-DataSourceNames "FSQ"
        Returns: "Files and Folders, System State, MS SQL"
    #>
    
    param(
        [Parameter(Mandatory=$false)]
        [string]$apString
    )
    
    if ([string]::IsNullOrWhiteSpace($apString)) {
        return "None"
    }
    
    $sources = @()
    
    # Map each character to its data source name based on N-able Cove API
    if ($apString -match 'F') { $sources += "Files and Folders" }
    if ($apString -match 'S') { $sources += "System State" }
    if ($apString -match 'Q') { $sources += "MS SQL" }
    if ($apString -match 'X') { $sources += "VSS Exchange" }
    if ($apString -match 'N') { $sources += "Network Shares" }
    if ($apString -match 'W') { $sources += "VMware Virtual Machines" }
    if ($apString -match 'Z') { $sources += "VSS MS SQL" }
    if ($apString -match 'P') { $sources += "VSS SharePoint" }
    if ($apString -match 'Y') { $sources += "Oracle" }
    if ($apString -match 'H') { $sources += "Hyper-V" }
    if ($apString -match 'L') { $sources += "MySQL" }
    
    if ($sources.Count -eq 0) {
        return "Unknown: $apString"
    }
    
    return $sources -join ', '
}  ## Decode AP column data source codes

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

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

    $RestrictedPartnerLevel = @("Root","Sub-root")

    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$script:Level = $Partner.result.result.Level
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
        }else{
        Write-Output $Script:strLineSeparator
        Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
        }

    if ($partner.error) {
        write-output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername

    }

}  ## get PartnerID and Partner Level    

Function CallJSON($url,$object) {

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($object)
    $web = [System.Net.WebRequest]::Create($url)
    $web.Method = "POST"
    $web.ContentLength = $bytes.Length
    $web.ContentType = "application/json"
    $stream = $web.GetRequestStream()
    $stream.Write($bytes,0,$bytes.Length)
    $stream.close()
    $reader = New-Object System.IO.Streamreader -ArgumentList $web.GetResponse().GetResponseStream()
    $result = $reader.ReadToEnd() | ConvertFrom-Json
    $reader.Close()
    return $result
}  ## Call JSON Web Request Function

Function Send-EnumeratePartners {
    # ----- Get Partners via EnumeratePartners -----
    
    # (Create the JSON object to call the EnumeratePartners function)
        $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
            Add-Member -PassThru NoteProperty visa $Script:visa |
            Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
            Add-Member -PassThru NoteProperty params @{
                                                        parentPartnerId = $PartnerId 
                                                        fetchRecursively = "true"
                                                        fields = (0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22) 
                                                        } |
            Add-Member -PassThru NoteProperty id '1')| ConvertTo-Json -Depth 5
    
    # (Call the JSON Web Request Function to get the EnumeratePartners Object)
            [array]$Script:EnumeratePartnersSession = CallJSON $urlJSON $objEnumeratePartners
    
            $Script:visa = $EnumeratePartnersSession.visa
    
            #Write-Output    $Script:strLineSeparator
            #Write-Output    "  Using Visa:" $Script:visa
            #Write-Output    $Script:strLineSeparator
    
    # (Added Delay in case command takes a bit to respond)
            Start-Sleep -Milliseconds 100
    
    # (Get Result Status of EnumerateAccountProfiles)
            $EnumeratePartnersSessionErrorCode = $EnumeratePartnersSession.error.code
            $EnumeratePartnersSessionErrorMsg = $EnumeratePartnersSession.error.message
    
    # (Check for Errors with EnumeratePartners - Check if ErrorCode has a value)
            if ($EnumeratePartnersSessionErrorCode) {
                Write-Output    $Script:strLineSeparator
                Write-Output    "  EnumeratePartnersSession Error Code:  $EnumeratePartnersSessionErrorCode"
                Write-Output    "  EnumeratePartnersSession Message:  $EnumeratePartnersSessionErrorMsg"
                Write-Output    $Script:strLineSeparator
                Write-Output    "  Exiting Script"
    # (Exit Script if there is a problem)
    
                #Break Script
            }
                Else {
    # (No error)
    
            $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
            
            $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
        
            $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                    
            $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
            
            
            if ($AllPartners) {
                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                Write-Output    $Script:strLineSeparator
                Write-Output    "  All Partners Selected"
            }else{
                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single
        
                if($null -eq $Selection) {
                    # Cancel was pressed
                    # Run cancel script
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  No Partners Selected"
                    Break
                }
                else {
                    # OK was pressed, $Selection contains what was chosen
                    # Run OK script
                    [int]$script:PartnerId = $script:Selection.Id
                    [String]$script:PartnerName = $script:Selection.Name
                }
            }

    }
    
}  ## EnumeratePartners API Call

Function Send-GetDevices {

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'EnumerateAccountStatistics'
    $data.params = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId = [int]$PartnerId
    $data.params.query.Filter = $DeviceFilter
    $data.params.query.Columns = @("AU","AR","AN","TM","MN","AL","LN","OP","OI","OS","OT","PD","AP","PF","PN","CD","TS","TL","T3","US","TB","TM","I81","AA843","AA77","AA2048","AA58","FA","F3","F5","SA","S3","S5","ZA","Z3","Z5","T0","T7","YV","VE","I78","OV","AA2646","AA2647","AA2648","AA2650","AA2906","AA3309","ES","I86","I87","I88","I89","I90","AA3444")
    $data.params.query.OrderBy = "AU ASC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = $devicecount
    $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }  

    $Script:DeviceResponse = Invoke-RestMethod @params 

    $Script:DeviceDetail = @()
    
    ForEach ( $DeviceResult in $DeviceResponse.result.result ) {
        $apCode = $DeviceResult.Settings.AP -join ''
        
        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{
            AccountID = [Int]$DeviceResult.AccountId;
            PartnerID        = [string]$DeviceResult.PartnerId;
            DeviceName       = $DeviceResult.Settings.AN -join '' ;
            ComputerName     = $DeviceResult.Settings.MN -join '' ;
            DeviceAlias      = $DeviceResult.Settings.AL -join '' ;
            PartnerName      = $DeviceResult.Settings.AR -join '' ;
            Reference        = $DeviceResult.Settings.PF -join '' ;
            Creation         = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
            TimeStamp        = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;
            LastSuccess      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '') ;
            Last28Days       = (($DeviceResult.Settings.TB -join '')[-1..-28] -join '') -replace("8",[char]0x26a0) -replace("7",[char]0x23f9) -replace("6",[char]0x23f9) -replace("5",[char]0x2611) -replace("2",[char]0x274e) -replace("1",[char]0x2BC8) -replace("0",[char]0x274c) ;
            Last28           = (($DeviceResult.Settings.TB -join '')[-1..-28] -join '') -replace("8","!") -replace("7","!") -replace("6","?") -replace("5","+") -replace("2","-") -replace("1",">") -replace("0","X") ;
            SelectedGB       = (($DeviceResult.Settings.T3 -join '') / 1GB) ;
            UsedGB           = (($DeviceResult.Settings.US -join '') / 1GB) ;
            DataSources      = $apCode ;
            DataSourceNames  = Get-DataSourceNames $apCode ;
            TotalM365        = $DeviceResult.Settings.TM -join '' ;   
            Account          = $DeviceResult.Settings.AU -join '' ;
            Location         = $DeviceResult.Settings.LN -join '' ;
            Notes            = $DeviceResult.Settings.AA843 -join '' ;
            GUIPassword      = $DeviceResult.Settings.AA2048 -join '' ;                                                                    
            TempInfo         = $DeviceResult.Settings.AA77 -join '' ;
            Product          = $DeviceResult.Settings.PN -join '' ;
            ProductID        = $DeviceResult.Settings.PD -join '' ;
            Profile          = $DeviceResult.Settings.OP -join '' ;
            ProfileID        = $DeviceResult.Settings.OI -join '' ;
            OS               = $DeviceResult.Settings.OS -join '' ;                                                                
            OSType           = $DeviceResult.Settings.OT -join '' ;
            Physicality      = $DeviceResult.Settings.I81 -join '';
            FFDuration       = New-TimeSpan -Seconds ($DeviceResult.Settings.FA -join '') | ForEach-Object { "{0:D2}:{1:D2}:{2:D2}" -f ($_.Days * 24 + $_.Hours), $_.Minutes, $_.Seconds } ;
            FFSelected       = (($DeviceResult.Settings.F3 -join '') / 1GB) ;
            FFSent           = (($DeviceResult.settings.F5 -join '') / 1GB) ;
            SSDuration       = New-TimeSpan -Seconds ($DeviceResult.Settings.SA -join '') | ForEach-Object { "{0:D2}:{1:D2}:{2:D2}" -f ($_.Days * 24 + $_.Hours), $_.Minutes, $_.Seconds } ;
            SSSelected       = (($DeviceResult.Settings.S3 -join '') / 1GB) ;
            SSSent           = (($DeviceResult.Settings.S5 -join '') / 1GB) ;
            SQLDuration      = New-TimeSpan -Seconds ($DeviceResult.Settings.ZA -join '') | ForEach-Object { "{0:D2}:{1:D2}:{2:D2}" -f ($_.Days * 24 + $_.Hours), $_.Minutes, $_.Seconds } ;
            SQLSelected      = (($DeviceResult.Settings.Z3 -join '') / 1GB) ;
            SQLSent          = (($DeviceResult.Settings.Z5 -join '') / 1GB) ;
            LastStatus       = $DeviceResult.Settings.T0     -join '' ;
            TotalErrors      = $DeviceResult.Settings.T7     -join '' ;
            LSVStatus        = $DeviceResult.Settings.YV     -join '' ;
            LSVEnabled       = $DeviceResult.Settings.VE     -join '' ;
            ActiveDataSources = $DeviceResult.Settings.I78   -join '' ;
            ProfileVersion   = $DeviceResult.Settings.OV     -join '' ;
            Vols             = $DeviceResult.Settings.AA2646 -join '' ;
            Includes         = $DeviceResult.Settings.AA2647 -join '' ;
            Excludes         = $DeviceResult.Settings.AA2648 -join '' ;
            USB_Vols         = $DeviceResult.Settings.AA2650 -join '' ;
            DetectedDataSources = $DeviceResult.Settings.AA2906 -join '' ;
            MissedVols       = $DeviceResult.Settings.AA3309 -join '' ;
            EncryptionStatus = $DeviceResult.Settings.ES     -join '' ;
            FIPSStatus       = $DeviceResult.Settings.I86    -join '' ;
            FIPSDetails      = $DeviceResult.Settings.I87    -join '' ;
            mTLSStatus       = $DeviceResult.Settings.I88    -join '' ;
            mTLSDetails      = $DeviceResult.Settings.I89    -join '' ;
            DRaaS            = $DeviceResult.Settings.I90    -join '' ;
            SeedHistory      = $DeviceResult.Settings.AA3444 -join ''
        }
    } ## End ForEach Loop
}  ## EnumerateAccountStatistics API Call to Get Devices

Function Format-CsvSafe {
    ## Prevents Excel formula injection by prefixing strings that start with =, +, -, or @
    Param ([string]$Value)
    if ($Value -match '^[=+\-@]') { return "`t$Value" }
    return $Value
}  ## Format-CsvSafe

Function Send-GetExportReport {
    ## Post-run EnumerateAccountStatistics export using the Henry Schein custom view columns
    $url    = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data   = @{}
    $data.jsonrpc = '2.0'
    $data.id      = '2'
    $data.visa    = $Script:visa
    $data.method  = 'EnumerateAccountStatistics'
    $data.params  = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId         = [int]$PartnerId
    $data.params.query.Filter            = $DeviceFilter
    $data.params.query.Columns           = @("AN","MN","AU","AR","OT","T3","US","TB","T0","T7","CD","TL","TS","YV","PD","PN","OS","LN","VE","I81","AA2646","I78","OI","OV","OP","AA3308","AA3135","AA3136","AA2649","AA2647","AA2648","AA2650","AA2906","AA3309","ES","I86","I87","I88","I89","I90","AA3444")
    $data.params.query.OrderBy           = "AU ASC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount      = $DeviceCount
    $data.params.query.Totals            = @("COUNT(AT==1)","SUM(T3)","SUM(US)")

    $jsondata = (ConvertTo-Json $data -depth 6)
    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }

    $Script:ExportResponse = Invoke-RestMethod @params
    $Script:ExportDetail   = @()

    ForEach ($DeviceResult in $Script:ExportResponse.result.result) {
        $Script:ExportDetail += New-Object -TypeName PSObject -Property @{
            AccountID            = [Int]$DeviceResult.AccountId
            PartnerID            = [string]$DeviceResult.PartnerId
            DeviceName           = $DeviceResult.Settings.AN     -join ''
            ComputerName         = $DeviceResult.Settings.MN     -join ''
            Account              = $DeviceResult.Settings.AU     -join ''
            PartnerName          = $DeviceResult.Settings.AR     -join ''
            OSType               = $DeviceResult.Settings.OT     -join ''
            SelectedGB           = [math]::Round(([double]($DeviceResult.Settings.T3 -join '')) / 1GB, 3)
            UsedGB               = [math]::Round(([double]($DeviceResult.Settings.US -join '')) / 1GB, 3)
            BackupHistory        = $DeviceResult.Settings.TB     -join ''
            LastStatus           = $DeviceResult.Settings.T0     -join ''
            TotalErrors          = $DeviceResult.Settings.T7     -join ''
            CreationDate         = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '')
            LastSuccess          = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '')
            LastStatusTime       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '')
            LSVStatus            = $DeviceResult.Settings.YV     -join ''
            ProductID            = $DeviceResult.Settings.PD     -join ''
            ProductName          = $DeviceResult.Settings.PN     -join ''
            OS                   = $DeviceResult.Settings.OS     -join ''
            Location             = $DeviceResult.Settings.LN     -join ''
            LSVEnabled           = $DeviceResult.Settings.VE     -join ''
            Physicality          = $DeviceResult.Settings.I81    -join ''
            Vols                 = Format-CsvSafe ($DeviceResult.Settings.AA2646 -join '')
            ActiveDataSources    = $DeviceResult.Settings.I78    -join ''
            ProfileID            = $DeviceResult.Settings.OI     -join ''
            ProfileVersion       = $DeviceResult.Settings.OV     -join ''
            ProfileName          = $DeviceResult.Settings.OP     -join ''
            OriginalSelections   = Format-CsvSafe ($DeviceResult.Settings.AA3308 -join '')
            SimplifiedSelections = Format-CsvSafe ($DeviceResult.Settings.AA3135 -join '')
            HumanReadable        = Format-CsvSafe ($DeviceResult.Settings.AA3136 -join '')
            Filters              = Format-CsvSafe ($DeviceResult.Settings.AA2649 -join '')
            Includes             = Format-CsvSafe ($DeviceResult.Settings.AA2647 -join '')
            Excludes             = Format-CsvSafe ($DeviceResult.Settings.AA2648 -join '')
            USB_Vols             = Format-CsvSafe ($DeviceResult.Settings.AA2650 -join '')
            DetectedDataSources  = Format-CsvSafe ($DeviceResult.Settings.AA2906 -join '')
            MissedVols           = Format-CsvSafe ($DeviceResult.Settings.AA3309 -join '')
            EncryptionStatus     = $DeviceResult.Settings.ES     -join ''
            FIPSStatus           = $DeviceResult.Settings.I86    -join ''
            FIPSDetails          = $DeviceResult.Settings.I87    -join ''
            mTLSStatus           = $DeviceResult.Settings.I88    -join ''
            mTLSDetails          = $DeviceResult.Settings.I89    -join ''
            DRaaS                = $DeviceResult.Settings.I90    -join ''
            SeedHistory          = $DeviceResult.Settings.AA3444 -join ''
        }
    } ## End ForEach Loop
}  ## Send-GetExportReport  ## RETIRED - export now built from per-device rows during processing

Function Get-AccountInfoById {
    Param (
    [int]$accountid                        ## Account
    )

    $url = "https://api.backup.management/jsonapi"

    $dataAccountInfo = @{
        method  = "GetAccountInfoById"
        params  = @{
            accountId = [int]$accountid
        }
        jsonrpc = "2.0"
        visa    = $visa
        id      = "jsonrpc"
    } | ConvertTo-Json -Depth 5

    # Retry logic for transient connection failures
    $maxRetries  = 3
    $retryCount  = 0
    $retryDelay  = 2
    $Script:repsvr = $null
    $Script:repsvr = @{ repurl = $null; AccountId = $null; Token = $null; Name = $null }

    while ($retryCount -lt $maxRetries) {
        try {
            $Script:responseAccountInfo = Invoke-RestMethod -Uri $url -Method POST -ContentType 'application/json' -Body $dataAccountInfo -TimeoutSec 30 -ErrorAction Stop

            $hostPart = ($script:responseAccountInfo.result.homeNodeInfo.CommonInfo.Host -split ':')[0]
            $Script:repsvr.repurl    = ("https://" + $hostPart + "/repserv_json").ToLower()
            $Script:repsvr.AccountId = $Script:responseAccountInfo.result.result.Id
            $Script:repsvr.Token     = $Script:responseAccountInfo.result.result.Token
            $Script:repsvr.Name      = $Script:responseAccountInfo.result.result.Name
            return  ## Success
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Warning "GetAccountInfoById failed for $accountid (attempt $retryCount of $maxRetries): $($_.Exception.Message). Retrying in $retryDelay s..."
                Start-Sleep -Seconds $retryDelay
                $retryDelay = $retryDelay * 2
            } else {
                Write-Error "GetAccountInfoById failed for $accountid after $maxRetries attempts: $($_.Exception.Message)"
                ## repsvr remains null - caller must check
            }
        }
    }
}  ## Get Account Info by ID

Function Send-UpdateCustomColumn($DeviceId,$ColumnId,$Message) {

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json; charset=utf-8")
    $headers.Add("Authorization", "Bearer $script:visa")
    
    # Sanitize message to remove problematic characters
    if ($Message) {
        # First, normalize the string to ensure proper encoding
        try {
            $Message = [Text.Encoding]::UTF8.GetString([Text.Encoding]::UTF8.GetBytes($Message))
        } catch {
            # If UTF-8 encoding fails, strip to ASCII
            $Message = [System.Text.RegularExpressions.Regex]::Replace($Message, '[^\x20-\x7E]', '?')
        }
        
        # Remove null characters
        $Message = $Message -replace '\x00', ''
        # Replace replacement character and invalid Unicode
        $Message = $Message -replace '[\uFFFD�]', '?'
        # Remove control characters (except space, tab, CR, LF)
        $Message = $Message -replace '[\x00-\x08\x0B-\x0C\x0E-\x1F]', ''
        # Replace any remaining non-printable characters
        $Message = $Message -replace '[^\x20-\x7E\x0A\x0D\x09]', '?'
    }
    
    # Build payload as object and convert to JSON
    # Need to create the nested array structure that the API expects: [[columnId, "message"]]
    $valuesArray = ,@([int]$ColumnId, [string]$Message)
    
    $payload = @{
        jsonrpc = "2.0"
        id = "jsonrpc"
        method = "UpdateAccountCustomColumnValues"
        params = @{
            accountId = [int]$DeviceId
            values = $valuesArray
        }
    }
    
    # Convert to JSON - ConvertTo-Json properly escapes backslashes as \\
    # The JSON parser will decode \\ back to \ when the API receives it
    $body = $payload | ConvertTo-Json -Depth 5 -Compress
    
    try {
        $script:updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body -ErrorAction Stop
        if ($script:updateCC.error.message) {
            Write-Warning "Custom column update error for device $DeviceId`: $($script:updateCC.error.message)"
        }
    }
    catch {
        Write-Warning "Failed to update custom column for device $DeviceId`: $($_.Exception.Message)"
    }

}

Function Send-BatchUpdateCustomColumns {
    <#
    .SYNOPSIS
        Updates multiple custom columns in a single API call
    
    .DESCRIPTION
        Batch updates multiple custom columns for a device to reduce API calls and improve performance.
        Uses the UpdateAccountCustomColumnValues API with multiple column/value pairs.
    
    .PARAMETER DeviceId
        The Account ID of the device
    
    .PARAMETER Updates
        Array of arrays, each containing [columnId, message]
        Example: @(@(2649, "filter text"), @(3308, "selection text"))
    
    .OUTPUTS
        Returns $true if successful, $false if failed
    #>
    
    param(
        [Parameter(Mandatory=$true)]
        [int]$DeviceId,
        
        [Parameter(Mandatory=$true)]
        [array]$Updates
    )

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json; charset=utf-8")
    $headers.Add("Authorization", "Bearer $script:visa")
    
    # Sanitize all messages
    $sanitizedUpdates = @()
    foreach ($update in $Updates) {
        $columnId = [int]$update[0]
        $message = [string]$update[1]
        
        if ($message) {
            # Normalize the string to ensure proper encoding
            try {
                $message = [Text.Encoding]::UTF8.GetString([Text.Encoding]::UTF8.GetBytes($message))
            } catch {
                $message = [System.Text.RegularExpressions.Regex]::Replace($message, '[^\x20-\x7E]', '?')
            }
            
            # Remove problematic characters
            $message = $message -replace '\x00', ''
            $message = $message -replace '[\uFFFD�]', '?'
            $message = $message -replace '[\x00-\x08\x0B-\x0C\x0E-\x1F]', ''
            $message = $message -replace '[^\x20-\x7E\x0A\x0D\x09]', '?'
        }
        
        $sanitizedUpdates += ,@($columnId, $message)
    }
    
    # Build payload with multiple column updates
    $payload = @{
        jsonrpc = "2.0"
        id = "jsonrpc"
        method = "UpdateAccountCustomColumnValues"
        params = @{
            accountId = [int]$DeviceId
            values = $sanitizedUpdates
        }
    }
    
    $body = $payload | ConvertTo-Json -Depth 5 -Compress

    $maxRetries = 3
    $retryCount = 0
    $retryDelay = 2
    while ($retryCount -lt $maxRetries) {
        try {
            $script:updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body -ErrorAction Stop -TimeoutSec 30
            if ($script:updateCC.error.message) {
                Write-Warning "Batch custom column update error for device $DeviceId`: $($script:updateCC.error.message)"
                return $false
            }
            return $true
        }
        catch {
            $retryCount++
            if ($_.Exception.Message -like "*SSL connection*" -or
                $_.Exception.Message -like "*timeout*" -or
                $_.Exception.Message -like "*did not properly respond*" -or
                $_.Exception.Message -like "*connected host has failed to respond*" -or
                $_.Exception.Message -like "*underlying connection was closed*") {
                if ($retryCount -lt $maxRetries) {
                    Write-Warning "Batch update transient failure for $DeviceId (attempt $retryCount of $maxRetries): $($_.Exception.Message). Retrying in $retryDelay s..."
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay = $retryDelay * 2
                } else {
                    Write-Warning "Failed to batch update custom columns for device $DeviceId after $maxRetries attempts: $($_.Exception.Message)"
                    return $false
                }
            } else {
                Write-Warning "Failed to batch update custom columns for device $DeviceId`: $($_.Exception.Message)"
                return $false
            }
        }
    }
    return $false
}

Function Get-DeviceAuditQuery {
    param(
        [int]$AccountId,
        [string]$Query = 
@"
0 != 1 and (
    Operation == 'SetMultipleBackupSelections' or 
    Operation == 'ClearBackupSelections' or 
    Operation == 'SaveFilterOptions' or 
    Operation == 'execute command: set backup filter' or 
    Operation == 'execute command: clear selection' or 
    Operation == 'execute command: set selection' or 
    Operation == 'SwitchToSeedingMode' or 
    Operation == 'SwitchToPostSeedingMode' or 
    Operation == 'execute command: switch back from seeding mode'
)
"@,

        [string]$OrderBy = "Time DESC",
        [int]$Offset = 0,
        [int]$Size = 500
    )

    Get-AccountInfoById $accountid

    # Guard: skip audit call if repserver URL is missing or malformed (GetAccountInfoById failed)
    if (-not $Script:repsvr.repurl -or $Script:repsvr.repurl -notmatch '^https://[^/]+/repserv_json$') {
        Write-Warning "Skipping audit query for $accountid - repserver URL invalid or unavailable: '$($Script:repsvr.repurl)'"
        return
    }

    $headers = @{
        "Accept" = "application/json"
        "Content-Type" = "application/json"
    }
    
    $body = @{
        id = 1
        jsonrpc = "2.0"
        method = "QueryAudit"
        params = @{
            accountId = $Script:repsvr.AccountId 
            orderBy = $OrderBy
            query = $Query
            range = @{
                Offset = $Offset
                Size = $Size
            }
            account = $Script:repsvr.Name 
            token = $Script:repsvr.Token
        }
        visa = $Script:Visa
    } | ConvertTo-Json -Depth 10
    
    # Retry logic for connection failures
    $maxRetries = 3
    $retryCount = 0
    $retryDelay = 1  # seconds
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            $Script:DeviceAuditQueryResponse = Invoke-RestMethod -Uri $Script:repsvr.repurl -Method POST -Headers $headers -ContentType "application/json" -Body $body -TimeoutSec 8
            $success = $true
            
            if ($debugdata) { return $Script:DeviceAuditQueryResponse }
        }
        catch {
            $retryCount++
            
            if ($_.Exception.Message -like "*underlying connection was closed*" -or 
                $_.Exception.Message -like "*connection that was expected to be kept alive*" -or
                $_.Exception.Message -like "*timeout*" -or
                $_.Exception.Message -like "*did not properly respond*" -or
                $_.Exception.Message -like "*connected host has failed to respond*" -or
                $_.Exception.Message -like "*SSL connection*") {
                
                if ($retryCount -lt $maxRetries) {
                    Write-Warning "Connection failed for device $($Script:repsvr.Name)| $($Script:repsvr.AccountId) @ $($Script:repsvr.repurl). Retry $retryCount of $maxRetries in $retryDelay seconds..."
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay = $retryDelay * 2  # Exponential backoff
                }
                else {
                    Write-Error "Failed to get audit query for device $($Script:repsvr.Name)| $($Script:repsvr.AccountId) @ $($Script:repsvr.repurl) after $maxRetries retries: $($_.Exception.Message)"
                    # Don't throw - continue to next device
                    $Script:DeviceAuditQueryResponse = $null
                    return
                }
            }
            else {
                # Non-connection error, fail immediately
                Write-Error "Failed to get audit query for device $($Script:repsvr.Name)| $($Script:repsvr.AccountId) @ $($Script:repsvr.repurl): $($_.Exception.Message)"
                return
            }
        }
    }
} ## Get Device Audit Query

Function Invoke-SSLTest {
    <#
    .SYNOPSIS
        Selection Simplification Logic Test - Simplifies chronological selection/exclusion data
    
    .DESCRIPTION
        Parses chronological selection data (format: "date |+ path | - path | ...") and simplifies
        by removing redundant toggles, keeping only the final state of each path.
    
    .PARAMETER SelectionString
        The raw selection string from the custom column (e.g., "25-08-15 13:58 |+ * | - F: | + F: | - F:")
    
    .PARAMETER DeviceName
        Name of the device being processed (for logging)
    
    .PARAMETER AccountID
        Account ID of the device (for logging)
    
    .OUTPUTS
        Returns a hashtable with:
        - SimplifiedString: The simplified selection string
        - HumanReadable: Plain English description
        - AnalysisLog: Detailed step-by-step analysis
    #>
    
    param(
        [Parameter(Mandatory=$true)]
        [string]$SelectionString,
        
        [Parameter(Mandatory=$false)]
        [string]$DeviceName = "Unknown",
        
        [Parameter(Mandatory=$false)]
        [string]$AccountID = "Unknown"
    )
    
    # Initialize results
    $result = @{
        Original = $SelectionString
        SimplifiedString = ""
        HumanReadable = ""
        AnalysisLog = @()
    }
    
    # Check if string is empty or null
    if ([string]::IsNullOrWhiteSpace($SelectionString)) {
        $result.SimplifiedString = ""
        $result.HumanReadable = "No selections configured"
        $result.AnalysisLog += "Empty or null selection string"
        return $result
    }
    
    # Clean up invalid/corrupted Unicode characters
    try {
        # Remove null characters and invalid Unicode
        $SelectionString = $SelectionString -replace '\x00', ''
        # Replace replacement character and other invalid Unicode (U+FFFD, etc.)
        $SelectionString = $SelectionString -replace '[\uFFFD�]', '?'
        # Remove or replace non-ASCII characters that might cause JSON issues
        # Keep common path characters but replace problematic Unicode
        $SelectionString = $SelectionString -replace '[\u0080-\uFFFF&&[^\u00A0-\u00FF]]', '?'
        # Normalize line endings
        $SelectionString = $SelectionString -replace '\r\n', ' ' -replace '\n', ' ' -replace '\r', ' '
    }
    catch {
        $result.SimplifiedString = $SelectionString
        $result.HumanReadable = "Error processing special characters"
        $result.AnalysisLog += "Character encoding error: $($_.Exception.Message)"
        return $result
    }
    
    # Split by pipe and extract date and operations
    $parts = $SelectionString -split '\|' | ForEach-Object { $_.Trim() }
    
    if ($parts.Count -lt 2) {
        # No operations found, return as-is
        $result.SimplifiedString = $SelectionString
        $result.HumanReadable = "Invalid format - no operations found"
        $result.AnalysisLog += "Cannot parse - insufficient parts"
        return $result
    }
    
    # First part is the date
    $dateStamp = $parts[0]
    $result.AnalysisLog += "Processing chronologically:"
    $result.AnalysisLog += "Date: $dateStamp"
    
    # Track final state of each path
    $pathStates = [ordered]@{}
    $operationCount = 0
    
    # Process each operation
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $operation = $parts[$i].Trim()
        
        if ($operation -match '^([+-])\s*(.+)$') {
            $operationCount++
            $action = $matches[1]
            $path = $matches[2].Trim()
            
            # Track the operation
            $actionText = if ($action -eq '+') { "Select" } else { "Exclude" }
            $result.AnalysisLog += "$operationCount. $actionText`: $path"
            
            # Store final state (will overwrite previous states of same path)
            $pathStates[$path] = $action
        }
    }
    
    # Detect toggles
    $result.AnalysisLog += ""
    $result.AnalysisLog += "Toggles detected:"
    $toggleDetected = $false
    
    # Count operations per path
    $pathCounts = @{}
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $operation = $parts[$i].Trim()
        if ($operation -match '^[+-]\s*(.+)$') {
            $path = $matches[1].Trim()
            if (-not $pathCounts.ContainsKey($path)) {
                $pathCounts[$path] = 0
            }
            $pathCounts[$path]++
        }
    }
    
    foreach ($path in $pathCounts.Keys) {
        if ($pathCounts[$path] -gt 1) {
            $toggleDetected = $true
            $finalState = if ($pathStates[$path] -eq '+') { "included" } else { "excluded" }
            $result.AnalysisLog += "- $path`: toggled $($pathCounts[$path]) times, final state = $finalState"
        }
    }
    
    if (-not $toggleDetected) {
        $result.AnalysisLog += "No toggles detected - each path appears only once"
    }
    
    # Build simplified string
    $simplifiedParts = @($dateStamp)
    foreach ($path in $pathStates.Keys) {
        $simplifiedParts += "$($pathStates[$path]) $path"
    }
    $result.SimplifiedString = $simplifiedParts -join ' |'
    
    # Build human readable description
    $selections = @()
    $exclusions = @()
    $hasWildcard = $false
    
    foreach ($path in $pathStates.Keys) {
        if ($pathStates[$path] -eq '+') {
            $selections += $path
            if ($path -eq '*') {
                $hasWildcard = $true
            }
        } else {
            $exclusions += $path
        }
    }
    
    # Sort selections and exclusions: volumes (C:, D:\, E:) first, then paths with content after backslash
    # Check if there's content after the backslash by seeing if it matches \\. (backslash followed by any character)
    $selections = $selections | Sort-Object @{Expression={$_ -match '\\.'}}, @{Expression={$_}}
    $exclusions = $exclusions | Sort-Object @{Expression={$_ -match '\\.'}}, @{Expression={$_}}
    
    if ($hasWildcard) {
        if ($exclusions.Count -gt 0) {
            $result.HumanReadable = "All except: " + ($exclusions -join ', ')
        } else {
            $result.HumanReadable = "All (*)"
        }
    } elseif ($selections.Count -gt 0) {
        $result.HumanReadable = "Backup only: " + ($selections -join ', ')
        if ($exclusions.Count -gt 0) {
            $result.HumanReadable += " (excluding: " + ($exclusions -join ', ') + ")"
        }
    } elseif ($exclusions.Count -gt 0) {
        $result.HumanReadable = "Exclude: " + ($exclusions -join ', ')
    } else {
        $result.HumanReadable = "No selections configured"
    }
    
    $result.AnalysisLog += ""
    $result.AnalysisLog += "Simplified Result:"
    $result.AnalysisLog += $result.SimplifiedString
    $result.AnalysisLog += ""
    $result.AnalysisLog += "Human Readable:"
    $result.AnalysisLog += $result.HumanReadable
    
    return $result
} ## SSLTest Function

Function Parse-DeviceAuditResponse {
    param(
        $Response = $script:DeviceAuditQueryResponse,
        [Parameter(Mandatory)] [ValidateSet("Selection", "Filter")] $Type
    )

    $script:parsedResults = @()
    
    if ($type -eq 'Selection') {
        foreach ($item in $Response.result.result | Where-Object {
            $_.Operation -eq "ClearBackupSelections" -or
            $_.Operation -eq "SetMultipleBackupSelections" -or

            $_.Operation -eq "execute command: clear selection" -or
            $_.Operation -eq "execute command: set selection"
        }) {

            $details = $item.Details

            $plugin = if ($details -match 'plugin=([^=]*?)(?=\w+=|$)') { 
                $matches[1].Trim() 
            } elseif ($item.Operation -eq 'execute command: clear selection') {
                if ($details -match 'datasources = ((?:Fs|SystemState|MsSql|Exchange|NetworkShares|VMWare|VssMsSql|VssSharePoint|Oracle|Sims|VssHyperV|MySql|LinuxSystemState)(?:, *(?:Fs|SystemState|MsSql|Exchange|NetworkShares|VMWare|VssMsSql|VssSharePoint|Oracle|Sims|VssHyperV|MySql|LinuxSystemState))*)') {
                    $matches[1]
                }
            } elseif( $item.Operation -eq 'execute command: set selection') {
                if ($details -match 'datasource = ((?:Fs|SystemState|MsSql|Exchange|NetworkShares|VMWare|VssMsSql|VssSharePoint|Oracle|Sims|VssHyperV|MySql|LinuxSystemState)(?:, *(?:Fs|SystemState|MsSql|Exchange|NetworkShares|VMWare|VssMsSql|VssSharePoint|Oracle|Sims|VssHyperV|MySql|LinuxSystemState))*)') {
                    $datasources = $matches[1] -split ', *'
                    $validSources = @('Fs','SystemState','MsSql','Exchange','NetworkShares','VMWare','VssMsSql','VssSharePoint','Oracle','Sims','VssHyperV','MySql','LinuxSystemState')
                    if ($datasources.Count -eq 1 -and $validSources -contains $datasources[0]) {
                        $datasources[0]
                    } else {
                        'Invalid'
                    }
                } else {
                    'Invalid'
                }
            } else { 
                $details
            }
                        
            $path = if ($item.Operation -eq 'execute command: set selection') { 
                # Replace 'Include' with '+', 'exclude' with '-'
                $cleanDetails = $details -replace 'Include', '| +' -replace 'exclude', '| -'
                # Remove 'datasource =' and everything up to the first + or - (or 'Include'/'exclude')
                if ($cleanDetails -match 'datasource\s*=\s*.*?([+-])') {
                    $cleanDetails = $cleanDetails.Substring($cleanDetails.IndexOf($matches[1]))
                }
                $cleanDetails.replace(' =', '')
            } else { 
                ([regex]::Matches($details, 'path=([^=]*?)(?=\w+=|$)') | ForEach-Object { 
                    $val = $_.Groups[1].Value
                    if ($val -eq " ") { '* ' } else { $val }
                }) -join '| '
            }        

            $selected = ([regex]::Matches($details, 'selected=([^=]*?)(?=\w+=|$)') | ForEach-Object { $_.Groups[1].Value }) -replace ("Inclusive","+") -replace ("Exclusive","-") -join '| '

            $includes = @()
            $selectedItems = $selected -split '\|'
            $pathItems = $path -split '\|'

            for ($i = 0; $i -lt $selectedItems.Count; $i++) {
                $sel = $selectedItems[$i].Trim()
                $pth = if ($i -lt $pathItems.Count) { $pathItems[$i].Trim() } else { "" }
                if ($sel) {
                    $includes += "$sel $pth"
                }
            }

            $includes = $includes -join ' | '

            if (($item.Operation -eq 'execute command: set selection') -and ($plugin -ne 'Invalid')) {
                $Includes = $Path

            } elseif ($item.Operation -eq 'execute command: clear selection') {
                $Includes = 'Cleared'
                

            } else {
                $Includes = $Includes
            }

            if ($plugin -like '*Fs*') { $plugin = $plugin -replace 'Fs', 'Files and folders' }
            if ($plugin -like '*SystemState*') { $plugin = $plugin -replace 'SystemState', 'System State (VSS)' }
            if ($plugin -like '*MsSql*') { $plugin = $plugin -replace 'MsSql', 'MS SQL' }
            if ($plugin -like '*Exchange*') { $plugin = $plugin -replace 'Exchange', 'MS Exchange' }
            if ($plugin -like '*NetworkShares*') { $plugin = $plugin -replace 'NetworkShares', 'Network Shares' }
            if ($plugin -like '*VMWare*') { $plugin = $plugin -replace 'VMWare', 'VMware' }
            if ($plugin -like '*VssMsSql*') { $plugin = $plugin -replace 'VssMsSql', 'VSS MS SQL' }
            if ($plugin -like '*VssHyperV*') { $plugin = $plugin -replace 'VssHyperV', 'VSS Hyper-V' }


         
            $script:parsedResults += [PSCustomObject]@{
                AccountId = $script:repsvr.AccountId
                AccountName = $script:repsvr.Name
                ActionId = $item.ActionId
                Timestamp = [DateTimeOffset]::FromUnixTimeSeconds($item.Timestamp).ToString("yyyy-MM-dd HH:mm:ss")
                User = $item.User
                Operation = $item.Operation
                Plugin = $plugin
                Selected = $selected
                Path = $path
                Selections = $Includes
                Filters = $null

                URL = $script:repsvr.repurl
                token = $script:repsvr.token
            }
        }
    }
    
    if ($type -eq 'Filter') {
        foreach ($item in $Response.result.result | Where-Object {

            $_.Operation -eq "SaveFilterOptions" -or
            $_.Operation -eq "execute command: set backup filter"

        }) {

            $details = $item.Details
            $Filters = if ($item.Operation -eq "SaveFilterOptions") { 
                if ($details -match '"([^"]+)"') {
                    (([regex]::Matches($details, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }) -join '|')
                } else {
                    "Cleared"
                }
      
            } elseif ($item.Operation -eq 'execute command: set backup filter') { 
                ($details.replace("add = ", "").replace("+ = ", "").replace("del = ", "- ").replace("- = ", "- ").replace("clean = ", "Cleared")).Trim()
            } else { 
                $null 
            }
            
            $script:parsedResults += [PSCustomObject]@{
                AccountId = $script:repsvr.AccountId
                AccountName = $script:repsvr.Name
                ActionId = $item.ActionId
                Timestamp = [DateTimeOffset]::FromUnixTimeSeconds($item.Timestamp).ToString("yyyy-MM-dd HH:mm:ss")
                User = $item.User
                Operation = $item.Operation
                Plugin = $null
                Selected = $null
                Path = $null
                Filters = $Filters

                URL = $script:repsvr.repurl
                token = $script:repsvr.token
            }
        }
    }


    return $script:parsedResults | Sort-Object ActionId
# Example usage:
# $auditResponse = Invoke-DeviceAuditQuery
# $parsedData = Parse-DeviceAuditResponse -Response $auditResponse

} ## Parse Device Audit Response based on Type

Function Get-FiltersSinceLastSaveCleared {
 

    # Get all parsed filter responses
    $parsedFilters = Parse-DeviceAuditResponse -Type 'Filter'

    # Find the latest "Cleared" filter timestamp
    $lastCleared = $parsedFilters | Where-Object { $_.Filters -eq 'Cleared' } | Sort-Object Timestamp -Descending | Select-Object -First 1
    $lastClearedTimestamp = if ($lastCleared) { $lastCleared.Timestamp } else { $null }

    # Find the latest 'SaveFilterOptions' operation timestamp
    $lastSaveFilter = $parsedFilters | Where-Object { $_.Operation -eq 'SaveFilterOptions' } | Sort-Object Timestamp -Descending | Select-Object -First 1
    $lastSaveFilterTimestamp = if ($lastSaveFilter) { $lastSaveFilter.Timestamp } else { $null }

    # Determine the later timestamp
    if ($lastClearedTimestamp -and $lastSaveFilterTimestamp) {
        $lastClearedDate = [DateTime]($lastClearedTimestamp )
        $lastSaveFilterDate = [DateTime]($lastSaveFilterTimestamp )
        $latestTimestamp = if ($lastClearedDate -gt $lastSaveFilterDate) { $lastClearedTimestamp } else { $lastSaveFilterTimestamp }
    } elseif ($lastClearedTimestamp) {
        $latestTimestamp = $lastClearedTimestamp
    } elseif ($lastSaveFilterTimestamp) {
        $latestTimestamp = $lastSaveFilterTimestamp
    } else {
        $latestTimestamp = $null
    }

    # If no "Cleared" or "SaveFilterOptions", show all; otherwise, show last and newer
    $filterResults = if ($latestTimestamp) {
        $parsedFilters | Where-Object {
            ($_.Filters -eq 'Cleared' -and $_.Timestamp -eq $latestTimestamp) -or
            ($_.Operation -eq 'SaveFilterOptions' -and $_.Timestamp -eq $latestTimestamp) -or
            ([DateTime]$_.Timestamp -gt [DateTime]$latestTimestamp)
        }
    } else {
        $parsedFilters
    }

    if ($debugdata) {
        $filterResults | Out-GridView -Title "Filters since last 'Cleared' or 'SaveFilterOptions' action (Last: $latestTimestamp)"
    }

    $allFilters = @()
    foreach ($filter in $filterResults.Filters) {
        if ($filter -and $filter -ne 'Cleared') {
            if ($filter.StartsWith('-')) {
                $items = $filter -split '\|'
                foreach ($item in $items) {
                    if ($item.Trim()) {
                        $allFilters += "-$($item.Trim().TrimStart('- '))"
                    }
                }
            } else {
                $items = $filter -split '\|'
                foreach ($item in $items) {
                    if ($item.Trim()) {
                        $allFilters += $item.Trim()
                    }
                }
            }
        }
    }
    $uniqueFilters = ($allFilters | Where-Object { $_ } | Sort-Object -Unique) -join '|'
    $lastTimestamp = if ($filterResults) { $filterResults[-1].Timestamp } else { $null }
    $formattedTimestamp = if ($lastTimestamp) { ([datetime]$lastTimestamp).ToString("yy-MM-dd HH:mm |") } else { $null }

    $script:resultString = $formattedTimestamp+$uniqueFilters

    # Return result instead of posting directly (will be batched later)
    return $resultString

  
} ## Get Filters Since Last Cleared or Save
  
Function Get-SelectionsSinceLastSaveCleared {
    param(
        [Parameter(Mandatory)] $Datasource
    )

    # Get all parsed filter responses
    $ParsedSelections = Parse-DeviceAuditResponse -Type 'Selection' | Where-Object { $_.Plugin -like "*$Datasource*" }

    # Find the latest "Cleared" selection timestamp
    $lastCleared = $ParsedSelections | Where-Object { ($_.Operation -eq 'ClearBackupSelections') -or ($_.Operation -eq 'execute command: clear selection') } | Sort-Object Timestamp -Descending | Select-Object -First 1
    $lastClearedTimestamp = if ($lastCleared) { $lastCleared.Timestamp } else { $null }

  

    # If no "Cleared Selections"show all; otherwise, show last and newer
    $SelectionResults = if ($lastClearedTimestamp) {
        $ParsedSelections | Where-Object {
            ($_.Selections -eq 'Cleared' -and $_.Timestamp -eq $lastClearedTimestamp) -or
            ([DateTime]$_.Timestamp -gt [DateTime]$lastClearedTimestamp)
        }
    } else {
        $ParsedSelections
    }

    if ($debugdata) { $SelectionResults | Out-GridView -Title "Selections for $datasource since last 'Cleared' action (Last: $lastClearedTimestamp)" }

    $SelectionString = @()


    if ($SelectionResults.Selections) {
        $selectionString = $SelectionResults.Selections.Trim() -join " | "

        if ($selectionString -match '\+ \*') {
            $lastIndex = $selectionString.LastIndexOf('+ *')
            if ($lastIndex -ge 0) {
                $selectionString = $selectionString.Substring($lastIndex)
            }
        }

    } else {
        $selectionString = "Managed by profile"
    }

    
    $lastTimestamp = if ($SelectionResults) { $SelectionResults[-1].Timestamp } else { $null }
    $formattedTimestamp = if ($lastTimestamp) { ([datetime]$lastTimestamp).ToString("yy-MM-dd HH:mm |") } else { $null }

    $script:resultString = $formattedTimestamp+$selectionString

    # Return result instead of posting directly (will be batched later)
    return $resultString

  
} ## Get Selections Since Last Cleared or Save

Function Get-SeedHistory {
    # Returns compact pipe-delimited seed event history from the already-fetched audit response.
    # Format: "26-03-06(L)Seed→H:\ |26-03-07(L)PostSeed |26-03-15(R)SwitchBack"
    # Returns empty string if no seed events found.

    $seedEvents = $Script:DeviceAuditQueryResponse.result.result | Where-Object {
        $_.Operation -eq 'SwitchToSeedingMode' -or
        $_.Operation -eq 'SwitchToPostSeedingMode' -or
        $_.Operation -eq 'execute command: switch back from seeding mode'
    } | Sort-Object { [long]$_.Timestamp }

    if (-not $seedEvents) { return "" }

    $parts = foreach ($event in $seedEvents) {
        $date = [DateTimeOffset]::FromUnixTimeSeconds([long]$event.Timestamp).ToString("yy-MM-dd")
        $userType = switch ($event.User) {
            'local user'  { '(L)' }
            'remote user' { '(R)' }
            default       { '(?)' }
        }
        $label = switch ($event.Operation) {
            'SwitchToSeedingMode' {
                $path = if ($event.Details -match 'seedingPath=([^\s,]+)') { $matches[1] } else { '' }
                if ($path) { "Seed$([char]0x2192)$path" } else { 'Seed' }
            }
            'SwitchToPostSeedingMode'                        { 'PostSeed'   }
            'execute command: switch back from seeding mode' { 'SwitchBack' }
        }
        "$date$userType$label"
    }

    return $parts -join ' |'

} ## Get-SeedHistory

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

#region ----- Main Script ----

$switch = $PSCmdlet.ParameterSetName

Send-APICredentialsCookie

Write-Output $Script:strLineSeparator
Write-Output "" 

Send-GetPartnerInfo $Script:cred0

if ($AllPartners) {}else{Send-EnumeratePartners}

Send-GetDevices $partnerId

if ($AllDevices) {
    $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,TotalM365,ComputerName,Creation,TimeStamp,LastSuccess,Last28Days,ProductId,Product,ProfileId,Profile,DataSources,DataSourceNames,SelectedGB,UsedGB,Location,OS,OSType,Physicality,companycc,contractcc,FFDuration,FFSelected,FFSent,SSDuration,SSSelected,SSSent,SQLDuration,SQLSelected,SQLSent
}else{
    $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,TotalM365,ComputerName,Creation,TimeStamp,LastSuccess,Last28Days,ProductId,Product,ProfileId,Profile,DataSources,DataSourceNames,SelectedGB,UsedGB,Location,OS,OSType,Physicality,companycc,contractcc,FFDuration,FFSelected,FFSent,SSDuration,SSSelected,SSSent,SQLDuration,SQLSelected,SQLSent  | Out-GridView -title "Current Partner | $partnername" -OutputMode Multiple}

if($null -eq $SelectedDevices) {
    # Cancel was pressed
    # Run cancel script
    Write-Output    $Script:strLineSeparator
    Write-Output    "  No Devices Selected"
    Break
}else{
    # OK was pressed, $Selection contains what was chosen
    # Run OK script

    If ($Script:Export) {
        # Initialize export file paths (progress CSV written during run; overwritten by Send-GetExportReport at end)
        $Script:exportCsvPath  = "$ExportPath\$($CurrentDate)_Export_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $Script:exportXlsxPath = $Script:exportCsvPath.Replace(".csv",".xlsx")
        $Script:exportColumnOrder = @(
            'AccountID','PartnerID','DeviceName','ComputerName','Account','PartnerName',
            'OSType','OS','Physicality','ProductID','ProductName','ProfileID','ProfileName','ProfileVersion',
            'CreationDate','LastStatusTime','LastSuccess','LastStatus','TotalErrors',
            'BackupHistory','SelectedGB','UsedGB',
            'ActiveDataSources','LSVEnabled','LSVStatus',
            'OriginalSelections','SimplifiedSelections','HumanReadable','Filters',
            'Includes','Excludes','Vols','USB_Vols','MissedVols','DetectedDataSources',
            'EncryptionStatus','FIPSStatus','FIPSDetails','mTLSStatus','mTLSDetails',
            'Location','DRaaS','SeedHistory','ColumnUpdateSuccess'
        )
        $Script:progressExportRows = [System.Collections.Generic.List[object]]::new()
    }

    # Initialize log file path and counters
    $Script:logFileName   = "$ExportPath\$($CurrentDate)_SSLTest_Analysis_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).log"
    $Script:batchSize     = 50
    $Script:totalDevices  = $Script:SelectedDevices.Count
    $Script:startTime     = Get-Date
    $Script:deviceCounter = 0

    # Write log header to disk before processing begins
    $Script:SSLTestLog | Out-File -FilePath $Script:logFileName -Encoding UTF8

    if ($PSVersionTable.PSVersion.Major -ge 7) {

        Write-Output "  PowerShell $($PSVersionTable.PSVersion) detected - running in parallel (ThrottleLimit=$ThrottleLimit)"
        Write-Output $Script:strLineSeparator

        # Capture function bodies as strings (scriptblocks cannot be passed via $using: in ForEach-Object -Parallel)
        $fn_ConvertUnixTime      = ${function:Convert-UnixTimeToDateTime}.ToString()
        $fn_GetAccountInfoById   = ${function:Get-AccountInfoById}.ToString()
        $fn_GetDeviceAuditQuery  = ${function:Get-DeviceAuditQuery}.ToString()
        $fn_ParseAuditResponse   = ${function:Parse-DeviceAuditResponse}.ToString()
        $fn_GetFiltersSince      = ${function:Get-FiltersSinceLastSaveCleared}.ToString()
        $fn_GetSelectionsSince   = ${function:Get-SelectionsSinceLastSaveCleared}.ToString()
        $fn_BatchUpdateColumns   = ${function:Send-BatchUpdateCustomColumns}.ToString()
        $fn_InvokeSSLTest        = ${function:Invoke-SSLTest}.ToString()
        $fn_FormatCsvSafe        = ${function:Format-CsvSafe}.ToString()
        $fn_GetSeedHistory       = ${function:Get-SeedHistory}.ToString()

        # Thread-safe shared objects
        $syncHash  = [hashtable]::Synchronized(@{ Counter = 0; Lock = [object]::new() })
        $logBag    = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
        $failedBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $resultBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $visaToken = $Script:visa
        $batchSz   = $Script:batchSize
        $totalDev  = $Script:totalDevices
        $startTm   = $Script:startTime

        $Script:SelectedDevices | ForEach-Object -Parallel {
            $device = $_

            # Re-define all needed functions in this runspace from string bodies
            ${function:Convert-UnixTimeToDateTime}         = [scriptblock]::Create($using:fn_ConvertUnixTime)
            ${function:Get-AccountInfoById}                = [scriptblock]::Create($using:fn_GetAccountInfoById)
            ${function:Get-DeviceAuditQuery}               = [scriptblock]::Create($using:fn_GetDeviceAuditQuery)
            ${function:Parse-DeviceAuditResponse}          = [scriptblock]::Create($using:fn_ParseAuditResponse)
            ${function:Get-FiltersSinceLastSaveCleared}    = [scriptblock]::Create($using:fn_GetFiltersSince)
            ${function:Get-SelectionsSinceLastSaveCleared} = [scriptblock]::Create($using:fn_GetSelectionsSince)
            ${function:Send-BatchUpdateCustomColumns}      = [scriptblock]::Create($using:fn_BatchUpdateColumns)
            ${function:Invoke-SSLTest}                     = [scriptblock]::Create($using:fn_InvokeSSLTest)
            ${function:Format-CsvSafe}                     = [scriptblock]::Create($using:fn_FormatCsvSafe)
            ${function:Get-SeedHistory}                    = [scriptblock]::Create($using:fn_GetSeedHistory)

            # Per-thread variable setup
            $Script:visa = $using:visaToken
            $debugdata   = $false   # Out-GridView not supported in parallel threads
            $sh          = $using:syncHash
            $lb          = $using:logBag
            $fb          = $using:failedBag
            $rb          = $using:resultBag

            # Get audit data for this device
            Get-DeviceAuditQuery $device.AccountID

            # If audit query failed entirely, queue for second pass and skip
            if (-not $Script:DeviceAuditQueryResponse) {
                $fb.Add($device)
                Write-Warning "  Queued for second pass: $($device.DeviceName) | $($device.AccountID)"
                [System.Threading.Monitor]::Enter($sh.Lock)
                try { $sh.Counter++ } finally { [System.Threading.Monitor]::Exit($sh.Lock) }
                return
            }
            Write-Output "  Audit | $($device.DeviceName) | $($device.AccountID) | $($Script:repsvr.Name) from $($Script:repsvr.repurl)"
            Write-Output "  ---------"

            $filterString      = Get-FiltersSinceLastSaveCleared
            $originalSelection = Get-SelectionsSinceLastSaveCleared -Datasource "Files and folders"
            $seedHistoryString = Get-SeedHistory

            $simplifiedMessage    = ""
            $humanReadableMessage = ""
            $selectionString = if ($originalSelection) { [string]$originalSelection } else { "" }
        
            if ($selectionString -and $selectionString -notlike "*Managed by profile" -and $selectionString.Trim() -ne "") {
                try {
                    $sslResult   = Invoke-SSLTest -SelectionString $selectionString -DeviceName $device.DeviceName -AccountID $device.AccountID
                    $skipLogging = $sslResult.HumanReadable -eq "All (*)"

                    $hasProfile         = -not [string]::IsNullOrWhiteSpace($device.Profile)
                    $hasFilesAndFolders = $device.DataSources -match 'F'

                    $simplifiedMessage = $sslResult.SimplifiedString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''

                    if ($hasProfile -and -not $hasFilesAndFolders) {
                        $humanReadableMessage = "Managed by profile, FS source not active"
                    } else {
                        $humanReadableMessage = $sslResult.HumanReadable
                    }

                    if (-not $skipLogging) {
                        $logEntry  = "`n" + ("=" * 80)
                        $logEntry += "`nDevice: $($device.DeviceName) | AccountID: $($device.AccountID)"
                        $logEntry += "`n" + ("=" * 80)
                        $logEntry += "`nDevice Profile: $($device.Profile)"
                        $logEntry += "`nData Sources Active: $($device.DataSources) ($($device.DataSourceNames))"
                        $logEntry += "`nFiles and Folders Active: $hasFilesAndFolders"
                        $logEntry += "`n--- Custom Column AA3135 (Simplified) ---"
                        $logEntry += "`nWill Post: $simplifiedMessage"
                        $logEntry += "`n--- Custom Column AA3136 (Human Readable) ---"
                        if ($hasProfile -and -not $hasFilesAndFolders) {
                            $logEntry += "`nLogic: Device has profile but FS source NOT active"
                        } elseif ($hasProfile) {
                            $logEntry += "`nLogic: Device has profile AND FS source IS active"
                        } else {
                            $logEntry += "`nLogic: No profile (device-level config)"
                        }
                        $logEntry += "`nWill Post: $humanReadableMessage"
                        $logEntry += "`n--- SSLTest Analysis ---"
                        $logEntry += "`nOriginal:`n$($sslResult.Original)"
                        foreach ($line in $sslResult.AnalysisLog) { $logEntry += "`n$line" }
                        $lb.Add($logEntry)
                }

                Write-Output "  SSLTest: $($sslResult.HumanReadable)"
            }
            catch {
                Write-Warning "SSLTest failed for $($device.DeviceName): $($_.Exception.Message)"
                $lb.Add("`nERROR: $($device.DeviceName) | $($device.AccountID): $($_.Exception.Message)")
            }
            } elseif ($selectionString -like "*Managed by profile") {
                Write-Output "  Selections managed by Profile"
                $messageWithoutDate   = $selectionString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''
                $simplifiedMessage    = $messageWithoutDate
                $humanReadableMessage = $messageWithoutDate
            } else {
                Write-Output "  No selection data found"
            }

            # Batch update custom columns - include AA3444 only if seed history found
            $colUpdates = @(
                @(2649, $filterString),
                @(3308, $originalSelection),
                @(3135, $simplifiedMessage),
                @(3136, $humanReadableMessage)
            )
            if ($seedHistoryString) { $colUpdates += ,@(3444, $seedHistoryString) }
            $batchSuccess = Send-BatchUpdateCustomColumns -DeviceId $device.AccountID -Updates $colUpdates
            $colCount = $colUpdates.Count
            if ($batchSuccess) {
                Write-Output "  Updated $colCount columns | $($device.DeviceName)"
            } else {
                Write-Warning "  Batch update failed | $($device.DeviceName)"
            }

            # Add row to progress export bag
            $rb.Add([PSCustomObject]@{
                AccountID            = $device.AccountID
                PartnerID            = $device.PartnerID
                DeviceName           = $device.DeviceName
                ComputerName         = $device.ComputerName
                Account              = $device.Account
                PartnerName          = $device.PartnerName
                OSType               = $device.OSType
                OS                   = $device.OS
                Physicality          = $device.Physicality
                ProductID            = $device.ProductID
                ProductName          = $device.Product
                ProfileID            = $device.ProfileID
                ProfileName          = $device.Profile
                ProfileVersion       = $device.ProfileVersion
                CreationDate         = $device.Creation
                LastStatusTime       = $device.TimeStamp
                LastSuccess          = $device.LastSuccess
                LastStatus           = $device.LastStatus
                TotalErrors          = $device.TotalErrors
                BackupHistory        = $device.Last28
                SelectedGB           = $device.SelectedGB
                UsedGB               = $device.UsedGB
                ActiveDataSources    = $device.ActiveDataSources
                LSVEnabled           = $device.LSVEnabled
                LSVStatus            = $device.LSVStatus
                OriginalSelections   = Format-CsvSafe "$originalSelection"
                SimplifiedSelections = Format-CsvSafe "$simplifiedMessage"
                HumanReadable        = Format-CsvSafe "$humanReadableMessage"
                Filters              = Format-CsvSafe "$filterString"
                Includes             = Format-CsvSafe ($device.Includes)
                Excludes             = Format-CsvSafe ($device.Excludes)
                Vols                 = Format-CsvSafe ($device.Vols)
                USB_Vols             = $device.USB_Vols
                MissedVols           = Format-CsvSafe ($device.MissedVols)
                DetectedDataSources  = Format-CsvSafe ($device.DetectedDataSources)
                EncryptionStatus     = $device.EncryptionStatus
                FIPSStatus           = $device.FIPSStatus
                FIPSDetails          = $device.FIPSDetails
                mTLSStatus           = $device.mTLSStatus
                mTLSDetails          = $device.mTLSDetails
                Location             = $device.Location
                DRaaS                = $device.DRaaS
                SeedHistory          = if ($seedHistoryString) { $seedHistoryString } else { $device.SeedHistory }
                ColumnUpdateSuccess  = $batchSuccess
            })

            # Thread-safe counter increment
            [System.Threading.Monitor]::Enter($sh.Lock)
            try { $sh.Counter++ } finally { [System.Threading.Monitor]::Exit($sh.Lock) }
            $count = $sh.Counter

            # Checkpoint metrics every $batchSz devices
            if ($count % $using:batchSz -eq 0) {
                $elapsed   = (Get-Date) - $using:startTm
                $remaining = $using:totalDev - $count
                $rate      = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($count / $elapsed.TotalMinutes, 2) } else { 0 }
                $eta       = if ($rate -gt 0) { [TimeSpan]::FromMinutes($remaining / $rate) } else { [TimeSpan]::Zero }
                Write-Output ""
                Write-Output "  ---------"
                Write-Output "  Checkpoint - $count of $($using:totalDev) | Rate: $rate dev/min | ETA: $($eta.ToString('hh\:mm\:ss'))"
                Write-Output "  ---------"
                Write-Output ""
            }

        } -ThrottleLimit $ThrottleLimit

        # Write collected log entries to file
        if ($logBag.Count -gt 0) {
            $logBag | Out-File -FilePath $Script:logFileName -Encoding UTF8 -Append
        }
        $Script:deviceCounter = $syncHash.Counter

        # Flush parallel progress rows to export CSV
        if ($Script:Export -and $resultBag.Count -gt 0) {
            $resultBag.ToArray() | Select-Object $Script:exportColumnOrder |
                Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
            Write-Output "  Progress export: $($resultBag.Count) devices → $Script:exportCsvPath"
        }

        # Second-pass: retry failed devices sequentially (node may have recovered)
        if ($failedBag.Count -gt 0) {
            Write-Output ""
            Write-Output $Script:strLineSeparator
            Write-Output "  Second pass: retrying $($failedBag.Count) failed devices sequentially..."
            Write-Output $Script:strLineSeparator
            foreach ($device in $failedBag) {
                Get-VisaTime
                Get-DeviceAuditQuery $device.AccountID
                Write-Output "  ---------"
                Write-Output "  Retry | $($device.DeviceName) | $($device.AccountID) | from $($Script:repsvr.repurl)"
                Write-Output "  ---------"
                if (-not $Script:DeviceAuditQueryResponse) {
                    Write-Warning "  Second pass also failed for $($device.DeviceName) | $($device.AccountID) - skipping"
                    continue
                }
                $filterString      = Get-FiltersSinceLastSaveCleared
                $originalSelection = Get-SelectionsSinceLastSaveCleared -Datasource "Files and folders"
                $seedHistoryString = Get-SeedHistory
                $simplifiedMessage    = ""
                $humanReadableMessage = ""
                $selectionString = if ($originalSelection) { [string]$originalSelection } else { "" }
                if ($selectionString -and $selectionString -notlike "*Managed by profile" -and $selectionString.Trim() -ne "") {
                    try {
                        $sslResult = Invoke-SSLTest -SelectionString $selectionString -DeviceName $device.DeviceName -AccountID $device.AccountID
                        $simplifiedMessage    = $sslResult.SimplifiedString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''
                        $hasProfile = -not [string]::IsNullOrWhiteSpace($device.Profile)
                        $hasFilesAndFolders = $device.DataSources -match 'F'
                        $humanReadableMessage = if ($hasProfile -and -not $hasFilesAndFolders) { "Managed by profile, FS source not active" } else { $sslResult.HumanReadable }
                        Write-Output "  SSLTest: $($sslResult.HumanReadable)"
                    } catch { Write-Warning "SSLTest failed for $($device.DeviceName): $($_.Exception.Message)" }
                } elseif ($selectionString -like "*Managed by profile") {
                    $msg = $selectionString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''
                    $simplifiedMessage = $msg; $humanReadableMessage = $msg
                }
                $colUpdates = @(
                    @(2649, $filterString), @(3308, $originalSelection),
                    @(3135, $simplifiedMessage), @(3136, $humanReadableMessage)
                )
                if ($seedHistoryString) { $colUpdates += ,@(3444, $seedHistoryString) }
                $batchSuccess = Send-BatchUpdateCustomColumns -DeviceId $device.AccountID -Updates $colUpdates
                $colCount = $colUpdates.Count
                if ($batchSuccess) { Write-Output "  Second pass: updated $colCount columns | $($device.DeviceName)" }
                else { Write-Warning "  Second pass: batch update failed | $($device.DeviceName)" }
                if ($Script:Export) {
                    $Script:progressExportRows.Add([PSCustomObject]@{
                        AccountID = $device.AccountID; PartnerID = $device.PartnerID
                        DeviceName = $device.DeviceName; ComputerName = $device.ComputerName
                        Account = $device.Account; PartnerName = $device.PartnerName
                        OSType = $device.OSType; OS = $device.OS; Physicality = $device.Physicality
                        ProductID = $device.ProductID; ProductName = $device.Product
                        ProfileID = $device.ProfileID; ProfileName = $device.Profile; ProfileVersion = $device.ProfileVersion
                        CreationDate = $device.Creation; LastStatusTime = $device.TimeStamp
                        LastSuccess = $device.LastSuccess; LastStatus = $device.LastStatus; TotalErrors = $device.TotalErrors
                        BackupHistory = $device.Last28; SelectedGB = $device.SelectedGB; UsedGB = $device.UsedGB
                        ActiveDataSources = $device.ActiveDataSources; LSVEnabled = $device.LSVEnabled; LSVStatus = $device.LSVStatus
                        OriginalSelections = Format-CsvSafe "$originalSelection"
                        SimplifiedSelections = Format-CsvSafe "$simplifiedMessage"
                        HumanReadable = Format-CsvSafe "$humanReadableMessage"
                        Filters = Format-CsvSafe "$filterString"
                        Includes = Format-CsvSafe ($device.Includes); Excludes = Format-CsvSafe ($device.Excludes)
                        Vols = Format-CsvSafe ($device.Vols); USB_Vols = $device.USB_Vols
                        MissedVols = Format-CsvSafe ($device.MissedVols)
                        DetectedDataSources = Format-CsvSafe ($device.DetectedDataSources)
                        EncryptionStatus = $device.EncryptionStatus; FIPSStatus = $device.FIPSStatus
                        FIPSDetails = $device.FIPSDetails; mTLSStatus = $device.mTLSStatus; mTLSDetails = $device.mTLSDetails
                        Location = $device.Location; DRaaS = $device.DRaaS; SeedHistory = if ($seedHistoryString) { $seedHistoryString } else { $device.SeedHistory }
                        ColumnUpdateSuccess = $batchSuccess
                    })
                }
            }
            # Flush second-pass rows to progress CSV
            if ($Script:Export -and $Script:progressExportRows.Count -gt 0) {
                $Script:progressExportRows | Select-Object $Script:exportColumnOrder |
                    Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -Append
                $Script:progressExportRows.Clear()
            }
            Write-Output $Script:strLineSeparator
        }

    } else {

        Write-Warning "  PowerShell 7+ required for parallel processing. Running sequentially on PS $($PSVersionTable.PSVersion)..."

        foreach ($device in $Script:SelectedDevices) {

            Get-VisaTime
            Get-DeviceAuditQuery $device.AccountID
            Start-Sleep -Milliseconds 25

            Write-Output $Script:strLineSeparator
            Write-Output "  Audit | $($device.DeviceName) | $($device.AccountID) | $($script:repsvr.Name) from $($script:repsvr.repurl)"
            Write-Output $Script:strLineSeparator

            $filterString      = Get-FiltersSinceLastSaveCleared
            $originalSelection = Get-SelectionsSinceLastSaveCleared -Datasource "Files and folders"
            $seedHistoryString = Get-SeedHistory

            $simplifiedMessage    = ""
            $humanReadableMessage = ""
            $selectionString = if ($originalSelection) { [string]$originalSelection } else { "" }

            if ($selectionString -and $selectionString -notlike "*Managed by profile" -and $selectionString.Trim() -ne "") {
                try {
                    $sslResult   = Invoke-SSLTest -SelectionString $selectionString -DeviceName $device.DeviceName -AccountID $device.AccountID
                    $skipLogging = $sslResult.HumanReadable -eq "All (*)"
                    $hasProfile         = -not [string]::IsNullOrWhiteSpace($device.Profile)
                    $hasFilesAndFolders = $device.DataSources -match 'F'
                    $simplifiedMessage = $sslResult.SimplifiedString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''
                    if ($hasProfile -and -not $hasFilesAndFolders) {
                        $humanReadableMessage = "Managed by profile, FS source not active"
                    } else {
                        $humanReadableMessage = $sslResult.HumanReadable
                    }
                    if (-not $skipLogging) {
                        $Script:SSLTestLog += "`n" + "=" * 80
                        $Script:SSLTestLog += "`nDevice: $($device.DeviceName) | AccountID: $($device.AccountID)"
                        $Script:SSLTestLog += "`n" + "=" * 80
                        $Script:SSLTestLog += "`nDevice Profile: $($device.Profile)"
                        $Script:SSLTestLog += "`nData Sources: $($device.DataSources) | FS Active: $hasFilesAndFolders"
                        $Script:SSLTestLog += "`n--- AA3135: $simplifiedMessage"
                        $Script:SSLTestLog += "`n--- AA3136: $humanReadableMessage"
                        $Script:SSLTestLog += "`n--- SSLTest ---`nOriginal: $($sslResult.Original)"
                        foreach ($line in $sslResult.AnalysisLog) { $Script:SSLTestLog += "`n$line" }
                        $Script:SSLTestLog += "`n"
                    }
                    Write-Output "  SSLTest Applied: $($sslResult.HumanReadable)"
                }
                catch {
                    Write-Warning "SSLTest failed for device $($device.DeviceName): $($_.Exception.Message)"
                    $Script:SSLTestLog += "`nERROR: $($device.DeviceName) | $($device.AccountID): $($_.Exception.Message)`n"
                }
            } elseif ($selectionString -like "*Managed by profile") {
                Write-Output "  Selections configured by Profile"
                $messageWithoutDate   = $selectionString -replace '^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\|\s*', ''
                $simplifiedMessage    = $messageWithoutDate
                $humanReadableMessage = $messageWithoutDate
            } else {
                Write-Output "  No selection data found to process"
            }

            $colUpdates = @(
                @(2649, $filterString),
                @(3308, $originalSelection),
                @(3135, $simplifiedMessage),
                @(3136, $humanReadableMessage)
            )
            if ($seedHistoryString) { $colUpdates += ,@(3444, $seedHistoryString) }
            $batchSuccess = Send-BatchUpdateCustomColumns -DeviceId $device.AccountID -Updates $colUpdates
            $colCount = $colUpdates.Count
            if ($batchSuccess) {
                Write-Output "  Successfully updated $colCount custom columns in batch"
            } else {
                Write-Warning "  Failed to batch update custom columns for device $($device.DeviceName)"
            }

            if ($Script:Export) {
                $Script:progressExportRows.Add([PSCustomObject]@{
                    AccountID = $device.AccountID; PartnerID = $device.PartnerID
                    DeviceName = $device.DeviceName; ComputerName = $device.ComputerName
                    Account = $device.Account; PartnerName = $device.PartnerName
                    OSType = $device.OSType; OS = $device.OS; Physicality = $device.Physicality
                    ProductID = $device.ProductID; ProductName = $device.Product
                    ProfileID = $device.ProfileID; ProfileName = $device.Profile; ProfileVersion = $device.ProfileVersion
                    CreationDate = $device.Creation; LastStatusTime = $device.TimeStamp
                    LastSuccess = $device.LastSuccess; LastStatus = $device.LastStatus; TotalErrors = $device.TotalErrors
                    BackupHistory = $device.Last28; SelectedGB = $device.SelectedGB; UsedGB = $device.UsedGB
                    ActiveDataSources = $device.ActiveDataSources; LSVEnabled = $device.LSVEnabled; LSVStatus = $device.LSVStatus
                    OriginalSelections = Format-CsvSafe "$originalSelection"
                    SimplifiedSelections = Format-CsvSafe "$simplifiedMessage"
                    HumanReadable = Format-CsvSafe "$humanReadableMessage"
                    Filters = Format-CsvSafe "$filterString"
                    Includes = Format-CsvSafe ($device.Includes); Excludes = Format-CsvSafe ($device.Excludes)
                    Vols = Format-CsvSafe ($device.Vols); USB_Vols = $device.USB_Vols
                    MissedVols = Format-CsvSafe ($device.MissedVols)
                    DetectedDataSources = Format-CsvSafe ($device.DetectedDataSources)
                    EncryptionStatus = $device.EncryptionStatus; FIPSStatus = $device.FIPSStatus
                    FIPSDetails = $device.FIPSDetails; mTLSStatus = $device.mTLSStatus; mTLSDetails = $device.mTLSDetails
                    Location = $device.Location; DRaaS = $device.DRaaS; SeedHistory = if ($seedHistoryString) { $seedHistoryString } else { $device.SeedHistory }
                    ColumnUpdateSuccess = $batchSuccess
                })
            }

            $Script:deviceCounter++
            if ($Script:deviceCounter % $Script:batchSize -eq 0) {
                $elapsedTime      = (Get-Date) - $Script:startTime
                $devicesRemaining = $Script:totalDevices - $Script:deviceCounter
                $devicesPerMinute = [math]::Round($Script:deviceCounter / $elapsedTime.TotalMinutes, 2)
                $eta              = if ($devicesPerMinute -gt 0) { [TimeSpan]::FromMinutes($devicesRemaining / $devicesPerMinute) } else { [TimeSpan]::Zero }
                Write-Output ""
                Write-Output $Script:strLineSeparator
                Write-Output "  Batch Save Checkpoint - $Script:deviceCounter of $Script:totalDevices devices processed"
                Write-Output "  Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))"
                Write-Output "  Devices Remaining: $devicesRemaining"
                Write-Output "  Average Rate: $devicesPerMinute devices/min"
                Write-Output "  Estimated Time Remaining: $($eta.ToString('hh\:mm\:ss'))"
                Write-Output $Script:strLineSeparator
                Write-Output ""
                if ($Script:SSLTestLog) {
                    $Script:SSLTestLog | Out-File -FilePath $Script:logFileName -Encoding UTF8
                }
                if ($Script:Export -and $Script:progressExportRows.Count -gt 0) {
                    if (Test-Path $Script:exportCsvPath) {
                        $Script:progressExportRows | Select-Object $Script:exportColumnOrder |
                            Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -Append
                    } else {
                        $Script:progressExportRows | Select-Object $Script:exportColumnOrder |
                            Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
                    }
                    $Script:progressExportRows.Clear()
                }
            }

        } ## End ForEach Loop (sequential)
    }
}


## Generate Final Export

Write-output $Script:strLineSeparator

# Final metrics and log flush (sequential path: append remaining SSLTestLog entries)
if ($Script:SSLTestLog -and $Script:SSLTestLog.Length -gt 200) {
    $Script:SSLTestLog | Out-File -FilePath $Script:logFileName -Encoding UTF8
}
$finalElapsedTime      = (Get-Date) - $Script:startTime
$finalDevicesPerMinute = if ($finalElapsedTime.TotalMinutes -gt 0) { [math]::Round($Script:deviceCounter / $finalElapsedTime.TotalMinutes, 2) } else { 0 }
Write-Output ""
Write-Output $Script:strLineSeparator
Write-Output "  Processing Complete!"
Write-Output "  SSLTest Analysis Log = $Script:logFileName"
Write-Output "  Total Devices Processed: $Script:deviceCounter of $Script:totalDevices"
Write-Output "  Total Elapsed Time: $($finalElapsedTime.ToString('hh\:mm\:ss'))"
Write-Output "  Average Rate: $finalDevicesPerMinute devices/min"
Write-Output $Script:strLineSeparator

if ($Script:Export) {
    Write-Output ""
    Write-Output $Script:strLineSeparator
    # Flush any remaining sequential rows not yet written at a checkpoint
    if ($Script:progressExportRows.Count -gt 0) {
        if (Test-Path $Script:exportCsvPath) {
            $Script:progressExportRows | Select-Object $Script:exportColumnOrder |
                Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -Append
        } else {
            $Script:progressExportRows | Select-Object $Script:exportColumnOrder |
                Export-Csv -Path $Script:exportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
        }
        $Script:progressExportRows.Clear()
    }
    $exportedCount = (Import-Csv $Script:exportCsvPath -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Output "  Export devices: $exportedCount"
    Write-Output "  Export CSV  = $Script:exportCsvPath"
    if (Test-Path HKLM:SOFTWARE\Classes\Excel.Application) {
        Save-CSVasExcel $Script:exportCsvPath
        Write-Output "  Export XLSX = $Script:exportXlsxPath"
    }
    if ($Launch) {
        if (Test-Path HKLM:SOFTWARE\Classes\Excel.Application) {
            Start-Process $Script:exportXlsxPath
            Write-Output "  Opening XLSX: $Script:exportXlsxPath"
        } else {
            Start-Process $Script:exportCsvPath
            Write-Output "  Opening CSV: $Script:exportCsvPath"
        }
    }
    Write-Output $Script:strLineSeparator
}

Write-Output ""

Start-Sleep -seconds 10
#endregion ----- Main Script ----