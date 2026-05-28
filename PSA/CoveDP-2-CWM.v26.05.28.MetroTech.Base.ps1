<# ----- About: ----
    # N-able Cove Data Protection to ConnectWise Manage PSA Integration using the Cove Data Protection MaxValue Plus Usage Report
    # Revision v26.05.28 - MetroTech: Added alphabetical sorting and green coloring for Cove products in read-only usage snapshots
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # https://github.com/backupNerd
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
    # For use with N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Get Current Device Statistics
    # Get Cove Maximum Value Usage Report for Period
    # Export usage to CSV/XLS & Pivot Table
    # Optionally Import Usage from CSV
    # Connect and update usage in ConnectWise Manage (CWM) PSA
    #
    # Use the -Period switch parameter to define Usage Dates (yyyy-MM or MM-yyyy)
    # Use the -Last switch parameter to define the number of months to count back (default=1) i.e. 0 current, 1 prior month
    # Use the -DeviceCount ## (default=15000) parameter to define the maximum number of devices returned
    # Use the -TrialState $true switch to treat trial partners as billable
    # Use the -PartnerName parameter to override the credential-stored partner for partner lookup (useful when stored credential partner is too high-level)
    # Use the -RunMode parameter to choose Interactive, PullOnly, UploadOnly, or PullAndUpload workflow execution
    # Use the -NoCWMUpdate switch (default: $true) as WhatIf-style testing to perform CWM authentication and lookup checks without updating CWM quantities
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherlands)
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path. Default is the execution path of the script.
    # Use the -DebugCDP switch to display debug info for Cove Data Protection usage and statistics lookup
    # Use the -DebugCWM switch to display debug info for ConnectWise Manage data lookup
    # Use the -GetCDPUsage switch to pull Cove HWM usage data from product
    # Use the -LoadCDPUsage switch to load Cove HWM usage data from a CSV file
    # Use the -ConnectCWM $true switch to test connection to ConnectWise Manage (Must be $true to generate an importable CSV file)
    # Use the -NoCWMUpdate switch to control write behavior: $true = what-if/read-only (default), $false = apply quantity updates
    # Use the -CWMAgreementBehavior parameter to specify how to look up CWM Agreements: 'Type' (default) matches by agreement type; 'Name' matches by agreement name
    # Use the -CWMAgreementTypes parameter to specify the list of CWM Agreement type names to match when CWMAgreementBehavior = 'Type'
    # Use the -CWMAgreementName parameter to specify the ConnectWise Manage Agreement Name to match when CWMAgreementBehavior = 'Name'
    # Use the -CWMAgreementSearchType parameter to specify how to match the CWM Agreement Name or Type (Equals, Contains, StartsWith, EndsWith)
    # Use the -CWMAdditionSearchType parameter to specify how to match the ConnectWise Manage Addition product ID (Equals or StartsWith)
    # Use the -CombineServers switch to combine Physical + Virtual server quantities into a single CWM addition (default=$true); when $true, Phys/Virt additions are zeroed
    # Use the -PhysServerFUGB, -VirtServerFUGB, and -WorkstationFUGB parameters to set the Fair Use GB for Physical Servers, Virtual Servers, and Workstations respectively
    # Edit the $CWMExcludedCompanies array (near line 140) to list CWM company names that should display usage but never have CWM additions updated

    # Use the -ClearCDPCredentials parameter to remove encrypted Cove API credentials at start of script
    # Use the -ClearCWMCredentials parameter to remove encrypted ConnectWise Manage API credentials from the local machine
    # Use the -IgnoreDeletedDevices switch to exclude devices with a DeviceDeletionDate from quantity totals and CWM updates

    # In Interactive mode, set GetCDPUsage and LoadCDPUsage $false to be prompted for Pull, Upload, Combined, or Exit
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console/export.htm
    #
    # Special thanks to Chris Taylor for the PS Module for ConnectWise Manage.
    # https://www.powershellgallery.com/packages/ConnectWiseManageAPI
    # https://github.com/christaylorcodes/ConnectWiseManageAPI

# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)][datetime]$Period,                                           ## Lookup Date yyyy-MM or MM-yyyy
    [Parameter(Mandatory=$False)][ValidateRange(0,24)]$Last = 0,                              ## Count back # Last Months i.e. 0 current, 1 Prior Month
    [Parameter(Mandatory=$False)][Alias("PartnerName")][String]$PartnerNameOverride = "",    ## Optional partner override used for Send-GetPartnerInfo when cred partner is too high-level
    [Parameter(Mandatory=$False)][int]$DeviceCount = 5000,                                    ## Change Maximum Number of backup device results to return
    [Parameter(Mandatory=$False)][string]$DeviceFilter = "AT==1 OR AT==2",                    ## Filter devices returned: AT==1 (Backup Manager), AT==2 (M365), or both
    [Parameter(Mandatory=$False)][Switch]$CombineServers = $true,                             ## Combine Physical + Virtual server qtys into a single CWM addition (CombinedServerQty). When $true, Phys/Virt additions are zeroed.

    [Parameter(Mandatory=$False)][ValidateSet("Interactive","PullOnly","UploadOnly","PullAndUpload")][String]$RunMode = "Interactive", ## Workflow mode. Interactive prompts; others are unattended and fail fast on invalid combos.
    [Parameter(Mandatory=$False)][Switch]$GetCDPUsage = $false,                               ## Set $true to Pull usage from Product, else prompt for Get or Import 
    [Parameter(Mandatory=$False)][Switch]$LoadCDPUsage = $false,                              ## Set $true to Load usage from External CSV File, else prompt for Get or Import

    [Parameter(Mandatory=$False)][Switch]$ConnectCWM = $true,                                 ## Set $true to Test CWM API connection 
    [Parameter(Mandatory=$False)][Switch]$NoCWMUpdate = $true,                                 ## WhatIf-style dry-run CWM path: authenticate and resolve CWM lookups, but never update quantities

    [Parameter(Mandatory=$False)][string]$Delimiter = ',',                                    ## Specify ',' or ';' Delimiter for XLS & CSV file
    [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",                       ## Export Path
    [Parameter(Mandatory=$False)][switch]$Launch = $false,                                    ## Launch CDP XLS or CSV Usage file

    [Parameter(Mandatory=$False)][int]$PhysServerFUGB = 0,                                   ## Selected Size Fair Use / Included GB, 0 to pass full value
    [Parameter(Mandatory=$False)][int]$VirtServerFUGB = 0,                                   ## Selected Size Fair Use / Included GB, 0 to pass full value
    [Parameter(Mandatory=$False)][int]$WorkstationFUGB = 0,                                  ## Selected Size Fair Use / Included GB, 0 to pass full value
    
    [Parameter(Mandatory=$False)][ValidateSet("Name","Type")][String]$CWMAgreementBehavior = "Type",  ## Agreement lookup mode: 'Name' matches by CWMAgreementName; 'Type' matches by agreement type names in CWMAgreementTypes
    [Parameter(Mandatory=$False)][String]$CWMAgreementName = "CoveDataProtection",            ## ConnectWise Manage Agreement Name (Preferred = 'CoveDataProtection')
    [Parameter(Mandatory=$False)][String[]]$CWMAgreementTypes = @("Co-Managed Essential","Co-Managed Professional","Legacy Time & Materials","Managed Flat Rate (Worry Free)"),  ## Agreement type names to match when CWMAgreementBehavior = 'Type'
      
    [Parameter(Mandatory=$False)][String]$CWMAgreementSearchType = "Equals",                  ## ConnectWise Manage Agreement Search Type. Can be either Equals, Contains, StartsWith, EndsWith
    [Parameter(Mandatory=$False)][String]$CWMAdditionSearchType = "Equals",                   ## ConnectWise Manage Addition Search Type. Can be either Equals or StartsWith

    [Parameter(Mandatory=$False)][Switch]$DebugCDP = $true,                                  ## Enable Debug for Cove Max Value / Device Statistics
    [Parameter(Mandatory=$False)][Switch]$DebugCWM = $true,                                   ## Enable Debug for ConnectWise Manage
    [Parameter(Mandatory=$False)][Switch]$TrialState = $false,                                ## Treat partners in trial state as Billable when set $true
    [Parameter(Mandatory=$False)][Switch]$IgnoreDeletedDevices = $true,                       ## Exclude devices with a DeviceDeletionDate from quantities and CWM totals

    [Parameter(Mandatory=$False)][switch]$ClearCDPCredentials = $false,                       ## Remove Stored Cove API Credentials at start of script
    [Parameter(Mandatory=$False)][switch]$ClearCWMCredentials = $false                       ## Remove Stored CWM API Credentials at start of script


)

#region ----- Environment, Variables, Names and Paths ----

#region ----- Cove Products to Addition Mapping  ----
$CDP2CWMProductMapping = @{

    ## Cove Data Protection Product ID >>> ConnectWise Manage Product ID (Products must be manually created in ConnectWise Manage)

    "PhysicalServerQty"     = "CDP-Physical-Servers"    ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Physical-Servers', Product Description = 'Cove Data Protection - Physical Server Backup')
    "VirtualServerQty"      = "CDP-Virtual-Servers"     ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Virtual-Servers', Product Description = 'Cove Data Protection - Virtual Server Backup')
    "CombinedServerQty"     = "Cove Server Backup with NAS"    ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Total-Servers', Product Description = 'Cove Data Protection - Total Server Backup')
    "WorkstationQty"        = "Cove Workstation Backup Agents"        ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Workstations', Product Description = 'Cove Data Protection - Workstation Professional Backup')
    "DocumentsQty"          = "CDP-Documents"           ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Documents', Product Description = 'Cove Data Protection - Workstation Documents Backup')
    "O365UsersQty"          = "CDP-M365-Users"          ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-M365-Users', Product Description = 'Cove Data Protection - M365 User Backup')
    "RecoveryTestingQty"    = "CDP-Continuity"          ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Continuity', Product Description = 'Cove Data Protection - Continuity License')      
    "CoveGBSelected"        = "CDP-GBSelected-Size"     ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-GBSelected-Size', Product Description = 'Cove Data Protection - GB Selected Size')
        
    ## 
    ## If you need to have different prices for a product (E.g. Not for Profits, Education, Government) then create products which use above Product ID's as a prefix
    ## followed by a unique suffix. (E.g. CDP-Physical-Servers-NFP, CDP-Physical-Servers-Edu, CDP-Physical-Servers-Gov).
    ## 
    ## Note: Due to the way the script looks up Product Additions you can't have multiple active Additions with the same Product ID (if $CWMAdditionSearchType = "Equals" ) 
    ## or Product ID prefix (if $CWMAdditionSearchType = "StartsWith" ).
    ## 

}  ## Cove to PSA Product Mapping
#endregion ----- Cove Products to Addition Mapping  ----

#region ----- Excluded CWM Companies  ----
$CWMExcludedCompanies = @(
    ## CWM companies listed here will have Cove usage displayed but CWM additions will NOT be updated.
    ## Matching is exact and case-insensitive against the resolved CWM company name.

    '5 Percent Nutrition'
    'Metro-Internal'

)
#endregion ----- Excluded CWM Companies  ----

Clear-Host
#Requires -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection to ConnectWise Manage PSA Integration"
$host.UI.RawUI.WindowTitle = $ConsoleTitle      ## Comment out for Automation Policy Use
Write-Output "$ConsoleTitle`n`n$($myInvocation.MyCommand.Path)"
$Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
if ($myInvocation.MyCommand.Path) { Split-path $MyInvocation.MyCommand.Path | Push-Location } ## Set terminal path to match script location (does not support editor 'Run Selection')
if ($exportpath -eq "") {$exportpath = Split-Path $MyInvocation.MyCommand.Path}

Function Display-CurrentParameters ($ParamLabel) {
    $Parameters = [ordered]@{
        Period                  = $Period
        MonthsPrior             = $Last
        PartnerNameOverride     = $PartnerNameOverride
        DeviceCount             = $DeviceCount
        DeviceFilter            = $DeviceFilter
        RunMode                 = $RunMode
        ConnectCWM              = $ConnectCWM
        GetCDPUsage             = $GetCDPUsage
        LoadCDPUsage            = $LoadCDPUsage
        NoCWMUpdate             = $NoCWMUpdate
        Launch                  = $Launch
        Delimiter               = $Delimiter
        ExportPath              = $ExportPath
        PhysServerFUGB          = "$PhysServerFUGB GB"
        VirtServerFUGB          = "$VirtServerFUGB GB"
        WorkstationFUGB         = "$WorkstationFUGB GB"
        CWMAgreementName        = $CWMAgreementName
        CWMAgreementSearchType  = $CWMAgreementSearchType
        CWMAdditionSearchType   = $CWMAdditionSearchType
        CWMAgreementBehavior    = $CWMAgreementBehavior
        CWMAgreementTypes       = ($CWMAgreementTypes -join ', ')
        CWMExcludedCompanies    = ($CWMExcludedCompanies -join ', ')
        CombineServers          = $CombineServers
        ClearCDPCredentials     = $ClearCDPCredentials
        ClearCWMCredentials     = $ClearCWMCredentials
        IgnoreDeletedDevices    = $IgnoreDeletedDevices
        TrialState              = $TrialState
        DebugCDP                = $DebugCDP
        DebugCWM                = $DebugCWM
    }

    Write-Output "`n  $ParamLabel"
    $maxKeyLength = ($Parameters.Keys | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $Parameters.GetEnumerator() | ForEach-Object {
        Write-Output ("    - {0,-$maxKeyLength} = {1}" -f $_.Key, $_.Value)
    }
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------------"
$CurrentDate = Get-Date -format "yy-MM-dd_HH-mm"
$urljson = "https://api.backup.management/jsonapi"
$ShouldProcessCWMAdditions = $false

#region ----- Workflow mode and validation ----

$IsInteractiveMode = $RunMode -eq "Interactive"

switch ($RunMode) {
    "PullOnly" {
        $GetCDPUsage = $true
        $LoadCDPUsage = $false
        $ShouldProcessCWMAdditions = $false
    }
    "UploadOnly" {
        $GetCDPUsage = $false
        $LoadCDPUsage = $true
        $ShouldProcessCWMAdditions = $true
    }
    "PullAndUpload" {
        $GetCDPUsage = $true
        $LoadCDPUsage = $false
        $ShouldProcessCWMAdditions = $true
    }
    "Interactive" {
        ## Respect explicitly passed switches and only prompt when neither pull nor upload is selected.
        $ShouldProcessCWMAdditions = $LoadCDPUsage
    }
}

if ($NoCWMUpdate -and $ShouldProcessCWMAdditions) {
    Write-Host "  [NoCWMUpdate] CWM quantity updates disabled; running CWM connectivity, lookup, and read-only addition snapshot workflow." -ForegroundColor Yellow
}

Display-CurrentParameters "Current Parameters:"

if ($LoadCDPUsage -and $GetCDPUsage) {Write-warning " Script Switch parameters 'GetCDPUsage' and 'LoadCDPUsage' can not be used together, exiting script.";break}
## If both GetCDPUsage and LoadCDPUsage are set, exit the script
if ($ShouldProcessCWMAdditions -and -not $ConnectCWM) {Write-warning " Workflow selection requires 'ConnectCWM' to be enabled, exiting script.";break}
## Any workflow that processes CWM additions requires ConnectCWM

if (-not $LoadCDPUsage -and -not $GetCDPUsage) {
    if (-not $IsInteractiveMode) {
        Write-warning " RunMode '$RunMode' is unattended and requires a valid workflow selection; no interactive prompts are allowed. Exiting script."
        break
    }

    do {
        $choice = Read-Host "`nNo workflow selected. `nDo you want to: `n`n (P)ull Cove usage only`n (U)pload usage to CWM from CSV`n (C)ombined Pull + Upload`n e(X)it?`n`nEnter 'P', 'U', 'C', or 'X'"
        switch ($choice.ToUpper()) {
            'P' {
                $GetCDPUsage = $true
                $LoadCDPUsage = $false
                $ShouldProcessCWMAdditions = $false
                break
            }
            'U' {
                $GetCDPUsage = $false
                $LoadCDPUsage = $true
                $ShouldProcessCWMAdditions = $true
                break
            }
            'C' {
                $GetCDPUsage = $true
                $LoadCDPUsage = $false
                $ShouldProcessCWMAdditions = $true
                break
            }
            'X' {
                Write-Output "Exiting script."
                exit
            }
            default {
                Write-Output "Invalid choice, please enter 'P', 'U', 'C' or 'X'."
            }
        }
    } until ($choice.ToUpper() -eq 'P' -or $choice.ToUpper() -eq 'U' -or $choice.ToUpper() -eq 'C' -or $choice.ToUpper() -eq 'X')

    if ($NoCWMUpdate -and $ShouldProcessCWMAdditions) {
        Write-Host "  [NoCWMUpdate] CWM quantity updates disabled; running CWM connectivity, lookup, and read-only addition snapshot workflow." -ForegroundColor Yellow
    }
}

#endregion ----- Workflow mode and validation ----

#endregion ----- Environment, Variables, Names and Paths ----

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

    Write-Output "  Enter Exact, Case Sensitive Partner Name for Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for Backup.Management API'

    ## v10 unified XML format - PartnerName and Username stored as properties, Password DPAPI encrypted
    $CDPCredentials = [PSCustomObject]@{
        PartnerName = $Script:PartnerName
        Username    = $BackupCred.UserName
        Password    = ($BackupCred.Password | ConvertFrom-SecureString)  ## DPAPI encrypted
    }
    $CDPCredentials | Export-Clixml -Path $APIcredfile -Force

    Start-Sleep -milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present (v10 unified XML format)

Function Get-APICredentials {

    $Script:True_path          = "C:\ProgramData\MXB\"
    $Script:APIcredfile        = Join-Path -Path $True_Path -ChildPath "${env:computername}_${env:username}_API_Credentials.Secure.xml"  ## v10 unified XML format
    $Script:APIcredpath        = Split-Path -Path $APIcredfile

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
            ## Load from v10 unified XML format
            Write-Output    $Script:strLineSeparator
            Write-Output "  Backup API Credential File Present"
            $APIcredentials = Import-Clixml -Path $APIcredfile

            $Script:cred0 = $APIcredentials.PartnerName
            $Script:cred1 = $APIcredentials.Username
            $Script:cred2 = $APIcredentials.Password | ConvertTo-SecureString  ## Kept as SecureString - decrypted only at point of use

            Write-Output    $Script:strLineSeparator
            Write-Output "  Stored Backup API Partner  = $Script:cred0"
            Write-Output "  Stored Backup API User     = $Script:cred1"
            Write-Output "  Stored Backup API Password = Encrypted"

        } else {
            Write-Output    $Script:strLineSeparator
            Write-Output "  Backup API Credential File Not Present"

            Set-APICredentials  ## Create API Credential File if Not Found
        }
    }
}  ## Get API credentials if present (v10 unified XML format)

Function Send-APICredentialsCookie {

    Get-APICredentials  ## Read API Credential File before Authentication

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.username = $Script:cred1
    $Script:_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2)
    try   { $plainPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto($Script:_bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Script:_bstr); $Script:_bstr = $null }
    $data.params.password = $plainPwd

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url `
        -TimeoutSec 30 `
        -SessionVariable Script:websession `
        -UseBasicParsing
    $data.params.password = $null  ## Clear plaintext password from request object
    $plainPwd = $null              ## Clear plaintext password variable
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:Authenticate = $webrequest | convertfrom-json

    #Debug Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($authenticate.visa) {
        $Script:visa = $authenticate.visa
        $Script:UserId = $authenticate.result.result.id
    }else{
        Write-Output    $Script:strLineSeparator
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator

        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Get-VisaTime {
    if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
            $visatime
            Send-APICredentialsCookie
        }
    }
}  ## Renew Visa

#endregion ----- Authentication ----

#region ----- Data Conversion ----
Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $DateTime = $epoch.AddSeconds($inputUnixTime)
    return $DateTime
    }else{ return ""}
}  ## Convert epoch time to date time

Function Convert-DateTimeToUnixTime($DateToConvert) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $NewExtensionDate = Get-Date -Date $DateToConvert
    [int64]$NewEpoch = (New-TimeSpan -Start $epoch -End $NewExtensionDate).TotalSeconds
    Return $NewEpoch
}  ## Convert date time to epoch time

Function Get-Period {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $TrialForm = New-Object Windows.Forms.Form
    $TrialForm.text = "Select Month and Year for Report"
    $TrialForm.Font = [System.Drawing.Font]::new('Arial',15, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
    $TrialForm.BackColor = [Drawing.Color]::SteelBlue
    $TrialForm.AutoSize = $False
    $TrialForm.MaximizeBox = $False
    $TrialForm.Size = New-Object Drawing.Size(750,350)
    $TrialForm.ControlBox = $True
    $TrialForm.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    $TrialForm.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(10,220)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $TrialForm.AcceptButton = $okButton
    $TrialForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(85,220)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $TrialForm.CancelButton = $cancelButton
    $TrialForm.Controls.Add($cancelButton)

    $Calendar = New-Object System.Windows.Forms.MonthCalendar
    $Calendar.Location = "10,10"
    $Calendar.MaxSelectionCount = 1
    $Calendar.MinDate = (Get-Date).AddMonths(-24)   # Minimum Date Displayed
    $Calendar.MaxDate = (Get-Date)
    $Calendar.SetCalendarDimensions([int]3,[int]1)  # 3x1 Grid
    $TrialForm.Controls.Add($Calendar)

    $TrialForm.Add_Shown($TrialForm.Activate())
    $result = $TrialForm.showdialog()
    #$Calendar.SelectionEnd

    if ($result -eq [Windows.Forms.DialogResult]::OK) {
        $Script:Period = $calendar.SelectionEnd
        Write-Output "Date selected: $(get-date($Period) -Format 'MMMM yyyy')"
    }

    if ($result -eq [Windows.Forms.DialogResult]::Cancel) {
        $Script:Period = get-date
    }
}  ## GUI Select Report Period

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

Function Send-GetPartnerInfo ($PartnerName, [int]$Depth = 0) {
    if ($Depth -ge 3) {
        Write-Error "  GetPartnerInfo: Maximum retry attempts (3) reached. Verify partner name and credentials."
        exit 1
    }

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfo'
    $data.params = @{}
    $data.params.name = [String]$PartnerName

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
    #$Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:Partner = $webrequest | convertfrom-json

    $RestrictedPartnerLevel = @("Root","SubRoot","Distributor")

    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$Script:Level = $Partner.result.result.Level
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-Output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
    } else {
        Write-Output $Script:strLineSeparator
        Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername ($Depth + 1)
    }

    if ($partner.error) {
        Write-Output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername ($Depth + 1)
    }
}  ## get PartnerID and Partner Level

Function Send-GetPartnerInfoByID ($PartnerId) {

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfoById'
    $data.params = @{}
    $data.params.partnerId = [String]$PartnerId

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:PartnerHistory = $webrequest | convertfrom-json


}  ## get PartnerID and Partner Level

Function Send-GetPartnerInfoHistory ([int]$PartnerId) {

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfoHistory'
    $data.params = @{}
    $data.params.partnerId = [int]$PartnerId
    $data.params.fields = (0,1,4,7)

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        [array]$PartnerHistory = $webrequest | convertfrom-json

        if ($PartnerHistory.result.result) {
       
            [int]$Script:HistoryParentid = $PartnerHistory.result.result[-1].Parentid[0]
            [int]$Script:HistoryPartnerId = $PartnerHistory.result[-1].result.Id[0]
            [String]$Script:HistoryLevel = $PartnerHistory.result.result[-1].Level[0]
            [String]$Script:HistoryPartnerName = $PartnerHistory.result.result[-1].Name[0]

            Write-Output $Script:strLineSeparator
            Write-Output "Parent - $Historyparentid - $Historylevel - $HistorypartnerId - $HistoryPartnerName"
            Write-Output $Script:strLineSeparator
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
    return $reader.ReadToEnd()| ConvertFrom-Json
    $reader.Close()
}  ## Generic Json Call

Function Send-EnumeratePartners {
    # ----- Get Partners via EnumeratePartners -----

    # (Create the JSON object to call the EnumeratePartners function)
    $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
        Add-Member -PassThru NoteProperty visa $Script:visa |
        Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
        Add-Member -PassThru NoteProperty params @{ parentPartnerId = $PartnerId
                                                    fetchRecursively = "false"
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
    # Break Script
    } else { # (No error)
        $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore

        $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
        $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
        $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}

        $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}

        $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} )


        if ($AllPartners) {
            $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
            Write-Output    $Script:strLineSeparator
            Write-Output    "  All Partners Selected"
        } else {
            $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner - $partnername" -OutputMode Single

            if($null -eq $Selection) {
                # Cancel was pressed
                # Run cancel script
                Write-Output    $Script:strLineSeparator
                Write-Output    "  No Partners Selected"
                Exit
            } else {
                # OK was pressed, $Selection contains what was chosen
                # Run OK script
                [int]$Script:PartnerId = $Script:Selection.Id
                [String]$Script:PartnerName = $Script:Selection.Name
            }
        }
    }

}  ## EnumeratePartners API Call

Function Send-GetDevices ($PartnerID, $DeviceFilter = $DeviceFilter) {

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $visa
    $data.method = 'EnumerateAccountStatistics'
    $data.params = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId = [int]$PartnerID
    $data.params.query.Filter = $DeviceFilter
    $data.params.query.Columns = @("AU","AR","AN","MN","AL","LN","OP","OI","OS","OT","PD","AP","PF","PN","CD","TS","TL","T3","US","I81","AA843")
    $data.params.query.OrderBy = "CD DESC"
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
    $Script:devTotal  = @($DeviceResponse.result.result).Count
    $Script:devIndex  = 0
    $Script:lastDevId = ''
    Write-Host "  Processing device statistics ..." -NoNewline -ForegroundColor Gray

    ForEach ( $DeviceResult in $DeviceResponse.result.result ) {
        $Script:devIndex++
        $Script:lastDevId = $DeviceResult.AccountId
        Write-Host ("`r  Processing device statistics  [ {0,4} / {1} ]  ID: {2}   " -f $Script:devIndex, $Script:devTotal, $Script:lastDevId) -NoNewline -ForegroundColor Gray
        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [Int]$DeviceResult.AccountId;
                                                                            PartnerID      = [string]$DeviceResult.PartnerId;
                                                                            DeviceName     = $DeviceResult.Settings.AN -join '' ;
                                                                            ComputerName   = $DeviceResult.Settings.MN -join '' ;
                                                                            DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                                                                            PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                            Reference      = $DeviceResult.Settings.PF -join '' ;
                                                                            Creation       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                                                                            TimeStamp      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;
                                                                            LastSuccess    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '') ;
                                                                            SelectedGB     = [Math]::Round([Decimal](($DeviceResult.Settings.T3 -join '') /1GB),3) ;
                                                                            UsedGB         = [Math]::Round([Decimal](($DeviceResult.Settings.US -join '') /1GB),3) ;
                                                                            DataSources    = $DeviceResult.Settings.AP -join '' ;
                                                                            Account        = $DeviceResult.Settings.AU -join '' ;
                                                                            Location       = $DeviceResult.Settings.LN -join '' ;
                                                                            Notes          = $DeviceResult.Settings.AA843 -join '' ;
                                                                            Product        = $DeviceResult.Settings.PN -join '' ;
                                                                            ProductID      = $DeviceResult.Settings.PD -join '' ;
                                                                            Profile        = $DeviceResult.Settings.OP -join '' ;
                                                                            OS             = $DeviceResult.Settings.OS -join '' ;
                                                                            OSType         = $DeviceResult.Settings.OT -join '' ;
                                                                            Physicality    = $DeviceResult.Settings.I81 -join '' ;
                                                                            ProfileID      = $DeviceResult.Settings.OI -join '' }
    }
    Write-Host ("`r  Processing device statistics  [ {0} / {0} ]  Done                              " -f $Script:devTotal) -ForegroundColor Green

}  ## EnumerateAccountStatistics API Call

Function Send-GetDeviceHistory ($DeviceID, $DeleteDate) {

    #$history = (get-date $Period)
    Get-VisaTime

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $visa
    $data.method = 'EnumerateAccountHistoryStatistics'
    $data.params = @{}
    $data.params.timeslice = Convert-DateTimeToUnixTime $DeleteDate
    $data.params.query = @{}
    $data.params.query.PartnerId = [int]$PartnerId
    $data.params.query.Filter = "(AU == $DeviceID)"
    $data.params.query.Columns = @("AU","AR","AN","MN","AL","LN","OP","OI","OS","OT","PD","AP","PF","PN","CD","TS","TL","T3","US","I81","AA843")
    $data.params.query.OrderBy = "CD DESC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = 1
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

    $Script:DeviceHistoryDetail = @()

    ForEach ( $DeviceHistory in $DeviceResponse.result.result ) {
        ## Progress line is handled in Join-Reports
        $Script:DeviceHistoryDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [Int]$DeviceHistory.AccountId;
                                                                                    PartnerID      = [string]$DeviceHistory.PartnerId;
                                                                                    DeviceName     = $DeviceHistory.Settings.AN -join '' ;
                                                                                    ComputerName   = $DeviceHistory.Settings.MN -join '' ;
                                                                                    DeviceAlias    = $DeviceHistory.Settings.AL -join '' ;
                                                                                    PartnerName    = $DeviceHistory.Settings.AR -join '' ;
                                                                                    Reference      = $DeviceHistory.Settings.PF -join '' ;
                                                                                    Creation       = Convert-UnixTimeToDateTime ($DeviceHistory.Settings.CD -join '') ;
                                                                                    TimeStamp      = Convert-UnixTimeToDateTime ($DeviceHistory.Settings.TS -join '') ;
                                                                                    LastSuccess    = Convert-UnixTimeToDateTime ($DeviceHistory.Settings.TL -join '') ;
                                                                                    SelectedGB     = [Math]::Round([Decimal](($DeviceHistory.Settings.T3 -join '') /1GB),3) ;
                                                                                    UsedGB         = [Math]::Round([Decimal](($DeviceHistory.Settings.US -join '') /1GB),3) ;
                                                                                    DataSources    = $DeviceHistory.Settings.AP -join '' ;
                                                                                    Account        = $DeviceHistory.Settings.AU -join '' ;
                                                                                    Location       = $DeviceHistory.Settings.LN -join '' ;
                                                                                    Notes          = $DeviceHistory.Settings.AA843 -join '' ;
                                                                                    Product        = $DeviceHistory.Settings.PN -join '' ;
                                                                                    ProductID      = $DeviceHistory.Settings.PD -join '' ;
                                                                                    Profile        = $DeviceHistory.Settings.OP -join '' ;
                                                                                    OS             = $DeviceHistory.Settings.OS -join '' ;
                                                                                    OSType         = $DeviceHistory.Settings.OT -join '' ;
                                                                                    Physicality    = $DeviceHistory.Settings.I81 -join '' ;
                                                                                    ProfileID      = $DeviceHistory.Settings.OI -join '' }
    }
}  ## EnumerateAccountStatistics API Call

Function GetMVReport {

    Param ([Parameter(Mandatory=$False)][Int]$PartnerId) #end param

    $script:Start = (Get-Date $Period -Day 1).AddMonths(0).ToString('yyyy-MM-dd')
    $script:End = (Get-Date $Period -Day 1).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
    Write-Host "  Requesting Maximum Value Report from $start to $end"

    $Script:TempMVReport = "$CurrentExportPath\TempMVReport.xlsx"
    remove-item $Script:TempMVReport -Force -Recurse -ErrorAction SilentlyContinue

    $url2 = "https://api.backup.management/statistics-reporting/high-watermark-usage-report?_=6e8d1e0fce68d&dateFrom=$($Start)T00%3A00%3A00Z&dateTo=$($end)T23%3A59%3A59Z&exportOutput=OneXlsxFile&partnerId=$($PartnerId)"
    $method = 'GET'

    $params = @{
        Uri         = $url2
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        ContentType = 'application/json; charset=utf-8'
        OutFile     = $Script:TempMVReport
    }

    #Write-Output  "$url2"

    Invoke-RestMethod @params

    if (Get-Module -ListAvailable -Name ImportExcel) {
        if ($debugCDP) { Write-Host "  Module ImportExcel Already Installed" }
    } else {
        try {
            ## https://powershell.one/tricks/parsing/excel
            ## https://github.com/dfinke/ImportExcel https://github.com/dfinke/ImportExcel
            Install-Module -Name ImportExcel -Confirm:$False -Force
        }
        catch [Exception] {
            $_.message
            exit
        }
    }
    
    $Script:MVReportxls = Import-Excel -path "$Script:TempMVReport" -asDate "*Date"
    
    $Script:MVReport = $Script:MVReportxls | select-object  Disti_Id,Disti,Disti_Legal,Disti_Ref,
                                                            SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,
                                                            Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,
                                                            SO_Id,SO,SO_Legal,SO_Ref,
                                                            EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,
                                                            Site_Id,Site,Site_Legal,Site_Ref,
                                                            CustomerId,CustomerName,CustomerReference,CustomerState,ProductionDate,ServiceType,
                                                            OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,CreationDate,DeviceDeletionDate,
                                                            UsedStorageGb,UsedStorageMvDate,SelectedSizeGb,SelectedSizeMvDate,O365Users,O365UsersMvDate,RecoveryTesting,RecoveryTestingMvDate,
                                                            @{n='PhysicalServer';e='$Null'},@{n='VirtualServer';e='$Null'},@{n='Workstation';e='$Null'},@{n='Documents';e='$Null'},@{n='FairUseGB';e='$Null'},@{n='OverageGB';e='$Null'} | sort-object CustomerId
    $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "0"){$_.RecoveryTesting = $null}}
    $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "True"){$_.RecoveryTesting = "1"}}
    $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "1"){$_.RecoveryTesting = "RecoveryTesting"}}
    $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "2"){$_.RecoveryTesting = "StandbyImage"}}
    #$Script:MVReport | foreach-object { if ($_.O365Users -eq "0"){$_.O365Users = $null}}

    $Script:mvTotal = @($Script:MVReport | Where-Object { $_.customerid }).Count
    $Script:mvIndex = 0
    Write-Host "  Building partner hierarchy  ..." -NoNewline -ForegroundColor Gray

    $Script:MVReport | foreach-object {
        if ($_.customerid) {
            $Script:mvIndex++
            Write-Host ("`r  Building partner hierarchy  [ {0,4} / {1} ]  ID: {2}   " -f $Script:mvIndex, $Script:mvTotal, $_.CustomerId) -NoNewline -ForegroundColor Gray

            if ($null -eq $Script:EnumerateAncestorPartners.result.result.id) {
                send-EnumerateAncestorPartners $_.customerid
            } ## If Initial Ancestor partner values are $null then Enum Ancestors, and Report value[0] if found

            if (($Script:EnumerateAncestorPartners.result.result.id) -and ( $_.customerid -ne $Script:EnumerateAncestorPartners.result.result.id[0])) {
                send-EnumerateAncestorPartners $_.customerid
            } ## If Ancestor partner[0] is not the same as the report partnerid then Enum Ancestors
            
            if (($Script:EnumerateAncestorPartners.result.result.id.count -ge 1) -and ($_.customerid -eq $Script:EnumerateAncestorPartners.result.result.id[0])) {

                $_.Disti_id             = $distributor_level.id
                $_.Disti                = $distributor_level.name
                $_.Disti_legal          = $distributor_level.Company.LegalCompanyName
                $_.Disti_ref            = $distributor_level.ExternalCode
                
                $_.SubDisti_id          = $subdistributor_level.id
                $_.SubDisti             = $subdistributor_level.name
                $_.SubDisti_legal       = $subdistributor_level.Company.LegalCompanyName
                $_.SubDisti_ref         = $subdistributor_level.ExternalCode
            
                $_.Reseller_id          = $reseller_level.id
                $_.Reseller             = $reseller_level.name
                $_.Reseller_legal       = $reseller_level.Company.LegalCompanyName
                $_.Reseller_ref         = $reseller_level.ExternalCode

                $_.SO_id                = $SO_level.id
                $_.SO                   = $SO_level.name
                $_.SO_legal             = $SO_level.Company.LegalCompanyName
                $_.SO_ref               = $SO_level.ExternalCode
                
                $_.EndCustomer_id       = $endcustomer_level.id
                $_.EndCustomer          = $endcustomer_level.name
                $_.EndCustomer_legal    = $endcustomer_level.Company.LegalCompanyName
                $_.EndCustomer_ref      = $endcustomer_level.ExternalCode
            
                $_.Site_id              = $site_level.id
                $_.Site                 = $site_level.name
                $_.Site_legal           = $site_level.Company.LegalCompanyName
                $_.Site_ref             = $site_level.ExternalCode
            
                $loop ++
            } ## If Ancestor partner[0] is the same as the report partnerid then use cached values  
                
            if ($Script:EnumerateAncestorPartners.error.message) {

                send-GetPartnerInfoHistory $_.customerid
                
                if ($debugCDP) { Write-Output "Getting Deleted Parent data for partnerid - $($_.CustomerId)" }

                if (($HistorypartnerId) -and ($Historylevel -eq "Site") ) {
                    if ($debugCDP) {Write-Output "Found Site data for partnerid - $($_.CustomerId)" }
                    $_.Site_id              = $HistorypartnerId
                    $_.Site                 = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "EndCustomer") ) {
                    if ($debugCDP) {Write-Output "Found EndCustomer data for partnerid - $($_.CustomerId)" }
                    $_.EndCustomer_id       = $HistorypartnerId
                    $_.EndCustomer          = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "ServiceOrganization") ) {
                    if ($debugCDP) {Write-Output "Found SO data for partnerid - $($_.CustomerId)" }
                    $_.SO_id          = $HistorypartnerId
                    $_.SO             = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Reseller") ) {
                    if ($debugCDP) {Write-Output "Found Reseller data for partnerid - $($_.CustomerId)" }
                    $_.Reseller_id          = $HistorypartnerId
                    $_.Reseller             = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Subdistributor") ) {
                    if ($debugCDP) {Write-Output "Found SubDisti data for partnerid - $($_.CustomerId)" }
                    $_.SubDisti_id           = $HistorypartnerId
                    $_.SubDisti              = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Distributor") ) {
                    if ($debugCDP) {Write-Output "Found Disti data for partnerid - $($_.CustomerId)" }
                    $_.Disti_id             = $HistorypartnerId
                    $_.Disti                = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
            } ## If partner is <anonymized> or deleted then use GetPartnerInfoHistory to get parent data  
        }
    }
    Write-Host ("`r  Building partner hierarchy  [ {0} / {0} ]  Done                              " -f $Script:mvTotal) -ForegroundColor Green
}  ## Get Maximum Value Report

Function Join-Reports {

    if (Get-Module -ListAvailable -Name Join-Object) {
        if ($debugCDP) { Write-Host "  Module Join-Object Already Installed" }
    } else {
        try {
            Install-Module -Name Join-Object -Confirm:$False -Force    ## https://www.powershellgallery.com/packages/Join-Object
        }
        catch [Exception] {
            $_.message
            exit
        }
    }
    $Script:MVPlus = Join-Object -left $Script:MVReport -right $Script:SelectedDevices -LeftJoinProperty DeviceId -RightJoinProperty AccountID -Prefix 'Now_' -Type AllInBoth
    $Script:deletedDevices = @($Script:MVPlus | Where-Object {$_.DeviceDeletionDate} | Sort-Object DeviceId)
    $Script:delTotal = $Script:deletedDevices.Count
    $Script:delIndex = 0
    if ($Script:delTotal -gt 0) {
        Write-Host "  Retrieving deleted device history ..." -NoNewline -ForegroundColor Gray
    }
    $Script:deletedDevices | foreach-object {
        $Script:delIndex++
        Write-Host ("`r  Retrieving deleted device history [ {0,4} / {1} ]  ID: {2}   " -f $Script:delIndex, $Script:delTotal, $_.DeviceId) -NoNewline -ForegroundColor Gray
        Send-GetDeviceHistory $_.deviceid $_.DeviceDeletionDate
        $_.Now_Physicality = $Script:DeviceHistoryDetail.physicality
        $_.Now_Product     = $Script:DeviceHistoryDetail.Product
        $_.Now_timestamp   = $Script:DeviceHistoryDetail.timestamp
        $_.Now_Reference   = $Script:DeviceHistoryDetail.Reference
    }
    if ($Script:delTotal -gt 0) {
        Write-Host ("`r  Retrieving deleted device history [ {0} / {0} ]  Done                              " -f $Script:delTotal) -ForegroundColor Green
    }
}  ## Install Join-Object PS module to merge statistics and Max Value Report

Function Send-EnumerateAncestorPartners ($PartnerID) {
    Get-VisaTime

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'EnumerateAncestorPartners'
    $data.params = @{}
    $data.params.partnerId = $PartnerId

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }

    [array]$Script:EnumerateAncestorPartners = Invoke-RestMethod @params

    $script:subroot_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "SubRoot"}
    $script:distributor_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Distributor"}
    $script:subdistributor_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "SubDistributor"}
    $script:reseller_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Reseller"}
    $script:SO_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "ServiceOrganization"}
    $script:Endcustomer_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "EndCustomer"}
    $script:Site_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Site"}
}  ## Get Parent Partners

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

Send-APICredentialsCookie

Write-Output $Script:strLineSeparator
Write-Output ""

if ([string]::IsNullOrWhiteSpace($PartnerNameOverride)) {
    Send-GetPartnerInfo $cred0
} else {
    Write-Output "  Using partner override from parameter: $PartnerNameOverride"
    Send-GetPartnerInfo $PartnerNameOverride
}

if (($null -eq $Script:Period) -and ($null -eq $last)) {Get-Period}
if (($null -eq $Script:Period) -and ($last -eq 0)) { $Script:Period = (get-date) }
if (($null -eq $Script:Period) -and ($last -ge 1)) { $Script:Period = ((get-date).addmonths($last/-1)) }
#$Script:Period
#$Last


$null = New-Item -Path "$ExportPath" -Name "$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)" -ItemType "directory" -Force

$Script:CurrentExportPath = "$Exportpath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)"

$Script:AuditLogFile = "$CurrentExportPath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)_AuditLog.csv"

Add-content -Path $AuditLogFile -Value "Company Id,Company Name,Agreement Name,Agreement Type,Additions Product ID,Additions Description,Additions Quantity,Updated Quantity"

$Script:ExceptionsLogFile = "$CurrentExportPath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)_ExceptionsLog.txt"

$Script:CSVOutputFile = "$CurrentExportPath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)_MaxValuePlus_DPP_$($Period.ToString('yyyy-MM'))_Statistics.csv"

$Script:CDPUsageoutputFile = "$CurrentExportPath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)_CDPUsageFile_$($Period.ToString('yyyy-MM'))_.csv"


if ($GetCDPUsage) {

    Send-GetDevices $partnerId

    $Script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,ComputerName,
                                                            Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,
                                                            SelectedGB,UsedGB,Location,OS,OSType,Physicality

    GetMVReport $partnerId
    Join-Reports

    if ($IgnoreDeletedDevices) {
        $removedCount = ($Script:MVPlus | Where-Object { $_.DeviceDeletionDate }).Count
        $Script:MVPlus = @($Script:MVPlus | Where-Object { -not $_.DeviceDeletionDate })
        if ($removedCount -gt 0) {
            Write-Host "  [IgnoreDeletedDevices] Excluded $removedCount deleted device(s) from totals." -ForegroundColor Yellow
        }
    }

    if ($null -eq $Script:SelectedDevices) {
        # Cancel was pressed
        # Run cancel script
        Write-Output  $Script:strLineSeparator
        Write-Warning "No Devices Selected`n"
        Exit
    }else{
        # OK was pressed, $Selection contains what was chosen
        # Run OK script
    }

    $Script:MVPlus | foreach-object { if (($_.Now_Physicality -eq "Undefined") -and ($_.O365Users)) {$_.Now_Physicality = "Cloud"; $_.OSType = "M365";  $_.ComputerName = "* M365 - $($_.DeviceName)"}}
    
    $Script:MVPlus | foreach-object { if ($_.SelectedsizeGB -gt 0) {$_.SelectedsizeGB = [math]::round($_.SelectedsizeGB,3)}}
    $Script:MVPlus | foreach-object { if ($_.UsedStorageGB -gt 0) {$_.UsedStorageGB = [math]::round($_.UsedStorageGB,3)}}
    
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq $null)) {$_.PhysicalServer = "1"}}
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Physical")) {$_.PhysicalServer = "1"}}#else{$_.PhysicalServer = "0"}}
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Virtual")) {$_.VirtualServer = "1"}}#else{$_.VirtualServer = "0"}}
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -eq "Documents")) {$_.Documents = "1"}}#else{$_.Documents = "0"}}
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -ne "Documents")) {$_.Workstation = "1"}}#else{$_.Workstation = "0"}}

    $Script:MVPlus | foreach-object { if ($_.OsType -eq "Server") {$_.FairUseGB = $PhysServerFUGB }}
    $Script:MVPlus | foreach-object { if ($_.OsType -eq "$null") {$_.FairUseGB = $PhysServerFUGB }}
    $Script:MVPlus | foreach-object { if ($_.PhysicalServer -eq 1) {$_.FairUseGB = $PhysServerFUGB }}
    $Script:MVPlus | foreach-object { if ($_.VirtualServer -eq 1) {$_.FairUseGB = $VirtServerFUGB }}
    $Script:MVPlus | foreach-object { if ($_.Workstation -eq 1) {$_.FairUseGB = $WorkstationFUGB }}
    $Script:MVPlus | foreach-object { if ($_.Documents -eq 1) {$_.FairUseGB = 0 ; $_.SelectedsizeGB = 0 }}
    $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -or ($_.PhysicalServer -eq 1) -or ($_.VirtualServer -eq 1) -or ($_.Workstation -eq 1) -or ($_.OsType -eq "$null") -or (($_.Documents = "$null") -and ($_.O365Users -eq 0))) {[decimal]$_.OverageGB = ([decimal]$_.SelectedSizeGB - [decimal]$_.FairuseGB) }}
    #PhysServerGBFU,VirtServerGBFU,WorkstationGBFU

    if ($TrialState) {
        $StateToProcess = "InTrial","InProduction"
    }else{
        $StateToProcess = "InProduction"
    }

    $Script:MVPlus | Where-object {$_.CustomerName -notlike '*Recycle Bin'} |
        Select-object   Disti_Id,Disti,Disti_Legal,Disti_Ref,
                        SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,
                        Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,
                        SO_Id,SO,SO_Legal,SO_Ref,
                        EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,
                        Site_Id,Site,Site_Legal,Site_Ref,
                        CustomerId,CustomerName,CustomerReference,ServiceType,
                        OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                        Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                        CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB|
        Export-CSV -path "$CSVOutputFile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8
        ## Generates the Cove CSV output file

    $Script:XLSOutputFile = $CSVOutputFile.Replace("csv","xlsx")

    $ConditionalFormat = $(New-ConditionalText -ConditionalType DuplicateValues -Range 'AC:AC' -BackgroundColor CornflowerBlue -ConditionalTextColor Black)

    $Script:MVPlus | Where-object {$_.CustomerState -eq "InTrial"} | where-object {$_.CustomerName -notlike '*Recycle Bin'} |
        Select-object   Disti_Id,Disti,Disti_Legal,Disti_Ref,
                        SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,
                        Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,
                        SO_Id,SO,SO_Legal,SO_Ref,
                        EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,
                        Site_Id,Site,Site_Legal,Site_Ref,
                        CustomerId,CustomerName,CustomerReference,ServiceType,
                        OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                        Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                        CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB |
        Export-Excel -Path "$XLSOutputFile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName TrialUsage -FreezeTopRowFirstColumn -WorksheetName "TRIAL $(get-date $Period -UFormat `"%b-%Y`")" -tablestyle Medium6 -BoldTopRow
        ## Generates the Trial Tab of the Cove XLS output file

    $pivotColumns = [ordered]@{
        PhysicalServer='sum';
        VirtualServer='sum';
        Workstation='sum';
        Documents='sum';
        RecoveryTesting='count';
        O365Users='sum';
        SelectedSizeGB='sum';
        UsedStorageGB='sum';
        FairUseGB='sum';
        OverageGB='sum'
    }

    $Script:MVPlus | Where-object {$_.CustomerState -in $StateToProcess} | where-object {$_.CustomerName -notlike '*Recycle Bin'} |
        Select-object   Disti_Id,Disti,Disti_Legal,Disti_Ref,
                        SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,
                        Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,
                        SO_Id,SO,SO_Legal,SO_Ref,
                        EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,
                        Site_Id,Site,Site_Legal,Site_Ref,
                        CustomerId,CustomerName,CustomerReference,ServiceType,
                        OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                        Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                        CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB |
        Export-Excel -Path "$XLSOutputFile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName ProductionUsage -FreezeTopRowFirstColumn -WorksheetName (get-date $Period -UFormat "%b-%Y") -tablestyle Medium6 -BoldTopRow -IncludePivotTable -PivotRows Disti,SubDisti,Reseller,SO,EndCustomer -PivotDataToColumn -PivotData $pivotColumns
        ## Generates the Production Tab of the Cove XLS output file, including a Pivot Table. Set script parameter $TrialState to $True to include Trial data

    if ($Launch) {
        if (test-path HKLM:SOFTWARE\Classes\Excel.Application) {
            Start-Process "$XLSOutputFile"
            Write-Output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
        } else {
            Start-Process "$CSVOutputFile"
            Write-Output $Script:strLineSeparator
            Write-Output "  Opening CSV file"
            Write-Output $Script:strLineSeparator
        }
    }  ## Launch CSV or XLS (if Excel is installed) (Requires -Launch Parameter)

    Write-Output $Script:strLineSeparator
    Write-Output "  Cove Max Value Plus Usage - CSV Path = `"$CSVOutputFile`""
    Write-Output "  Cove Max Value Plus Usage - XLS Path = `"$XLSOutputFile`""
    Write-Output ""

    ## ── Build per-customer usage summary and export CDPUsage CSV ──
    ## Runs here (outside $ConnectCWM) so the CDPUsage CSV is always generated when $GetCDPUsage is $True.
    ## $Script:CDPUsageCollection stores the built objects so the CWM block can use them for company/agreement lookups.
    $Script:CDPUsageCollection = @()

    $Usage = $MVplus | Where-object {$_.CustomerState -in $StateToProcess} |
        Select-Object Subdisti,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Id,
        SO_Id,SO,SO_Legal,SO_Ref,
        Endcustomer,EndCustomer_legal,Endcustomer_Ref,EndCustomer_id,
        PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,OverageGB | Sort-Object endcustomer -Unique

    # ── Console header (once) ──
    # $script:Start and $script:End are proper month-start/end strings (yyyy-MM-dd) computed earlier
    $hwmMonthLabel = "{0} to {1}" -f $script:Start, $script:End
    Write-Host ""
    Write-Host "  $('=' * 68)" -ForegroundColor Cyan
    Write-Host ("    Cove Data Protection  -  High Water Mark  -  {0}" -f $hwmMonthLabel) -ForegroundColor Cyan
    Write-Host "  $('=' * 68)" -ForegroundColor Cyan
    Write-Host ("  {0} partner(s) to process" -f $Usage.Count) -ForegroundColor Gray
    Write-Host "  $('-' * 68)" -ForegroundColor DarkGray
    $Script:HWMCounter = 0

    foreach ($customer in $Usage) {
        $Script:HWMCounter++
        $CDPusage =  @()
        $CDPusage += New-Object -TypeName PSObject -Property @{
            PartnerID           = $customer.endcustomer_id ;
            Partnername         = $customer.endcustomer ;
            LegalName           = $customer.endcustomer_Legal ;
            PartnerRef          = $customer.endcustomer_Ref ;
            Site                = $customer.Site ;
            EndCustomer         = $customer.EndCustomer ;
            SO                  = $customer.SO ;
            Reseller            = $customer.Reseller ; 
            SubDisti            = $customer.SubDisti ; 

            PhysicalServerQty   = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.physicalserver -ge 1)} |  measure-object physicalserver -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.physicalserver -ge 1)} |  measure-object physicalserver -sum).sum}else{0} ;
            VirtualServerQty    = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.virtualserver -ge 1)} |  measure-object virtualserver -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.virtualserver -ge 1)} |  measure-object virtualserver -sum).sum}else{0} ;
            WorkstationQty      = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.workstation -ge 1)} |  measure-object workstation -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.workstation -ge 1)} |  measure-object workstation -sum).sum}else{0} ;
            DocumentsQty        = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.documents -ge 1)} |  measure-object documents -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.documents -ge 1)} |  measure-object documents -sum).sum}else{0} ;
            O365UsersQty        = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.o365users -ge 1)} |  measure-object o365users -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.o365users -ge 1)} |  measure-object o365users -sum).sum}else{0} ;
            RecoveryTestingQty  = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.RecoveryTesting -ge 1)}).count))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.recoverytesting -ge 1)}).count}else{0} ;
            CoveGBSelected      = if ($null -ne (($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.Overagegb -ge 1)} |  measure-object Overagegb -sum).sum))
                {($MVplus | Where-Object {($_.endcustomer -eq $customer.endcustomer) -and ($_.Overagegb -ge 1)} |  measure-object Overagegb -sum).sum}else{0} ;
        }

        ## Compute combined server qty (always calculated; used by CWM push when $CombineServers = $true)
        $CDPusage | Add-Member -NotePropertyName CombinedServerQty -NotePropertyValue ([int]$CDPusage.PhysicalServerQty + [int]$CDPusage.VirtualServerQty) -Force

        # ── Per-partner compact output ──
        $hasUsage = (($CDPusage.PhysicalServerQty + $CDPusage.VirtualServerQty +
                      $CDPusage.WorkstationQty    + $CDPusage.DocumentsQty +
                      $CDPusage.O365UsersQty      + $CDPusage.RecoveryTestingQty) -gt 0) -or
                     ($CDPusage.CoveGBSelected -gt 0)

        # Detect misplaced / higher-tier usage (EndCustomer not set)
        $isMisplaced = [string]::IsNullOrWhiteSpace($CDPusage.Partnername)
        $actualLevel = ''; $actualName = ''; $actualId = ''
        if ($isMisplaced) {
            if ($customer.SO)           { $actualLevel = 'Service Organization'; $actualName = $customer.SO;       $actualId = $customer.SO_Id }
            elseif ($customer.Reseller) { $actualLevel = 'Reseller';             $actualName = $customer.Reseller; $actualId = $customer.Reseller_Id }
            elseif ($customer.SubDisti) { $actualLevel = 'Sub-Distributor';      $actualName = $customer.SubDisti; $actualId = '' }
            else                        { $actualLevel = 'Unknown';              $actualName = '(unknown)';        $actualId = '' }
        }

        $displayName  = if ($isMisplaced) { '(EndCustomer not assigned)' } else { $CDPusage.Partnername }
        $displayId    = if ($isMisplaced) { '(none)' }                    else { $CDPusage.PartnerID }
        $nameColor    = if ($isMisplaced) { 'Red' } elseif ($hasUsage) { 'Yellow' } else { 'DarkGray' }

        Write-Host ""
        Write-Host ("  [{0,3}/{1} ]  {2,-44} ID: {3}" -f $Script:HWMCounter, $Usage.Count, $displayName, $displayId) -ForegroundColor $nameColor
        if ($isMisplaced) {
            Write-Host ("             ! Level  : {0,-20}  -->  {1}  (ID: {2})" -f $actualLevel, $actualName, $actualId) -ForegroundColor Red
            Write-Host  "             ! Usage may be internal-use, deployed at wrong tenant level, or at a higher tier than EndCustomer" -ForegroundColor DarkYellow
        } else {
            if ($CDPusage.LegalName -and ($CDPusage.LegalName -ne $CDPusage.Partnername)) {
                Write-Host ("             Legal   : {0}" -f $CDPusage.LegalName) -ForegroundColor Gray
            }
            if ($CDPusage.PartnerRef) {
                Write-Host ("             Cus Ref : {0}" -f $CDPusage.PartnerRef) -ForegroundColor Gray
            }
        }
        if (-not $hasUsage) {
            Write-Host "             --  No billable usage this period" -ForegroundColor DarkGray
        } else {
            $usageColor = if ($isMisplaced) { 'Red' } else { 'Green' }
            Write-Host ("             Phys:{0,-4} Virt:{1,-4}" -f $CDPusage.PhysicalServerQty, $CDPusage.VirtualServerQty) -NoNewline -ForegroundColor $usageColor
            if ($CombineServers) { Write-Host (" Srvr:{0,-4}" -f $CDPusage.CombinedServerQty) -NoNewline -ForegroundColor DarkCyan }
            Write-Host (" Wkstn:{0,-4} Docs:{1,-4} M365:{2,-4} Cont:{3,-4} GB:{4}" -f `
                $CDPusage.WorkstationQty, $CDPusage.DocumentsQty,
                $CDPusage.O365UsersQty,   $CDPusage.RecoveryTestingQty,
                ([math]::Round([decimal]$CDPusage.CoveGBSelected, 2))) -ForegroundColor $usageColor
        }

        $CDPusage | Select-Object   SubDisti,Reseller,SO,EndCustomer,PartnerID,Partnername,LegalName,PartnerRef,
                                    PhysicalServerQty,VirtualServerQty,WorkstationQty,DocumentsQty,RecoveryTestingQty,O365UsersQty,@{n='CoveGBSelected';e={[int]$_.CoveGBSelected}} |
                                    Export-Csv -Path $CDPUsageoutputFile -Append -NoTypeInformation

        $Script:CDPUsageCollection += $CDPusage  ## Store for CWM block to use for company/agreement lookups
    }
    # ── Totals footer ──
    $tPhys  = ($Script:CDPUsageCollection | Measure-Object PhysicalServerQty  -Sum).Sum
    $tVirt  = ($Script:CDPUsageCollection | Measure-Object VirtualServerQty   -Sum).Sum
    $tSrvr  = ($Script:CDPUsageCollection | Measure-Object CombinedServerQty  -Sum).Sum
    $tWkstn = ($Script:CDPUsageCollection | Measure-Object WorkstationQty     -Sum).Sum
    $tDocs  = ($Script:CDPUsageCollection | Measure-Object DocumentsQty       -Sum).Sum
    $tM365  = ($Script:CDPUsageCollection | Measure-Object O365UsersQty       -Sum).Sum
    $tCont  = ($Script:CDPUsageCollection | Measure-Object RecoveryTestingQty -Sum).Sum
    $tGB    = [math]::Round(($Script:CDPUsageCollection | Measure-Object CoveGBSelected -Sum).Sum, 2)
    Write-Host ""
    Write-Host "  $('-' * 68)" -ForegroundColor DarkGray
    Write-Host ("  TOTALS     Phys:{0,-4} Virt:{1,-4}" -f $tPhys, $tVirt) -NoNewline -ForegroundColor White
    if ($CombineServers) { Write-Host (" Srvr:{0,-4}" -f $tSrvr) -NoNewline -ForegroundColor DarkCyan }
    Write-Host (" Wkstn:{0,-4} Docs:{1,-4} M365:{2,-4} Cont:{3,-4} GB:{4}" -f `
        $tWkstn, $tDocs, $tM365, $tCont, $tGB) -ForegroundColor White
    Write-Host "  $('=' * 68)" -ForegroundColor Cyan
    Write-Host ("  CDP Usage Output : {0}" -f $CDPUsageoutputFile) -ForegroundColor Gray
    Write-Host ""

}  ## Execute If GetCDPUsage Switch is $True

#region ----- ConnectWise Manage Body / Functions  ----
if ($ConnectCWM) {
    Function Open-FileName($initialDirectory) {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Exported Cove Usage Files(*.csv)|*CDPUsageFile*.csv"
        $OpenFileDialog.title = "Select a Usage File to Import (*.csv)"
        $result = $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        #$OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    } ## GUI Prompt for Filename to open

    Function InstallCWMPSModule {
        # Install/Update/Load the module
        Write-Output "  Checking | Installing | Updating ConnectWise Manage API PowerShell Module`n  https://www.powershellgallery.com/packages/ConnectWiseManageAPI`n  https://github.com/christaylorcodes/ConnectWiseManageAPI`n"
        Import-module 'PowershellGet'
        if (Get-InstalledModule 'ConnectWiseManageAPI' -ErrorAction SilentlyContinue){
            Update-Module 'ConnectWiseManageAPI'
        } else {
            Install-Module 'ConnectWiseManageAPI' -verbose
        }
        Import-Module 'ConnectWiseManageAPI'
    }  ## Install | Update ConnectWise Manage API PS Module

    
    Function Prompt-CWMAPICreds {
        $CWMAPICreds =$null
        $CWMAPICreds = @{}

        $CWMAPICreds.Server     = Read-Host "Enter the DNS Name of your ConnectWise Manage server e.g. staging.connectwisedev.com"
        $CWMAPICreds.Company    = Read-Host "Enter the Company you use when logging on to ConnectWise Manage"
        # Sensitive fields collected as SecureString and encrypted immediately - plaintext never materialised
        $CWMAPICreds.pubKey     = Read-Host "Enter the Public key created for this integration" -AsSecureString | ConvertFrom-SecureString
        $CWMAPICreds.privateKey = Read-Host "Enter the Private key created for this integration" -AsSecureString | ConvertFrom-SecureString
        $CWMAPICreds.clientId   = Read-Host "Enter your ClientID (You can create/retrieve your ClientID at https://developer.connectwise.com/ClientID)"            -AsSecureString | ConvertFrom-SecureString
    
        $CWMAPICreds | Export-Clixml -Path $CWMAPICredsFile -Force
    }
    
    Function Get-CWMAPICreds {
        if (Test-Path $CWMAPICredsFile) {
            $CWMAPICreds = Import-Clixml -Path $CWMAPICredsFile
            ## Sensitive fields returned as encrypted strings - decryption happens in AuthenticateCWM only
        } else {
            Prompt-CWMAPICreds
            $CWMAPICreds = Get-CWMAPICreds
        }
        return $CWMAPICreds
    }
        
    Function AuthenticateCWM {
        $CWMAPICreds = Get-CWMAPICreds
        if ($null -eq $CWMAPICreds) {
            Prompt-CWMAPICreds
            $CWMAPICreds = Get-CWMAPICreds
        }

        ## Decrypt sensitive fields locally - plaintext lives only for the duration of Connect-CWM call
        ## ZeroFreeBSTR ensures unmanaged BSTR memory is zeroed before release
        $Script:_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($CWMAPICreds.pubKey | ConvertTo-SecureString))
        try   { $_plainPub = [Runtime.InteropServices.Marshal]::PtrToStringAuto($Script:_bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Script:_bstr); $Script:_bstr = $null }

        $Script:_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($CWMAPICreds.privateKey | ConvertTo-SecureString))
        try   { $_plainPriv = [Runtime.InteropServices.Marshal]::PtrToStringAuto($Script:_bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Script:_bstr); $Script:_bstr = $null }

        $Script:_bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($CWMAPICreds.clientId | ConvertTo-SecureString))
        try   { $_plainCid = [Runtime.InteropServices.Marshal]::PtrToStringAuto($Script:_bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Script:_bstr); $Script:_bstr = $null }

        $plainCreds = @{
            Server     = $CWMAPICreds.Server
            Company    = $CWMAPICreds.Company
            pubKey     = $_plainPub
            privateKey = $_plainPriv
            clientId   = $_plainCid
        }
        Connect-CWM @plainCreds
        $plainCreds = $null   ## Clear plaintext hashtable
        $_plainPub = $null; $_plainPriv = $null; $_plainCid = $null  ## Clear individual plaintext variables

        if ($debugCWM) {
            Write-Output "`nDebugging ConnectWise Manage API Credentials:`n"
            Write-Output "Server: $($CWMAPICreds.Server)"
            Write-Output "Company: $($CWMAPICreds.Company)"
            Write-Output "Public Key: ********"
            Write-Output "Private Key: ********"
            Write-Output "Client ID: ********"
        }
    }  ## Connect to ConnectWise Manage instance

    Function GreenText  {
        process { Write-Host $_ -ForegroundColor Green }
    }  ## Output Green text

    Function LookupCWMProducts {
        foreach ($item in $CDP2CWMProductMapping.GetEnumerator()) {
            # https://arcanecode.com/2020/12/14/iterate-over-a-hashtable-in-powershell/
            if ($null -eq (Get-CWMProductCatalog -condition "identifier = `"$($item.Value)`"")) {
                Write-Warning "Product not in CWM catalog: [$($item.Value)] -- update CDP2CWMProductMapping or create product in CWM."
                Add-content -Path $ExceptionsLogFile -Value "Matching product NOT FOUND in ConnectWise Manage for Product ID: [$($item.Value)]."
            }
        }
    }  ## Lookup the Cove Data Protection Products in the ConnectWise Manage Product Catalog

    Function LookupCWMAgreements {
        Write-Output "`n  Checking for ConnectWise Manage Agreements"

        $GetCWMAgreementWithSingleRetry = {
            param(
                [Parameter(Mandatory=$true)][string]$Condition,
                [Parameter(Mandatory=$true)][string]$QueryLabel
            )

            for ($attempt = 1; $attempt -le 2; $attempt++) {
                try {
                    return Get-CWMAgreement -condition $Condition -all -ErrorAction Stop
                } catch {
                    if ($attempt -lt 2) {
                        Write-Warning "CWM agreement query failed for [$QueryLabel] on attempt $attempt of 2. Retrying once. Error: $($_.Exception.Message)"
                    } else {
                        Write-Warning "CWM agreement query failed for [$QueryLabel] after 2 attempts."
                        Add-content -Path $ExceptionsLogFile -Value "CWM agreement query failed after 2 attempts for [$QueryLabel]. Error: $($_.Exception.Message)"
                        return $null
                    }
                }
            }
        }

        if ($CWMAgreementBehavior -eq "Type") {
            ## Type mode: fetch all active agreements, then filter client-side by type name.
            ## (CWM API does not support type/name as a server-side condition field.)
            $allActive = & $GetCWMAgreementWithSingleRetry -Condition "agreementStatus = `"Active`"" -QueryLabel "agreementStatus=Active"
            if ($null -eq $allActive) {
                Break
            }
            $CWMAgreements = $allActive | Where-Object {
                $typeName = $_.type.name
                $CWMAgreementTypes | Where-Object {
                    if     ($CWMAgreementSearchType -eq "Contains")    { $typeName -like "*$_*" }
                    elseif ($CWMAgreementSearchType -eq "StartsWith")  { $typeName -like "$_*" }
                    elseif ($CWMAgreementSearchType -eq "EndsWith")    { $typeName -like "*$_" }
                    else                                               { $typeName -eq $_ }
                }
            }

            if ($null -eq $CWMAgreements) {
                Write-Warning "No active agreements of any matching type [$($CWMAgreementTypes -join ', ')] found in CWM."
                Add-content -Path $ExceptionsLogFile -Value "No matching active agreements were found for agreement types: [$($CWMAgreementTypes -join ', ')]."
                Break
            } else {
                Write-Output "    $($CWMAgreements.count) matching active agreements found for types: [$($CWMAgreementTypes -join ', ')].`n"
                $Script:CWMAgreementTable = $CWMAgreements | Select-Object `
                    @{n='CmpID';e={$_.company.id}},
                    @{n='clientName';e={$_.company.name}},
                    @{n='AgID';e={$_.id}},
                    @{n='agreementName';e={$_.name}},
                    @{n='agreementType';e={$_.type.name}},
                    @{n='Excluded';e={ if ($CWMExcludedCompanies -icontains $_.company.name) { '[EXCLUDED]' } else { '' } }} | Sort-Object clientName
                $Script:CWMAgreementTable | Format-Table -AutoSize
                $Script:CWMAgreementTable | Where-Object { $_.Excluded } | ForEach-Object {
                    Write-Host ("  [!] EXCLUDED: [{0}] {1} [!]" -f $_.CmpID, $_.clientName) -ForegroundColor Red
                }
                $_agSuffix = ($Partnername -replace ' \(.*\)','' -replace '[^a-zA-Z_0-9]','')
                $Script:CWMAgreementsOutputFile = "$CurrentExportPath\$($CurrentDate)_$($_agSuffix)_$($PartnerId)_CWMAgreements.csv"
                $Script:CWMAgreementTable | Export-Csv -Path $Script:CWMAgreementsOutputFile -NoTypeInformation
                Write-Output "    CWM Agreements CSV : $Script:CWMAgreementsOutputFile"
            }
        } else {
            ## Name mode (original behavior)
            if ($CWMAgreementSearchType -eq "Contains") {
                $CWMAgreements = & $GetCWMAgreementWithSingleRetry -Condition "name contains `"$CWMAgreementName`" and agreementStatus = `"Active`"" -QueryLabel "name contains [$CWMAgreementName]"    ## Contains $CWMAgreementName
            } elseif ($CWMAgreementSearchType -eq "StartsWith") {
                $CWMAgreements = & $GetCWMAgreementWithSingleRetry -Condition "name like `"$CWMAgreementName%`" and agreementStatus = `"Active`"" -QueryLabel "name startsWith [$CWMAgreementName]"       ## StartsWith $CWMAgreementName
            } elseif ($CWMAgreementSearchType -eq "EndsWith") {
                $CWMAgreements = & $GetCWMAgreementWithSingleRetry -Condition "name like `"%$CWMAgreementName`" and agreementStatus = `"Active`"" -QueryLabel "name endsWith [$CWMAgreementName]"       ## EndsWith $CWMAgreementName
            } else {
                $CWMAgreements = & $GetCWMAgreementWithSingleRetry -Condition "name = `"$CWMAgreementName`" and agreementStatus = `"Active`"" -QueryLabel "name equals [$CWMAgreementName]"           ## Matches $CWMAgreementName
            }

            if ($null -eq $CWMAgreements) {
                Write-Warning "No active agreements named [$CWMAgreementName] in CWM -- update CWMAgreementName or create agreements in CWM."
                Add-content -Path $ExceptionsLogFile -Value "No matching active agreements were found in your ConnectWise Manage instance for the Agreement Name: [$CWMAgreementName]."
                Break
            } else {
                Write-Output "    $($CWMAgreements.count) matching active agreements were found in ConnectWise Manage for the Agreement Name: [$CWMAgreementName].`n"
                $Script:CWMAgreementTable = $CWMAgreements | Select-Object `
                    @{n='CmpID';e={$_.company.id}},
                    @{n='clientName';e={$_.company.name}},
                    @{n='AgID';e={$_.id}},
                    @{n='agreementName';e={$_.name}},
                    @{n='agreementType';e={$_.type.name}},
                    @{n='Excluded';e={ if ($CWMExcludedCompanies -icontains $_.company.name) { '[EXCLUDED]' } else { '' } }} | Sort-Object clientName
                $Script:CWMAgreementTable | Format-Table -AutoSize
                $Script:CWMAgreementTable | Where-Object { $_.Excluded } | ForEach-Object {
                    Write-Host ("  [!] EXCLUDED: [{0}] {1} [!]" -f $_.CmpID, $_.clientName) -ForegroundColor Red
                }
                $_agSuffix = ($Partnername -replace ' \(.*\)','' -replace '[^a-zA-Z_0-9]','')
                $Script:CWMAgreementsOutputFile = "$CurrentExportPath\$($CurrentDate)_$($_agSuffix)_$($PartnerId)_CWMAgreements.csv"
                $Script:CWMAgreementTable | Export-Csv -Path $Script:CWMAgreementsOutputFile -NoTypeInformation
                Write-Output "    CWM Agreements CSV : $Script:CWMAgreementsOutputFile"
            }
        }
    }  ## Look for ConnectWise Agreements with the specified Agreement Name or Type in ConnectWise Manage

    Function LookupCWMCompany { # Get-CWMcompany -condition "id=28734"    
        $CWMcompany = $null

        Write-Host ""
        Write-Host "  $('-' * 68)" -ForegroundColor DarkGray
        Write-Host ("  Cove: [{0}] {1}" -f $CDPusage.PartnerID, $CDPusage.PartnerName) -ForegroundColor Cyan
        Write-Host ""

        ## Helper: print a detail row for each duplicate company result
        $PrintDuplicateCompanies = {
            param([array]$results, [string]$matchedBy)
            $activeCount = ($results | Where-Object { $_.deletedFlag -ne $true }).Count
            Write-Host ("  [i] Multiple ({0}) CWM companies matched by {1} -- preferring non-deleted, then lowest ID ({2} non-deleted):" -f $results.Count, $matchedBy, $activeCount) -ForegroundColor Cyan
            $results | Sort-Object id | ForEach-Object {
                $statusName  = if ($_.status)    { $_.status.name }    else { '(no status)' }
                $territory   = if ($_.territory) { $_.territory.name } else { '(no territory)' }
                $identifier  = if ($_.identifier){ $_.identifier }     else { '(no identifier)' }
                $deleted     = if ($_.deletedFlag -eq $true) { '  *** DELETED ***' } else { '' }
                Write-Host ("      [{0,6}]  {1,-36}  status: {2,-20}  territory: {3,-20}  id: {4}{5}" -f `
                    $_.id, $_.name, $statusName, $territory, $identifier, $deleted) -ForegroundColor DarkCyan
            }
        }

        $CWMCompanyID = (($CDPusage.PartnerName -split '\| ') -split ' ~')[2]
        if ($CWMCompanyID) {$CWMcompany = Get-CWMcompany -condition "id=$CWMCompanyID" | Select-Object -First 1}

        if ($CWMcompany) {
        ## Match Found for CWMCompanyId
        Write-Host "  [+] CWM client matched (by ID)   : [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
        } else {
            [array]$_result = Get-CWMcompany -condition "name=`"$($CDPusage.PartnerName)`""
            if (($_result | Where-Object { $_.deletedFlag -ne $true }).Count -gt 1) { & $PrintDuplicateCompanies $_result "Name" }
            $CWMcompany = $_result | Where-Object { $_.deletedFlag -ne $true } | Sort-Object id | Select-Object -First 1
            if (-not $CWMcompany -and $_result.Count -gt 0) {
                Write-Host "  [!] All ($($_result.Count)) CWM companies matched by Name are deleted -- skipping" -ForegroundColor Yellow
                Add-content -Path $ExceptionsLogFile -Value "All matching CWM companies for Name: [$($CDPusage.PartnerName)] are deleted -- skipping update."
            }
            if ($CWMcompany) {
                ## Match Found for Partner Name
                Write-Host "  [+] CWM client matched (by Name) : [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
            } else {
                ## NoMatch Found - Try Partner Legal Name
                [array]$_result = Get-CWMcompany -condition "name=`"$($CDPusage.LegalName)`""
                if (($_result | Where-Object { $_.deletedFlag -ne $true }).Count -gt 1) { & $PrintDuplicateCompanies $_result "Legal" }
                $CWMcompany = $_result | Where-Object { $_.deletedFlag -ne $true } | Sort-Object id | Select-Object -First 1
                if (-not $CWMcompany -and $_result.Count -gt 0) {
                    Write-Host "  [!] All ($($_result.Count)) CWM companies matched by Legal are deleted -- skipping" -ForegroundColor Yellow
                    Add-content -Path $ExceptionsLogFile -Value "All matching CWM companies for LegalName: [$($CDPusage.LegalName)] are deleted -- skipping update."
                }
                if ($CWMcompany) {
                    ## Match Found for Partner Legal Name
                    Write-Host "  [+] CWM client matched (by Legal): [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                } else {
                    ## NoMatch Found - Try Partner Reference Name
                    [array]$_result = Get-CWMcompany -condition "name=`"$($CDPusage.PartnerRef)`""
                    if (($_result | Where-Object { $_.deletedFlag -ne $true }).Count -gt 1) { & $PrintDuplicateCompanies $_result "Ref" }
                    $CWMcompany = $_result | Where-Object { $_.deletedFlag -ne $true } | Sort-Object id | Select-Object -First 1
                    if (-not $CWMcompany -and $_result.Count -gt 0) {
                        Write-Host "  [!] All ($($_result.Count)) CWM companies matched by Ref are deleted -- skipping" -ForegroundColor Yellow
                        Add-content -Path $ExceptionsLogFile -Value "All matching CWM companies for Ref: [$($CDPusage.PartnerRef)] are deleted -- skipping update."
                    }

                    if ($CWMcompany) {
                        ## Match Found for Partner Reference Name
                        Write-Host "  [+] CWM client matched (by Ref)  : [$($CWMcompany.ID)] $($CWMcompany.Name)" -ForegroundColor Green
                    } else {
                        ## NoMatches Found - Output Warning
                        Write-Host "  [!] No CWM client matched -- ID:[$($CDPusage.PartnerID)] Name:[$($CDPusage.PartnerName)] Legal:[$($CDPusage.LegalName)] Ref:[$($CDPusage.PartnerRef)]" -ForegroundColor Yellow
                        Add-content -Path $ExceptionsLogFile -Value "Matching client NOT FOUND in ConnectWise Manage for Cove Data Protection CustomerID: [$($CDPusage.PartnerID)], PartnerName: [$($CDPusage.PartnerName)], LegalName: [$($CDPusage.LegalName)], Reference: [$($CDPusage.PartnerRef)]"
                    }
                }
            }
        }

        if ($CWMcompany) {
            ## Check if this company is in the exclusion list
            if ($CWMExcludedCompanies -icontains $CWMcompany.name) {
                Write-Host ("  [!] EXCLUDED: [{0}] {1} -- in excluded companies list, skipping CWM update" -f $CWMcompany.id, $CWMcompany.name) -ForegroundColor Red
                Add-content -Path $ExceptionsLogFile -Value "Company [$($CWMcompany.Name)] (ID: $($CWMcompany.id)) is in the excluded companies list -- CWM additions NOT updated."
                return
            }
            if ($CWMAgreementBehavior -eq "Type") {
                ## Type mode: fetch ALL active agreements for the matched company (no server-side type filter —
                ## CWM API does not support type/name as a condition field). Filter + display client-side.
                [array]$allActiveAgreements = Get-CWMAgreement -Condition "company/id=$([int]$CWMcompany.id) AND agreementStatus = `"Active`"" -all

                [array]$CWMAgreement = $allActiveAgreements | Where-Object {
                    $typeName = $_.type.name
                    $CWMAgreementTypes | Where-Object {
                        if     ($CWMAgreementSearchType -eq "Contains")    { $typeName -like "*$_*" }
                        elseif ($CWMAgreementSearchType -eq "StartsWith")  { $typeName -like "$_*" }
                        elseif ($CWMAgreementSearchType -eq "EndsWith")    { $typeName -like "*$_" }
                        else                                               { $typeName -eq $_ }
                    }
                }

                if ($null -eq $CWMAgreement -or $CWMAgreement.Count -eq 0) {
                    ## 0 matches -- show all active agreements as diagnostic
                    Write-Host "  [i] Active agreements for [$($CWMcompany.Name)] -- no type match:" -ForegroundColor Cyan
                    if ($allActiveAgreements) {
                        $allActiveAgreements | ForEach-Object {
                            Write-Host ("      [{0,6}]  {1,-40}  type: {2}" -f $_.id, $_.name, $_.type.name) -ForegroundColor DarkCyan
                        }
                    } else {
                        Write-Host "      (no active agreements at all)" -ForegroundColor DarkGray
                    }
                    Write-Host "  [!] No active agreement of type [$($CWMAgreementTypes -join ' | ')] for: [$($CWMcompany.Name)]" -ForegroundColor Yellow
                    Add-content -Path $ExceptionsLogFile -Value "No active agreement matching any of the agreement types [$($CWMAgreementTypes -join ', ')] was found for ConnectWise Client Name: [$($CWMcompany.Name)]"
                } elseif ($CWMAgreement.Count -gt 1) {
                    ## >1 matches -- show only the matched agreements as diagnostic
                    Write-Host "  [i] Multiple type-matching agreements for [$($CWMcompany.Name)]:" -ForegroundColor Cyan
                    $CWMAgreement | ForEach-Object {
                        Write-Host ("      [{0,6}]  {1,-40}  type: {2}" -f $_.id, $_.name, $_.type.name) -ForegroundColor DarkCyan
                    }
                    Write-Host "  [!] Multiple ($($CWMAgreement.Count)) matching agreements found for: [$($CWMcompany.Name)] -- skipping update" -ForegroundColor Yellow
                    Add-content -Path $ExceptionsLogFile -Value "Multiple ($($CWMAgreement.Count)) active agreements of matching types [$($CWMAgreementTypes -join ', ')] found for ConnectWise Client Name: [$($CWMcompany.Name)]. Cannot determine which to update -- skipping."
                } else {
                    ## exactly 1 match -- green only, no [i] block
                    $CWMAgreement = $CWMAgreement[0]
                    Write-Host "  [+] Agreement [$($CWMAgreement.name)] (type: $($CWMAgreement.type.name)) found for: [$($CWMcompany.Name)]" -ForegroundColor Green
                    if ($ShouldProcessCWMAdditions) {
                        UpdateCWMQty    ## Update the Additions for the matching ConnectWise Manage Agreement with the quantities from the Cove Data Protection Usage
                    }
                }
            } else {
                ## Name mode (original behavior): look for active agreement matching CWMAgreementName
                if ($CWMAgreementSearchType -eq "Contains") {
                    $CWMAgreement = Get-CWMAgreement -Condition "company/id=$($CWMcompany.id) AND name contains `"$CWMAgreementName`" and agreementStatus = `"Active`""    ## Contains $CWMAgreementName
                } elseif ($CWMAgreementSearchType -eq "StartsWith") {
                    $CWMAgreement = Get-CWMAgreement -Condition "company/id=$($CWMcompany.id) AND name like `"$CWMAgreementName%`" and agreementStatus = `"Active`""       ## StartsWith $CWMAgreementName
                } elseif ($CWMAgreementSearchType -eq "EndsWith") {
                    $CWMAgreement = Get-CWMAgreement -Condition "company/id=$($CWMcompany.id) AND name like `"%$CWMAgreementName`" and agreementStatus = `"Active`""       ## EndsWith $CWMAgreementName
                } else {
                    $CWMAgreement = Get-CWMAgreement -Condition "company/id=$($CWMcompany.id) AND name = `"$CWMAgreementName`" and agreementStatus = `"Active`""           ## Matches $CWMAgreementName
                }
                If ($null -eq $CWMAgreement) {
                    Write-Host "  [!] No active agreement [$CWMAgreementName] for: [$($CWMcompany.Name)]" -ForegroundColor Yellow
                    Add-content -Path $ExceptionsLogFile -Value "Matching active agreement NOT FOUND in ConnectWise Manage for Agreement Name: [$CWMAgreementName] for ConnectWise Client Name: [$($CWMcompany.Name)]"
                } else {
                    Write-Host "  [+] Agreement [$CWMAgreementName] found for: [$($CWMcompany.Name)]" -ForegroundColor Green
                    if ($ShouldProcessCWMAdditions) {
                        UpdateCWMQty    ## Update the Additions for the matching ConnectWise Manage Agreement with the quantities from the Cove Data Protection Usage
                    }
                }
            }
        }
    }  ## Lookup Cove Data Protection Customer in ConnectWise Manage

    Function UpdateCWMQty {
        $ReadOnlyCWMUsageMode = ($NoCWMUpdate -and $ShouldProcessCWMAdditions)

        try {
            $CWMAdditions = Get-CWMAgreementAddition -AgreementID $CWMAgreement.id -Condition "agreementStatus = `"Active`"" -all -ErrorAction Stop
        } catch {
            Write-Host "  [!] Failed to retrieve CWM agreement additions for [$($CWMcompany.Name)] / [$($CWMAgreement.name)] -- continuing" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Failed to retrieve CWM agreement additions for ConnectWise Client Name: [$($CWMcompany.Name)] Agreement: [$($CWMAgreement.name)]. Error: $($_.Exception.Message)"
            return
        }

        $PhysicalServerMatched = $false
        $VirtualServerMatched = $false
        $CombinedServerMatched = $false
        $WorkstationMatched = $false
        $DocumentsMatched = $false
        $RecoveryTestingMatched = $false
        $O365UsersMatched = $false
        $CoveGBSelectedMatched = $false

        if ($debugCWM) {
            # Store before quantities keyed by product identifier for side-by-side comparison after update
            $beforeQty = @{}
            foreach ($add in $CWMAdditions) { $beforeQty[$add.product.identifier] = [decimal]$add.quantity }
            $maxIdLen = [Math]::Max(28, ($CWMAdditions | ForEach-Object { ("[C] " + $_.product.identifier).Length } | Measure-Object -Maximum).Maximum)
            Write-Host ""
            Write-Host ("  {0}  |  {1}  |  {2}" -f $CDPusage.partnername, $CWMcompany.name, $CWMAgreement.name) -ForegroundColor Cyan
            if ($ReadOnlyCWMUsageMode) {
                Write-Host "  Current addition quantities (read-only):" -ForegroundColor White
            } else {
                Write-Host "  Addition quantities before update:" -ForegroundColor White
            }
            Write-Host ("  {0,-$maxIdLen}  {1,8}" -f "Product", "Qty") -ForegroundColor White
            Write-Host ("  {0}  {1}" -f ("-" * $maxIdLen), "--------") -ForegroundColor DarkGray
            $sortedAdditions = $CWMAdditions | Sort-Object { $_.product.identifier }
            foreach ($add in $sortedAdditions) {
                $isCove = $add.product.identifier -like "Cove*"
                $productLabel = if ($isCove) { "[C] $($add.product.identifier)" } else { "    $($add.product.identifier)" }
                $displayColor = if ($isCove) { 'DarkCyan' } else { 'Gray' }
                Write-Host ("  {0,-$maxIdLen}  {1,8:N2}" -f $productLabel, ([decimal]$add.quantity)) -ForegroundColor $displayColor
            }
            Write-Host ""
        }

        if ($ReadOnlyCWMUsageMode) {
            foreach ($CWMAddition in $CWMAdditions) {
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),"""""
            }
            Write-Host "  [NoCWMUpdate] Captured current CWM addition quantities (all additions); no updates were applied." -ForegroundColor DarkCyan
            return
        }

        foreach ($CWMAddition in $CWMAdditions) {
            $Update = @{
                AgreementID = $CWMAgreement.id
                AdditionID = $CWMAddition.id
                Operation = 'replace'
                Path = 'quantity'
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.PhysicalServerQty))) {
                if ($CombineServers) {
                    ## CombineServers mode: zero out the split Physical addition
                    Write-Host "  [+] Matched: $($CDP2CWMProductMapping.PhysicalServerQty) (zeroed -- CombineServers active)" -ForegroundColor DarkCyan
                    Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),0"
                    Update-CWMAgreementAddition @Update -Value 0 | out-null
                } else {
                    Write-Host "  [+] Matched: $($CDP2CWMProductMapping.PhysicalServerQty)" -ForegroundColor Green
                    Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.PhysicalServerQty)"
                    Update-CWMAgreementAddition @Update -Value $CDPusage.PhysicalServerQty | out-null
                }
                $PhysicalServerMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.VirtualServerQty))) {
                if ($CombineServers) {
                    ## CombineServers mode: zero out the split Virtual addition
                    Write-Host "  [+] Matched: $($CDP2CWMProductMapping.VirtualServerQty) (zeroed -- CombineServers active)" -ForegroundColor DarkCyan
                    Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),0"
                    Update-CWMAgreementAddition @Update -Value 0 | out-null
                } else {
                    Write-Host "  [+] Matched: $($CDP2CWMProductMapping.VirtualServerQty)" -ForegroundColor Green
                    Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.VirtualServerQty)"
                    Update-CWMAgreementAddition @Update -Value $CDPusage.VirtualServerQty | out-null
                }
                $VirtualServerMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.CombinedServerQty))) {
                if ($CombineServers) {
                    ## CombineServers mode: push Phys+Virt sum to the combined addition
                    Write-Host "  [+] Matched: $($CDP2CWMProductMapping.CombinedServerQty) (combined Phys+Virt = $($CDPusage.CombinedServerQty))" -ForegroundColor Green
                    Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.CombinedServerQty)"
                    Update-CWMAgreementAddition @Update -Value $CDPusage.CombinedServerQty | out-null
                } else {
                    Write-Host "  [i] Skipped: $($CDP2CWMProductMapping.CombinedServerQty) (CombineServers not active)" -ForegroundColor DarkGray
                }
                $CombinedServerMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.WorkstationQty))) {
                Write-Host "  [+] Matched: $($CDP2CWMProductMapping.WorkstationQty)" -ForegroundColor Green
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.WorkstationQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.WorkstationQty | out-null
                $WorkstationMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.DocumentsQty))) {
                Write-Host "  [+] Matched: $($CDP2CWMProductMapping.DocumentsQty)" -ForegroundColor Green
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.DocumentsQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.DocumentsQty | out-null
                $DocumentsMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.RecoveryTestingQty))) {
                Write-Host "  [+] Matched: $($CDP2CWMProductMapping.RecoveryTestingQty)" -ForegroundColor Green
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.RecoveryTestingQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.RecoveryTestingQty | out-null
                $RecoveryTestingMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.O365UsersQty))) {
                Write-Host "  [+] Matched: $($CDP2CWMProductMapping.O365UsersQty)" -ForegroundColor Green
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.O365UsersQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.O365UsersQty | out-null
                $O365UsersMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.CoveGBSelected))) {
                Write-Host "  [+] Matched: $($CDP2CWMProductMapping.CoveGBSelected)" -ForegroundColor Green
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAgreement.type.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.CoveGBSelected)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.CoveGBSelected | out-null
                $CoveGBSelectedMatched = $true
            }
        }

        ## Show list of products Report an exception if Usage Qty > 0 and no matching Addition was found
        if ($debugCWM) {
            start-sleep -Milliseconds 2500    ## Wait before retrieving the updated Additions
            $CWMAdditions = Get-CWMAgreementAddition -AgreementID $CWMAgreement.id -Condition "agreementStatus = `"Active`"" -all
            $maxIdLen = [Math]::Max(28, ($CWMAdditions | ForEach-Object { ("[C] " + $_.product.identifier).Length } | Measure-Object -Maximum).Maximum)
            Write-Host ""
            Write-Host ("  {0}  |  {1}  |  {2}" -f $CDPusage.partnername, $CWMcompany.name, $CWMAgreement.name) -ForegroundColor Cyan
            Write-Host "  Addition quantities after update:" -ForegroundColor White
            Write-Host ("  {0,-$maxIdLen}  {1,8}  {2,8}  {3,8}" -f "Product", "Before", "After", "Change") -ForegroundColor White
            Write-Host ("  {0}  {1}  {2}  {3}" -f ("-" * $maxIdLen), "--------", "--------", "--------") -ForegroundColor DarkGray
            $sortedAdditions = $CWMAdditions | Sort-Object { $_.product.identifier }
            foreach ($add in $sortedAdditions) {
                $id     = $add.product.identifier
                $after  = [decimal]$add.quantity
                $before = if ($null -ne $beforeQty -and $beforeQty.ContainsKey($id)) { $beforeQty[$id] } else { [decimal]0 }
                $diff   = $after - $before
                $isCove = $id -like "Cove*"
                $productLabel = if ($isCove) { "[C] $id" } else { "    $id" }

                if ($diff -gt 0) {
                    $changeStr = "+{0:N2}" -f $diff
                    $rowColor  = 'Green'
                } elseif ($diff -lt 0) {
                    $changeStr = "{0:N2}" -f $diff
                    $rowColor  = 'Red'
                } else {
                    $changeStr = "+0.00"
                    $rowColor  = 'Yellow'
                }
                Write-Host ("  {0,-$maxIdLen}  {1,8:N2}  {2,8:N2}  {3,8}" -f $productLabel, $before, $after, $changeStr) -ForegroundColor $rowColor
            }
            Write-Host ""
        }  ## end if ($debugCWM)

        ## In CombineServers mode, warn on missing split additions only if their qtys are non-zero (they will be zeroed so warning is still useful).
        ## Always warn on missing CombinedServerQty addition when CombineServers is active and total servers > 0.
        if (-not $CombineServers -and $CDPusage.PhysicalServerQty -gt 0 -and $PhysicalServerMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.PhysicalServerQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.PhysicalServerQty)] in the ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if (-not $CombineServers -and $CDPusage.VirtualServerQty -gt 0 -and $VirtualServerMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.VirtualServerQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.VirtualServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CombineServers -and $CDPusage.CombinedServerQty -gt 0 -and $CombinedServerMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.CombinedServerQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.CombinedServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.WorkstationQty -gt 0 -and $WorkstationMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.WorkstationQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$CDP2CWMProductMapping.WorkstationQty] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.DocumentsQty -gt 0 -and $DocumentsMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.DocumentsQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.DocumentsQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.RecoveryTestingQty -gt 0 -and $RecoveryTestingMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.RecoveryTestingQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.RecoveryTestingQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.O365UsersQty -gt 0 -and $O365UsersMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.O365UsersQty)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.O365UsersQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.CoveGBSelected -gt 0 -and $CoveGBSelectedMatched -eq $false) {
            Write-Host "  [!] No addition [$($CDP2CWMProductMapping.CoveGBSelected)] in [$($CWMAgreement.name)] for [$($CWMcompany.Name)]" -ForegroundColor Yellow
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.CoveGBSelected)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }


    }  ## Update the Additions for the matching ConnectWise Manage Agreement with the quantities from the Cove Data Protection Usage

    InstallCWMPSModule         ## Install the ConnectWise Manage PowerShell module

    $Script:CWMAPICredsFile = join-path -Path $True_Path -ChildPath "$env:computername CWM_API_Credentials.Secure.metrotech.xml"
    $Script:CWMAPICredsPath = Split-path -path $CWMAPICredsFile


    if ($ClearCWMCredentials) {
        if (Test-Path $CWMAPICredsFile) {
            Remove-Item -Path $CWMAPICredsFile -Force
            Write-Output "ConnectWise Manage API Credentials file removed successfully."
        } else {
            Write-Warning "ConnectWise Manage API Credentials file not found at $CWMAPICredsFile."
        }
    }
    
    AuthenticateCWM            ## Authenticate with ConnectWise Manage
    LookupCWMProducts          ## Lookup the Data Protection Products in the ConnectWise Manage Product Catalog
    if ($debugCWM) {
        LookupCWMAgreements    ## Look for Agreements with the specified Agreement Name in ConnectWise Manage
    }

    If ($GetCDPUsage -and $Script:CDPUsageCollection) {
        ## Per-customer usage was already built and exported to CSV in the GetCDPUsage block above.
        ## This block handles only the CWM company/agreement lookup and optional quantity update.
        foreach ($CDPusage in $Script:CDPUsageCollection) {
            LookupCWMCompany    ## Lookup Cove Data Protection Customer in ConnectWise Manage
        }
    }
    if ($LoadCDPUsage) {
        if ($IsInteractiveMode) {
            # Interactive mode: prompt user to select import file.
            $OpenFileName = Open-FileName "$ExportPath"
            if ($null -eq $OpenFileName) { Break }
        } else {
            # Unattended mode: auto-select the newest CDP usage import file under ExportPath.
            $LatestUsageFile = Get-ChildItem -Path $ExportPath -Filter "*CDPUsageFile*.csv" -File -Recurse -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($null -eq $LatestUsageFile) {
                Write-Warning "No usage import file matching '*CDPUsageFile*.csv' found under ExportPath: $ExportPath"
                Break
            }

            $OpenFileName = $LatestUsageFile.FullName
            Write-Output "Auto-selected latest usage import file: $OpenFileName"
        }

        #Import CSV file to $cdpusage
        $Usage = Import-Csv -Path $OpenFileName
        #Validate CSV file
        $expectedHeaders =  "SubDisti","Reseller","SO","EndCustomer",
                            "PartnerID","PartnerName","LegalName","PartnerRef",
                            "PhysicalServerQty","VirtualServerQty",
                            "WorkstationQty","DocumentsQty",
                            "RecoveryTestingQty","O365UsersQty",
                            "CoveGBSelected"

        $actualHeaders = $Usage[0].PSObject.Properties.Name
        $missingHeaders = $expectedHeaders | Where-Object { $_ -notin $actualHeaders }
        
        if ($missingHeaders.Count -gt 0) {
            Write-Warning "Aborting script, missing headers in the CSV File: $($missingHeaders -join ', ')`n"
            Break
            # You can take additional actions here, such as logging or stopping the process.
        } else {
            Write-Host "`nAll expected headers are present!"
        }

        ## Inject CombinedServerQty before GridView so the column is visible at selection time
        if ($CombineServers) {
            $Usage | ForEach-Object {
                $_ | Add-Member -NotePropertyName CombinedServerQty -NotePropertyValue ([int]$_.PhysicalServerQty + [int]$_.VirtualServerQty) -Force
            }
        }

        $SelectedUsage = $Usage | Out-GridView -Title "Select End-Customer Usage to Process from $($openfilename)" -OutputMode Multiple
        
        if ($null -eq $SelectedUsage) {
            Write-Warning "No End-Customers Selected to Process, Exiting Script`n"
            break
        }else{
            Write-Output "`nProcessing $($SelectedUsage.count) Selected End-Customers"
        }
        
        #for each selected customer with usage LookupCWMCompany

        Foreach ($CDPusage in $SelectedUsage) {
            Write-Output "`nCove Data Protection Usage Import from File for Month ending $end"
            Write-Output "    PartnerId         : $($CDPusage.PartnerID)"
            Write-Output "    PartnerName       : $($CDPusage.Partnername)"
            Write-Output "    LegalName         : $($CDPusage.LegalName)"
            Write-Output "    PartnerRef        : $($CDPusage.PartnerRef)"
            Write-Output "    PhysicalServerQty : $($CDPusage.PhysicalServerQty)"
            Write-Output "    VirtualServerQty  : $($CDPusage.VirtualServerQty)"
            Write-Output "    WorkstationQty    : $($CDPusage.WorkstationQty)"
            Write-Output "    DocumentsQty      : $($CDPusage.DocumentsQty)"
            Write-Output "    M365UsersQty      : $($CDPusage.O365UsersQty)"
            Write-Output "    ContinuityQty     : $($CDPusage.RecoveryTestingQty)"
            Write-Output "    GBSelectedQty     : $($CDPusage.CoveGBSelected)`n"

            ## Recompute CombinedServerQty from CSV columns (not stored in CSV export)
            $CDPusage | Add-Member -NotePropertyName CombinedServerQty -NotePropertyValue ([int]$CDPusage.PhysicalServerQty + [int]$CDPusage.VirtualServerQty) -Force

            LookupCWMCompany    ## Lookup Cove Data Protection Customer in ConnectWise Manage
        }
    }
} else {
    Write-Warning "ConnectWise Manage can not be reached as the ConnectCWM switch is disabled."
}  ## Execute if ConnectCWM switch is enabled
    
#endregion ----- ConnectWise Manage Body / Functions  ----
Write-Output $Script:strLineSeparator
Write-Output "Log files written to: $CurrentExportPath"
Write-Output "  Audit Log File: $AuditLogFile"
Write-Output "  Exceptions Log File: $ExceptionsLogFile"
Write-Output "  MaxValue CSV Output File: $CSVOutputFile"
Write-Output "  CDP Usage Output File: $CDPUsageoutputFile"
if ($Script:CWMAgreementsOutputFile) { Write-Output "  CWM Agreements CSV: $Script:CWMAgreementsOutputFile" }
Write-Output $Script:strLineSeparator
Start-Sleep -seconds 5
