# ----- About: ----
    # N-able Cove Data Protection to ConnectWise Manage PSA Integration using the Cove Data Protection MaxValue Plus Usage Report
    # Revision v24.10.23
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
    #
    # Use the -Period switch parameter to define Usage Dates (yyyy-MM or MM-yyyy)
    # Use the -DeviceCount ## (default=15000) parameter to define the maximum number of devices returned
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherlands)
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path. Default is the execution path of the script.
    # Use the -DebugCDP switch to display debug info for Cove Data Protection usage and statistics lookup
    # Use the -DebugCWM switch to display debug info for ConnectWise Manage data lookup
    # Use the -GetCDPUsage switch to pull Cove HWM usage data from product
    # Use the -ImportCDPUsage switch to load Cove HWM usage data from a CSV file
    # Use the -SendToCWM switch to push Cove usage data ConnectWise Manage
    # Use the -CWMAgreementName parameter to specifiy the ConnectWise Manage Agreement Name to update with the Cove Data Protection usage
    # Use the -CWMAgreementSearchType switch to specifiy how to match the ConnectWise Manage Agreement Name
    # Use the -ClearCredentials parameter to remove stored Cove API credentials at start of script
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
        [Parameter(Mandatory=$False)][datetime]$Period,                                                 ## Lookup Date yyyy-MM or MM-yyyy
        [Parameter(Mandatory=$False)][ValidateRange(0,24)]$Last = 1,                                    ## Count back # Last Months i.e. 0 current, 1 Prior Month
        [Parameter(Mandatory=$False)][int]$DeviceCount = 10000,                                         ## Change Maximum Number of current devices results to return
        [Parameter(Mandatory=$False)][Switch]$GetCDPUsage = $true,                                     ## Pull Cove High Water Mark Usage fron Product
        [Parameter(Mandatory=$False)][Switch]$ImportCDPUsage = $false,                                  ## Import Cove Usage Data from External CSV File     
        [Parameter(Mandatory=$False)][Switch]$DebugCDP = $true,                                         ## Enable Debug for Cove Max Value / Device Statistics
        [Parameter(Mandatory=$False)][Switch]$TrialState = $false,                                      ## Treat Trial partners as Billable when set $true
        [Parameter(Mandatory=$False)][switch]$Launch = $true,                                           ## Launch XLS or CSV file
        [Parameter(Mandatory=$False)][string]$Delimiter = ',',                                          ## Specify ',' or ';' Delimiter for XLS & CSV file
        [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",                             ## Export Path
        [Parameter(Mandatory=$False)][int]$PhysServerFUGB = 2000,                                       ## Selected Size Fair Use, 0 to pass full value
        [Parameter(Mandatory=$False)][int]$VirtServerFUGB = 2000,                                       ## Selected Size Fair Use, 0 to pass full value
        [Parameter(Mandatory=$False)][int]$WorkstationFUGB = 500,                                       ## Selected Size Fair Use, 0 to pass full value
        [Parameter(Mandatory=$False)][Switch]$ConnectCWM = $true,                                       ## Test Connection to ConnectWise Manage, must be $true to generate an importable CSV file
        [Parameter(Mandatory=$False)][Switch]$UpdateCWM = $true,                                        ## Send Usage Data to ConnectWise Manage
        [Parameter(Mandatory=$False)][Switch]$DebugCWM = $true,                                         ## Enable Debug for ConnectWise Manage
        [Parameter(Mandatory=$False)][String]$CWMAgreementName = "CoveDataProtection",                  ## ConnectWise Manage Agreement Name (Preferred = 'CoveDataProtection' )
        [Parameter(Mandatory=$False)][String]$CWMAgreementSearchType = "Equals",                        ## ConnectWise Manage Agreement Search Type. Can be either Equals, Contains, StartsWith, EndsWith
        [Parameter(Mandatory=$False)][String]$CWMAdditionSearchType = "Equals",                         ## ConnectWise Manage Addition Search Type. Can be either Equals or StartsWith
        [Parameter(Mandatory=$False)][switch]$ClearCredentials                                          ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----

$CWMAPICreds = @{
    Server      = 'Endpoint.connectwise.com'                ## If using the cloud version of ConnectWise Manage then this will be the DNS Name of server for your region, otherwise the DNS Name for your ConnectWise Manage server.
    Company     = 'samplecompany'                           ## The Company you enter when logging on to ConnectWise Manage
    pubKey      = 'osEdehti-sample'                         ## The Public key created for this integration
    privateKey  = '5OQYikD-sample'                          ## Private key created for this integration
    clientId    = '81696495-xxxx-xxxx-xxxx-5df01b261886'    ## You can create/retrieve your ClientID at https://developer.connectwise.com/ClientID. If you are not a registed ConnectWise developer then you'll need to register first at https://register.developer.connectwise.com/.
}  ## ConnectWise Manage API Credentials

$CDP2CWMProductMapping = @{

    ## Cove Data Protection Product ID >>> ConnectWise Manage Product ID (Products must be manually created in ConnectWise Manage)

    "PhysicalServerQty"     = "CDP-Physical-Servers"    ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Physical-Servers', Product Description = 'Cove Data Protection - Physical Server Backup')
    "VirtualServerQty"      = "CDP-Virtual-Servers"     ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Virtual-Servers', Product Description = 'Cove Data Protection - Virtual Server Backup')
    "WorkstationQty"        = "CDP-Workstations"        ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Workstations', Product Description = 'Cove Data Protection - Workstation Professional Backup')
    "DocumentsQty"          = "CDP-Documents"           ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Documents', Product Description = 'Cove Data Protection - Workstation Documents Backup')
    "O365UsersQty"          = "CDP-M365-Users"          ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-M365-Users', Product Description = 'Cove Data Protection - M365 User Backup')
    "RecoveryTestingQty"    = "CDP-Continuity"          ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-Continuity', Product Description = 'Cove Data Protection - Continuity License')      
    "CoveGBSelected"        = "CDP-GBSelected-Size"     ## ConnectWise Manage Product ID (Recommended Product ID = 'CDP-GBSelected-Size', Product Description = 'Cove Data Protection - GB Selected Size')
        
    ## 
    ## If you need to have different prices for a product (E.g. Not for Profits, Education, Government) then create products which use above Product ID's as a prefix
    ## followed by a unique suffix. (E.g. CDP-Physical-Servers-NFP, CDP-Physical-Servers-Edu, CDP-Physical-Servers-Gov).
    ## 
    ## Note: Due to the way the script works you can't have multiple active Additions with the same Product ID or Product ID prefix.
    ## 

}  ## Cove to PSA Product Mapping

Clear-Host
#Requires -RunAsAdministrator
$ConsoleTitle = "Cove Data Protection to ConnectWise Manage PSA Integration"
$host.UI.RawUI.WindowTitle = $ConsoleTitle      ## Comment out for Automation Policy Use
Write-Output "$ConsoleTitle`n`n$($myInvocation.MyCommand.Path)"
$Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
if ($myInvocation.MyCommand.Path) { Split-path $MyInvocation.MyCommand.Path | Push-Location } ## Set terminal path to match script location (does not support editor 'Run Selection')
if ($exportpath -eq "") {$exportpath = Split-Path $MyInvocation.MyCommand.Path}

Write-Output "  Current Parameters:"
Write-Output "    - Period                  = $Period"
Write-Output "    - Months Prior            = $Last"
Write-Output "    - DeviceCount             = $DeviceCount"
Write-Output "    - GetCDPusage             = $GetCDPusage"
Write-Output "    - ImportCDPusage          = $ImportCDPusage"
Write-Output "    - TrialState              = $TrialState"
Write-Output "    - DebugCDP                = $DebugCDP"
Write-Output "    - Launch                  = $Launch"
Write-Output "    - Delimiter               = $Delimiter"
Write-Output "    - ExportPath              = $ExportPath"
Write-Output "    - PhysServerFUGB          = $PhysServerFUGB GB"
Write-Output "    - VirtServerFUGB          = $VirtServerFUGB GB"
Write-Output "    - WorkstationFUGB         = $WorkstationFUGB GB"
Write-Output "    - ConnectCWM              = $ConnectCWM"
Write-Output "    - UpdateCWM               = $UpdateCWM"
Write-Output "    - DebugCWM                = $DebugCWM"
Write-Output "    - CWMAgreementName        = $CWMAgreementName"
Write-Output "    - CWMAgreementSearchType  = $CWMAgreementSearchType"
Write-Output "    - CWMAdditionSearchType   = $CWMAdditionSearchType"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------------"
$CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm"
$urljson = "https://api.backup.management/jsonapi"

if ($ImportCDPUsage -and $GetCDPUsage) {Write-warning " Script Switch parameters 'GetCDPUsage' and 'ImportCDPUsage' can not be used together, exiting script.";break}
if (-not $ImportCDPUsage -and -not $GetCDPUsage) {
    do {
        $choice = Read-Host "`nParameters 'ImportCDPUsage' & 'GetCDPUsage' are not set. `nDo you want to (G)et usage from product, (I)mport usage from CSV, or e(X)it? Enter 'G', 'I', or 'X'"
        switch ($choice.ToUpper()) {
            'I' { $ImportCDPUsage = $true }
            'G' { $GetCDPUsage = $true }
            'X' { Write-Output "Exiting script."; exit }
            default { Write-Output "Invalid choice, please enter 'G', 'I', or 'X'." }
        }
    } until ($choice.ToUpper() -eq 'G' -or $choice.ToUpper() -eq 'I' -or $choice.ToUpper() -eq 'X')
}

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
        New-Item -ItemType Directory -Path $APIcredpath
    }

    Write-Output "  Enter Exact, Case Sensitive Partner Name for Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for Backup.Management API'
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
    } else {
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
            Write-Output "  Stored Backup API Partner  = $Script:cred0"
            Write-Output "  Stored Backup API User     = $Script:cred1"
            Write-Output "  Stored Backup API Password = Encrypted"

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
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url `
        -TimeoutSec 30 `
        -SessionVariable Script:websession `
        -UseBasicParsing
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
        Send-GetPartnerInfo $Script:partnername
    }

    if ($partner.error) {
        Write-Output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
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
            Write-Output "Parent | $Historyparentid | $Historylevel | $HistorypartnerId | $HistoryPartnerName"
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
            $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single

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

Function Send-GetDevices {

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
    $data.params.query.Filter = $Filter1
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

    ForEach ( $DeviceResult in $DeviceResponse.result.result ) {
        if ($debugCDP) {Write-Output "Getting statistics data for deviceid | $($DeviceResult.AccountId)"}
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
        if ($debugCDP) {Write-Output "Getting historic data for deleted device | $($DeviceHistory.AccountId)"}
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
    Write-Output "Requesting Maximum Value Report from $start to $end"

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
        Write-Host "Module ImportExcel Already Installed"
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

    $Script:MVReport | foreach-object {
        if ($_.customerid) {
            if ($null -eq $Script:EnumerateAncestorPartners.result.result.id) {
                send-EnumerateAncestorPartners $_.customerid
                write-Output "Getting ancestor data for initial partnerid | $($_.CustomerId) | #$loop "
                
                if ($Script:EnumerateAncestorPartners.result.result.id.count -gt 0) {
                    Write-Output "Found $($Script:EnumerateAncestorPartners.result.result.id[0])"
                }else{
                    #Write-Output $Script:EnumerateAncestorPartners.error.message
                }
            } ## If Initial Ancestor partner values are $null then Enum Ancestors, and Reporte value[0] if found

            if (($Script:EnumerateAncestorPartners.result.result.id) -and ( $_.customerid -ne $Script:EnumerateAncestorPartners.result.result.id[0])) {
                send-EnumerateAncestorPartners $_.customerid
                
                if ($debugCDP) {
                    write-output "Getting ancestor data for next partnerid | $($_.CustomerId) | #$loop "
                    if ($Script:EnumerateAncestorPartners.result.result.id.count -gt 0) {
                        Write-Output "Found $($Script:EnumerateAncestorPartners.result.result.id[0])"
                    }else{
                        #Write-Output $Script:EnumerateAncestorPartners.error.message
                    }
                }
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
                
                $_.EndCustomer_id       = $endcustomer_level.id
                $_.EndCustomer          = $endcustomer_level.name
                $_.EndCustomer_legal    = $endcustomer_level.Company.LegalCompanyName
                $_.EndCustomer_ref      = $endcustomer_level.ExternalCode
            
                $_.Site_id              = $site_level.id
                $_.Site                 = $site_level.name
                $_.Site_legal           = $site_level.Company.LegalCompanyName
                $_.Site_ref             = $site_level.ExternalCode
            
                $loop ++
                if ($debugCDP) {Write-Output "Using cached ancestor data for partnerid | $($_.CustomerId) | #$loop "; }
            } ## If Ancestor partner[0] is the same as the report partnerid then use cached values  
                
            if ($Script:EnumerateAncestorPartners.error.message) {

                send-GetPartnerInfoHistory $_.customerid
                
                if ($debugCDP) { Write-Output "Getting Deleted Parent data for partnerid | $($_.CustomerId)" }

                if (($HistorypartnerId) -and ($Historylevel -eq "Site") ) {
                    if ($debugCDP) {Write-Output "Found Site data for partnerid | $($_.CustomerId)" }
                    $_.Site_id              = $HistorypartnerId
                    $_.Site                 = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "EndCustomer") ) {
                    if ($debugCDP) {Write-Output "Found EndCustomer data for partnerid | $($_.CustomerId)" }
                    $_.EndCustomer_id       = $HistorypartnerId
                    $_.EndCustomer          = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Reseller") ) {
                    if ($debugCDP) {Write-Output "Found Reseller data for partnerid | $($_.CustomerId)" }
                    $_.Reseller_id          = $HistorypartnerId
                    $_.Reseller             = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Subdistributor") ) {
                    if ($debugCDP) {Write-Output "Found SubDisti data for partnerid | $($_.CustomerId)" }
                    $_.SubDisti_id           = $HistorypartnerId
                    $_.SubDisti              = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
                if (($HistorypartnerId) -and ($Historylevel -eq "Distributor") ) {
                    if ($debugCDP) {Write-Output "Found Disti data for partnerid | $($_.CustomerId)" }
                    $_.Disti_id             = $HistorypartnerId
                    $_.Disti                = $historypartnername
                    send-GetPartnerInfoHistory $historyparentid
                }
            } ## If partner is <anonymized> or deleted then use GetPartnerInfoHistory to get parent data  
        }
    }
}  ## Get Maximum Value Report

Function Join-Reports {

    if (Get-Module -ListAvailable -Name Join-Object) {
        Write-Host "  Module Join-Object Already Installed"
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
    $Script:MVPlus | Where-Object {$_.DeviceDeletionDate} | Sort-Object DeviceId | foreach-object { Send-GetDeviceHistory $_.deviceid $_.DeviceDeletionDate; $_.Now_Physicality = $Script:DeviceHistoryDetail.physicality; $_.Now_Product = $Script:DeviceHistoryDetail.Product; $_.Now_timestamp= $Script:DeviceHistoryDetail.timestamp; $_.Now_Reference= $Script:DeviceHistoryDetail.Reference }
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
    $script:Endcustomer_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "EndCustomer"}
    $script:Site_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Site"}
}  ## Get Parent Partners

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

Send-APICredentialsCookie

Write-Output $Script:strLineSeparator
Write-Output ""

Send-GetPartnerInfo $cred0

if (($null -eq $Script:Period) -and ($null -eq $last)) {Get-Period}
if (($null -eq $Script:Period) -and ($last -eq 0)) { $Script:Period = (get-date) }
if (($null -eq $Script:Period) -and ($last -ge 1)) { $Script:Period = ((get-date).addmonths($last/-1)) }
#$Script:Period
#$Last


$null = New-Item -Path "$ExportPath" -Name "$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)" -ItemType "directory" 

$Script:CurrentExportPath = "$Exportpath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)"

$Script:AuditLogFile = "$CurrentExportPath\$($CurrentDate)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId)_AuditLog.csv"

Add-content -Path $AuditLogFile -Value "Company Id,Company Name,Agreement Name,Additions Product ID,Additions Description,Additions Quantity,Updated Quantity"

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
                        EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,
                        Site_Id,Site,Site_Legal,Site_Ref,
                        CustomerId,CustomerName,CustomerReference,ServiceType,
                        OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                        Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                        CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB |
        Export-Excel -Path "$XLSOutputFile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName ProductionUsage -FreezeTopRowFirstColumn -WorksheetName (get-date $Period -UFormat "%b-%Y") -tablestyle Medium6 -BoldTopRow -IncludePivotTable -PivotRows Disti,SubDisti,Reseller,EndCustomer -PivotDataToColumn -PivotData $pivotColumns
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

    Function AuthenticateCWM {
        if ( ($CWMAPICreds.pubKey -eq '') -or ($CWMAPICreds.privateKey -eq '') -or ($CWMAPICreds.Company -eq '') -or ($CWMAPICreds.clientId -eq '') -or ($CWMAPICreds.Server -eq '') ) {
            Write-Warning "ConnectWise Manage API Credentials not found in this script.`nExiting.`n`n"
            Add-content -Path $ExceptionsLogFile -Value "ConnectWise Manage API Credentials not found in this script."
            Break
        } else {
            Connect-CWM @CWMAPICreds
        }
    }  ## Connect to ConnectWise Manage instance

    Function GreenText  {
        process { Write-Host $_ -ForegroundColor Green }
    }  ## Output Green text

    Function LookupCWMProducts {
        foreach ($item in $CDP2CWMProductMapping.GetEnumerator()) {
            # https://arcanecode.com/2020/12/14/iterate-over-a-hashtable-in-powershell/
            if ($null -eq (Get-CWMProductCatalog -condition "identifier = `"$($item.Value)`"")) {
                Write-Output "`n";
                Write-Warning "Matching product NOT FOUND in ConnectWise Manage for Product ID: [$($item.Value)].`nPlease update CDP2CWMProductMapping variables in this script or ensure a product with the Product ID [$($item.Value)] exists in ConnectWise Manage.`n`n"
                Add-content -Path $ExceptionsLogFile -Value "Matching product NOT FOUND in ConnectWise Manage for Product ID: [$($item.Value)]."
            }
        }
    }  ## Lookup the Cove Data Protection Products in the ConnectWise Manage Product Catalog

    Function LookupCWMAgreements {
        Write-Output "`n  Checking for ConnectWise Manage Agreements"

        if ($CWMAgreementSearchType -eq "Contains") {
            $CWMAgreements = Get-CWMAgreement -condition "name contains `"$CWMAgreementName`" and agreementStatus = `"Active`"" -all    ## Contains $CWMAgreementName
        } elseif ($CWMAgreementSearchType -eq "StartsWith") {
            $CWMAgreements = Get-CWMAgreement -condition "name like `"$CWMAgreementName%`" and agreementStatus = `"Active`"" -all       ## StartsWith $CWMAgreementName
        } elseif ($CWMAgreementSearchType -eq "EndsWith") {
            $CWMAgreements = Get-CWMAgreement -condition "name like `"%$CWMAgreementName`" and agreementStatus = `"Active`"" -all       ## EndsWith $CWMAgreementName
        } else {
            $CWMAgreements = Get-CWMAgreement -condition "name = `"$CWMAgreementName`" and agreementStatus = `"Active`"" -all           ## Matches $CWMAgreementName
        }

        if ($null -eq $CWMAgreements) {
            Write-Output "`n";
            Write-Warning "No matching active agreements were found in your ConnectWise Manage instance for the Agreement Name: [$CWMAgreementName].`nPlease update the CWMAgreementName variable in this script or ensure agreements with the Agreement Name: [$CWMAgreementName] exist in your ConnectWise Manage instance.`n`n"
            Add-content -Path $ExceptionsLogFile -Value "No matching active agreements were found in your ConnectWise Manage instance for the Agreement Name: [$CWMAgreementName]."
            Break
        }else {
            Write-Output "    $($CWMAgreements.count) matching active agreements were found in ConnectWise Manage for the Agreement Name: [$CWMAgreementName].`n"
            $CWMAgreements | select-object company, @{n='agreementName';e={$_.name}} -expand company | select-object @{n='clientName';e={$_.company.name}}, agreementName | sort-object @{e={$_.company.Identifier}} | Format-Table
        }
    }  ## Look for ConnectWise Agreements with the specified Agreement Name in ConnectWise Manage

    Function LookupCWMCompany { # Get-CWMcompany -condition "id=28734"    
        $CWMcompany = $null
        
        $CWMCompanyID = (($CDPusage.PartnerName -split '\| ') -split ' ~')[2]
        if ($CWMCompanyID) {$CWMcompany = Get-CWMcompany -condition "id=$CWMCompanyID"}

        if ($CWMcompany) {
        ## Match Found for CWMCompanyId
        Write-Output "Matching cient found in ConnectWise Manage for Cove Data Protection CustomerName: [$($CDPusage.PartnerName)]`n  >>> ConnectWise Client ID: [$($CWMcompany.ID)], Client Name: [$($CWMcompany.Name)]`n" | GreenText
        } else {
            $CWMcompany = Get-CWMcompany -condition "name=`"$($CDPusage.PartnerName)`""
            if ($CWMcompany) {
                ## Match Found for Partner Name
                Write-Output "Matching client found in ConnectWise Manage for Cove Data Protection CustomerName: [$($CDPusage.PartnerName)]`n  >>> ConnectWise Client ID: [$($CWMcompany.ID)], Client Name: [$($CWMcompany.Name)]`n" | GreenText
            } else {
                ## NoMatch Found - Try Partner Legal Name
                $CWMcompany = Get-CWMcompany -condition "name=`"$($CDPusage.LegalName)`""
                if ($CWMcompany) {
                    ## Match Found for Partner Legal Name
                    Write-Output "Matching client found in ConnectWise Manage for Cove Data Protection Customer Legal Name: [$($CDPusage.LegalName)]`n  >>> ConnectWise Client ID: [$($CWMcompany.ID)], Client Name: [$($CWMcompany.Name)]`n" | GreenText
                } else {
                    ## NoMatch Found - Try Partner Reference Name
                    $CWMcompany = Get-CWMcompany -condition "name=`"$($CDPusage.PartnerRef)`""
                    
                    if ($CWMcompany) {
                        ## Match Found for Partner Reference Name
                        Write-Output "Matching client found in ConnectWise Manage for Cove Data Protection Customer Reference: [$($CDPusage.PartnerRef)]`n  >>> ConnectWise Client ID: [$($CWMcompany.ID)], Client Name: [$($CWMcompany.Name)]`n"  | GreenText
                    } else {
                        ## NoMatches Found - Output Warning
                        Write-Warning "Matching client NOT FOUND in ConnectWise Manage for Cove Data Protection CustomerID: [$($CDPusage.PartnerID)], PartnerName: [$($CDPusage.PartnerName)], LegalName: [$($CDPusage.LegalName)], Reference: [$($CDPusage.PartnerRef)]`n`n"
                        Add-content -Path $ExceptionsLogFile -Value "Matching client NOT FOUND in ConnectWise Manage for Cove Data Protection CustomerID: [$($CDPusage.PartnerID)], PartnerName: [$($CDPusage.PartnerName)], LegalName: [$($CDPusage.LegalName)], Reference: [$($CDPusage.PartnerRef)]"
                    }
                }
            }
        }

        if ($CWMcompany) {
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
                Write-Warning "Matching active agreement NOT FOUND in ConnectWise Manage for Agreement Name: [$CWMAgreementName] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
                Add-content -Path $ExceptionsLogFile -Value "Matching active agreement NOT FOUND in ConnectWise Manage for Agreement Name: [$CWMAgreementName] for ConnectWise Client Name: [$($CWMcompany.Name)]"
            } else {
                Write-Output "Matching active agreement found in ConnectWise Manage for Agreement Name: [$CWMAgreementName] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                if ($updateCWM) {
                    UpdateCWMQty    ## Update the Additions for the matching ConnectWise Manage Agreement with the quantities from the Cove Data Protection Usage
                }
                
            }
        }
    }  ## Lookup Cove Data Protection Customer in ConnectWise Manage

    Function UpdateCWMQty {
        $CWMAdditions = Get-CWMAgreementAddition -AgreementID $CWMAgreement.id -Condition "agreementStatus = `"Active`"" -all

        $PhysicalServerMatched = $false
        $VirtualServerMatched = $false
        $WorkstationMatched = $false
        $DocumentsMatched = $false
        $RecoveryTestingMatched = $false
        $O365UsersMatched = $false
        $CoveGBSelectedMatched = $false

        if ($debugCWM) {
            Write-Output "  $($CDPusage.partnername) >>> $($CWMcompany.name) | $($CWMAgreement.name)"
            Write-Output "  Addition quantities before updates"
            $CWMAdditions | select-object product, quantity -expand product | select-object @{n='identifier';e={$_.product.identifier}}, quantity | format-table
        }

        foreach ($CWMAddition in $CWMAdditions) {
            $Update = @{
                AgreementID = $CWMAgreement.id
                AdditionID = $CWMAddition.id
                Operation = 'replace'
                Path = 'quantity'
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.PhysicalServerQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.PhysicalServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.PhysicalServerQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.PhysicalServerQty | out-null
                $PhysicalServerMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.VirtualServerQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.VirtualServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.VirtualServerQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.VirtualServerQty | out-null
                $VirtualServerMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.WorkstationQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.WorkstationQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.WorkstationQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.WorkstationQty | out-null
                $WorkstationMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.DocumentsQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.DocumentsQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for the ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.DocumentsQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.DocumentsQty | out-null
                $DocumentsMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.RecoveryTestingQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.RecoveryTestingQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.RecoveryTestingQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.RecoveryTestingQty | out-null
                $RecoveryTestingMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.O365UsersQty))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.O365UsersQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for the ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.O365UsersQty)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.O365UsersQty | out-null
                $O365UsersMatched = $true
            }
            if ($($CWMAddition.product.identifier).$CWMAdditionSearchType($($CDP2CWMProductMapping.CoveGBSelected))) {
                Write-Output "Matching active addition found for ConnectWise Product ID: [$($CDP2CWMProductMapping.CoveGBSelected)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n" | GreenText
                Add-content -Path $AuditLogFile -Value "`"$($CWMcompany.id)`",`"$($CWMcompany.name)`",`"$($CWMAgreement.name)`",`"$($CWMAddition.product.identifier)`",`"$($CWMAddition.description)`",$([Int]$CWMAddition.quantity),$($CDPusage.CoveGBSelected)"
                Update-CWMAgreementAddition @Update -Value $CDPusage.CoveGBSelected | out-null
                $CoveGBSelectedMatched = $true
            }
        }

        ## Show list of products Report an exception if Usage Qty > 0 and no matching Addition was found
        if ($debugCWM) {
            start-sleep -Milliseconds 2500    ## Wait before retrieving the updated Additions
            $CWMAdditions = Get-CWMAgreementAddition -AgreementID $CWMAgreement.id -Condition "agreementStatus = `"Active`"" -all
            Write-Output "  $($CDPusage.partnername) >>> $($CWMcompany.name) | $($CWMAgreement.name)"
            Write-Output "  Addition quantities after updates"
            $CWMAdditions | select-object product, quantity -expand product | select-object @{n='identifier';e={$_.product.identifier}}, quantity | format-table
        }

        ## Report an exception if Usage Qty > 0 and no matching Addition was found

        if ($CDPusage.PhysicalServerQty -gt 0 -and $PhysicalServerMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.PhysicalServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.PhysicalServerQty)] in the ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.VirtualServerQty -gt 0 -and $VirtualServerMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.VirtualServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.VirtualServerQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.WorkstationQty -gt 0 -and $WorkstationMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.WorkstationQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$CDP2CWMProductMapping.WorkstationQty] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.DocumentsQty -gt 0 -and $DocumentsMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.DocumentsQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.DocumentsQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.RecoveryTestingQty -gt 0 -and $RecoveryTestingMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.RecoveryTestingQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$CDP2CWMProductMapping.RecoveryTestingQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }
        if ($CDPusage.O365UsersQty -gt 0 -and $O365UsersMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.O365UsersQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.O365UsersQty)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }

        if ($CDPusage.CoveGBSelected -gt 0 -and $CoveGBSelectedMatched -eq $false) {
            Write-Warning "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.CoveGBSelected)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]`n`n"
            Add-content -Path $ExceptionsLogFile -Value "Matching active addition NOT FOUND in ConnectWise Manage for Product ID: [$($CDP2CWMProductMapping.CoveGBSelected)] in ConnectWise Agreement Name: [$($CWMAgreement.name)] for ConnectWise Client Name: [$($CWMcompany.Name)]"
        }


    }  ## Update the Additions for the matching ConnectWise Manage Agreement with the quantities from the Cove Data Protection Usage

    InstallCWMPSModule         ## Install the ConnectWise Manage PowerShell module
    AuthenticateCWM            ## Authenticate with ConnectWise Manage
    LookupCWMProducts          ## Lookup the Data Protection Products in the ConnectWise Manage Product Catalog
    if ($debugCWM) {
        LookupCWMAgreements    ## Look for Agreements with the specified Agreement Name in ConnectWise Manage
    }

    If ($GetCDPUsage) {
        $Usage = $MVplus | Where-object {$_.CustomerState -in $StateToProcess} |
            Select-Object Subdisti,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Id,
            Endcustomer,EndCustomer_legal,Endcustomer_Ref,EndCustomer_id,
            PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,OverageGB | Sort-Object endcustomer -Unique

        ## Update the Usage for each Customer
        foreach ($customer in $Usage) {
            $CDPusage =  @()
            $CDPusage += New-Object -TypeName PSObject -Property @{
                PartnerID           = $customer.endcustomer_id ;
                Partnername         = $customer.endcustomer ;
                LegalName           = $customer.endcustomer_Legal ;
                PartnerRef          = $customer.endcustomer_Ref ;
                Site                = $customer.Site ;
                EndCustomer         = $customer.EndCustomer ;
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

            Write-Output "`nCove Data Protection High Water Mark Usage for Month ending $end"
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
            
            $CDPusage | Select-Object   SubDisti,Reseller,EndCustomer,PartnerID,Partnername,LegalName,PartnerRef,
                                        PhysicalServerQty,VirtualServerQty,WorkstationQty,DocumentsQty,RecoveryTestingQty,O365UsersQty,@{n='CoveGBSelected';e={[int]$_.CoveGBSelected}} |
                                        Export-Csv -Path $CDPUsageoutputFile -Append -NoTypeInformation

            LookupCWMCompany    ## Lookup Cove Data Protection Customer in ConnectWise Manage
        }
    }
    if ($ImportCDPUsage) {
        #Open dialog to load import file
        $OpenFileName = Open-FileName "$ExportPath" ; if ($null -eq $OpenFileName) {Break} 
        #Import CSV file to $cdpusage
        $Usage = Import-Csv -Path $OpenFileName
        #Validate CSV file
        $expectedHeaders =  "SubDisti","Reseller","EndCustomer",
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

        $SelectedUsage = $Usage | Out-GridView -Title "Select End-Customer Usage to Process from $($openfilename)" -OutputMode Multiple
        
        if ($SelectedUsage -eq $null) {
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

            LookupCWMCompany    ## Lookup Cove Data Protection Customer in ConnectWise Manage
        }
    }
} else {
    Write-Warning "ConnectWise Manage can not be reached as the ConnectCWM switch is disabled."
}  ## Execute if ConnectCWM switch is enabled
    
#endregion ----- ConnectWise Manage Body / Functions  ----
    Start-Sleep -seconds 5
