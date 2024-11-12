# ----- About: ----
    # N-able Cove Data Protection MaxValue Plus Usage Report > Multi-select
    # Revision v24.11.12
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
    # For use with the Standalone edition of N-able | Cove Data Protection 
    # Tested with Cove Data Protection v24.10 and Powershell 5.1 & 7.4
    # Some script elements may be developed, tested or documented using AI
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials for Cove Data Protection API
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select one or multiple partners
    # Get Current Device Statistics
    # Get Cove Maximum Value Usage Report for Period
    # Export usage to CSV/XLS & Pivot Table
    #
    # Use the -Period switch parameter to define Usage Dates (yyyy-MM or MM-yyyy)
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherlands)
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path. Default is the execution path of the script.
    # Use the -DebugCDP switch to display debug info for Cove Data Protection usage and statistics lookup
    # Use the -ClearCredentials parameter to remove stored Cove API credentials at start of script
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console/export.htm
    #
    # https://www.n-able.com/legal/cove-data-protection-fair-use-policy
    # https://www.n-able.com/legal/cove-data-protection-fair-use-policy-dpp
    #
   
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (        
        [Parameter(Mandatory=$False)][datetime]$Period,                                                 ## Lookup Date yyyy-MM or MM-yyyy
        [Parameter(Mandatory=$False)][ValidateRange(0,24)] $Last = 1,                                   ## Count back # Last Months i.e. 0 current, 1 Prior Month
        [Parameter(Mandatory=$False)][switch]$AllPartners = $false,                                     ## If $true Skip GUI partner selection
        [Parameter(Mandatory=$False)][int]$DeviceCount = 5000,                                          ## Change Maximum Number of current devices results to return
        [Parameter(Mandatory=$False)][Switch]$DebugCDP = $false,                                        ## If $true Enable Debug for Cove Max Value / Device Statistics
        [Parameter(Mandatory=$False)][switch]$Launch = $true,                                           ## If $true Launch XLS or CSV file 
        [Parameter(Mandatory=$False)][string]$Delimiter = ',',                                          ## Specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",                             ## Export Path
        [Parameter(Mandatory=$False)][int]$PhysServerFUGB = 2000,                                       ## Set your contracted Selected Size Fare Use GB capacity per Physical Server
        [Parameter(Mandatory=$False)][int]$VirtServerFUGB = 2000,                                       ## Set your contracted Selected Size Fare Use GB capacity per Virtual Server
        [Parameter(Mandatory=$False)][int]$WorkstationFUGB = 500,                                       ## Set your contracted Selected Size Fare Use GB capacity per Workstation
        [Parameter(Mandatory=$False)][switch]$ClearCredentials                                          ## If $true Clear Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----

    Clear-Host
    #Requires -RunAsAdministrator
    $ConsoleTitle = "Cove Data Protection MaxValue Plus Usage Report"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-Output "  $ConsoleTitle`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    if ($ExportPath -eq "") {$ExportPath = Split-Path $MyInvocation.MyCommand.Path}
    if ($ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "CoveMaxValuePlusReport_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "CoveMaxValuePlusReport_$shortdate"}
    If ($ExportPath) {mkdir -force -path $ExportPath | Out-Null}
    Push-Location $dir

    Write-Output "  Current Parameters:"
    Write-Output "  -Period             = $Period"
    Write-Output "  -Month Prior        = $Last"
    Write-Output "  -AllPartners        = $AllPartners"
    Write-Output "  -AllDevices         = $AllDevices"
    Write-Output "  -DeviceCount        = $DeviceCount"
    Write-Output "  -Launch             = $Launch"
    Write-Output "  -ExportPath         = $ExportPath"
    Write-Output "  -Delimiter          = $Delimiter"
    Write-Output "  -PhysServerFUGB     = $PhysServerFUGB"
    Write-Output "  -VirtServerFUGB     = $VirtServerFUGB"
    Write-Output "  -WorkstationFUGB    = $WorkstationFUGB"
    Write-Output "  -DebugCDP           = $DebugCDP"
    Write-Output "  -SendToCWM          = $SendtoCWM"
    Write-Output "  -DebugCWM           = $DebugCWM"



    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
    $urljson = "https://api.backup.management/jsonapi"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 

            Write-Output "  Enter Exact, Case Sensitive Partner Name for Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($PartnerName.length -eq 0)
        $PartnerName | out-file $APIcredfile

        $BackupCred = Get-Credential -Message 'Enter Login Email and Password for Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
        
        Start-Sleep -milliseconds 300

        Send-APICredentialsCookie  ## Attempt API Authentication

    } ## Set API credentials if not present

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
                    Write-Output "  Stored Backup API Partner  = $Script:cred0"
                    Write-Output "  Stored Backup API User     = $Script:cred1"
                    Write-Output "  Stored Backup API Password = Encrypted"
                    
                }else{
                    Write-Output    $Script:strLineSeparator 
                    Write-Output "  Backup API Credential File Not Present"

                    Set-APICredentials  ## Create API Credential File if Not Found
                    }
                }

    } ## Get API credentials if present

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

    } ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Send-APICredentialsCookie
            }
            
        }
    } ## Renew Visa

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $DateTime = $epoch.AddSeconds($inputUnixTime)
        return $DateTime
        }else{ return ""}
    } ## Convert epoch time to date time 

    Function Convert-DateTimeToUnixTime($DateToConvert) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $NewExtensionDate = Get-Date -Date $DateToConvert
        [int64]$NewEpoch = (New-TimeSpan -Start $epoch -End $NewExtensionDate).TotalSeconds
        Return $NewEpoch
    } ## Convert date time to epoch time

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
        $Calendar.MinDate = (Get-Date).AddMonths(-24)   # Minimum Date Dispalyed
        $Calendar.MaxDate = (Get-Date)
        #$Calendar.MinDate = "01/01/2020"               # Minimum Date Displayed
        #$Calendar.MaxDate = "12/31/2021"               # Maximum Date Displayed
        $Calendar.SetCalendarDimensions([int]3,[int]1)  # 3x1 Grid
        $TrialForm.Controls.Add($Calendar)

        $TrialForm.Add_Shown($TrialForm.Activate())
        $result = $TrialForm.showdialog()
        #$Calendar.SelectionEnd

        if ($result -eq [Windows.Forms.DialogResult]::OK) {
            $Script:Period = $calendar.SelectionEnd
        
            #Write-Output "Date selected: $($Period.ToShortDateString())"
            Write-Output "Date selected: $(get-date($period) -Format 'MMMM yyyy')"
        }

        if ($result -eq [Windows.Forms.DialogResult]::Cancel) {
            $Script:Period = get-date
        }
    } ## GUI Select Report Period

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

Function Send-GetPartnerInfo2 ($PartnerName) { 
    $url = "https://api.backup.management/jsonapi"
    $data = @{
        jsonrpc = '2.0'
        id = '2'
        visa = $Script:visa
        method = 'GetPartnerInfo'
        params = @{
            name = [String]$PartnerName
        }
    }

    # Convert the data to JSON
    $jsonString = ConvertTo-Json $data -Depth 5

    # Make the REST request
    $response = Invoke-RestMethod -Method POST `
        -ContentType 'application/json; charset=UTF-8' `
        -Body $jsonString `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing

    # Store the response data
    $Script:websession = $websession
    $Script:Partner = $response
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

        # Create the JSON string
        $jsonString = ConvertTo-Json $data -Depth 5

        # Convert the JSON string to bytes (UTF-8 encoding)
        $jsonBytes = [System.Text.Encoding]::UTF8.GetBytes($jsonString)

        # Make the HTTP request
        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json; charset=UTF-8' `
            -Body $jsonBytes `
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
            }else{
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

    } ## get PartnerID and Partner Level    

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
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            $Script:PartnerId = Read-Host "  Enter Customer/ Tenant / Partner Id to lookup i.e. '192003'"
            Send-GetPartnerInfo $Script:partnername
        }

        if ($partner.error) {
            Write-Output "  $($partner.error.message)"
            $Script:PartnerId = Read-Host "  Enter Customer/ Tenant / Partner Id to lookup i.e. '192003'"
            Send-GetPartnerInfo $Script:partnername
        }

    } ## get PartnerID and Partner Level    

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
    } ## Generic Json Call

    Function Send-EnumeratePartners {
        # ----- Get Partners via EnumeratePartners -----
        
        # (Create the JSON object to call the EnumeratePartners function)
            $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
                Add-Member -PassThru NoteProperty visa $Script:visa |
                Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
                Add-Member -PassThru NoteProperty params @{
                                                            parentPartnerId = $PartnerId 
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
        }Else{ # (No error)
            $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,*,Company -ErrorAction Ignore
            
            $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
        
            $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                    
            $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
            
            
            if ($AllPartners) {
                $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid,@{l='VatNumber';e={$_.Company.VatNumber}} | sort-object Level,name
                Write-Output    $Script:strLineSeparator
                Write-Output    "  All Partners Selected"
            }else{
                [array]$Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid,@{l='VatNumber';e={$_.Company.VatNumber}} | sort-object Level,name | out-gridview -Title "Current Partner | $partnername | Select the <ParentPartner> for a consolidated usage report or multiple child partners for individual usage reports" -OutputMode Multiple
        
              
            }
        }
        
    } ## EnumeratePartners API Call

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
                                                                                SelectedGB     = [Math]::Round([Decimal](($DeviceResult.Settings.T3 -join '') /1GB),2) ;                        
                                                                                UsedGB         = [Math]::Round([Decimal](($DeviceResult.Settings.US -join '') /1GB),2) ;  
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

    } ## EnumerateAccountStatistics API Call

    Function Send-GetDeviceHistory ($DeviceID, $DeleteDate) {
        #$history = (get-date $period)
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
            if ($debugCDP) {Write-Output "Getting historic data for deleted deviceid | $($DeviceHistory.AccountId)"}
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
                                                                    SelectedGB     = [Math]::Round([Decimal](($DeviceHistory.Settings.T3 -join '') /1GB),2) ;                        
                                                                    UsedGB         = [Math]::Round([Decimal](($DeviceHistory.Settings.US -join '') /1GB),2) ; 
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
    } ## EnumerateAccountStatistics API Call

    Function GetMVReport {
        Param ([Parameter(Mandatory=$False)][Int]$PartnerId) #end param

        $script:Start = (Get-Date $period -Day 1).AddMonths(0).ToString('yyyy-MM-dd')
        $script:End = (Get-Date $period -Day 1).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
        Write-Output "  Requesting Maximum Value Report from $start to $end"
        
        $Script:TempMVReport = "$exportpath\TempMVReport.xlsx"
        remove-item $Script:TempMVReport -Force -Recurse -ErrorAction SilentlyContinue
        
        $script:url2 = "https://api.backup.management/statistics-reporting/high-watermark-usage-report?_=b19c8a811a5a7&dateFrom=$($Start)T00%3A00%3A00Z&dateTo=$($end)T23%3A59%3A59Z&exportOutput=OneXlsxFile&partnerId=$($PartnerId)"
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

        $Script:MVReportxls = Import-Excel -path "$Script:TempMVReport" -asDate "*Date" 

        $Script:MVReport = $Script:MVReportxls | select-object  Disti_Id,Disti,Disti_Legal,Disti_Ref,Disti_Vat,
                                                                SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,SubDisti_Vat,
                                                                Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Vat,
                                                                EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,EndCustomer_Vat,
                                                                Site_Id,Site,Site_Legal,Site_Ref,Site_Vat,
                                                                CustomerId,CustomerName,CustomerReference,CustomerState,ProductionDate,ServiceType,
                                                                OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,CreationDate,DeviceDeletionDate,
                                                                UsedStorageGb,UsedStorageMvDate,SelectedSizeGb,SelectedSizeMvDate,O365Users,O365UsersMvDate,RecoveryTesting,RecoveryTestingMvDate,
                                                                @{n='PhysicalServer';e='Null'},@{n='VirtualServer';e='Null'},@{n='Workstation';e='Null'},@{n='Documents';e='Null'},@{n='FairUseGB';e='Null'},@{n='OverageGB';e='Null'}

        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "0"){$_.RecoveryTesting = $null}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "True"){$_.RecoveryTesting = "1"}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "1"){$_.RecoveryTesting = "RecoveryTesting"}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "2"){$_.RecoveryTesting = "StandbyImage"}}
        #$Script:MVReport | foreach-object { if ($_.O365Users -eq "0"){$_.O365Users = $null}}
        $Script:MVReport | foreach-object { if ($_.customerid){
                send-EnumerateAncestorPartners $_.customerid
                
                $_.Disti_id             = $distributor_level.id
                $_.Disti                = $distributor_level.name
                $_.Disti_legal          = $distributor_level.Company.LegalCompanyName
                $_.Disti_ref            = $distributor_level.ExternalCode
                $_.Disti_vat            = $distributor_level.Company.VatNumber

                $_.SubDisti_id          = $subdistributor_level.id
                $_.SubDisti             = $subdistributor_level.name
                $_.SubDisti_legal       = $subdistributor_level.Company.LegalCompanyName
                $_.SubDisti_ref         = $subdistributor_level.ExternalCode
                $_.SubDisti_vat         = $subdistributor_level.Company.VatNumber

                $_.Reseller_id          = $reseller_level.id
                $_.Reseller             = $reseller_level.name
                $_.Reseller_legal       = $reseller_level.Company.LegalCompanyName
                $_.Reseller_ref         = $reseller_level.ExternalCode
                $_.Reseller_vat         = $reseller_level.Company.VatNumber

                $_.EndCustomer_id       = $endcustomer_level.id
                $_.EndCustomer          = $endcustomer_level.name
                $_.EndCustomer_legal    = $endcustomer_level.Company.LegalCompanyName
                $_.EndCustomer_ref      = $endcustomer_level.ExternalCode
                $_.EndCustomer_vat      = $endcustomer_level.Company.VatNumber

                $_.Site_id              = $site_level.id
                $_.Site                 = $site_level.name
                $_.Site_legal           = $site_level.Company.LegalCompanyName
                $_.Site_ref             = $site_level.ExternalCode
                $_.Site_vat             = $site_level.Company.VatNumber
            }
        }
    } ## Download Max Value Report

   Function Join-Reports {

        $Script:MVPlus = Join-Object -left $Script:MVReport -right $Script:SelectedDevices -LeftJoinProperty DeviceId -RightJoinProperty AccountID -Prefix 'Now_' -Type AllInBoth
        $Script:MVPlus | Where-Object {$_.DeviceDeletionDate} | foreach-object { Send-GetDeviceHistory $_.deviceid $_.DeviceDeletionDate; $_.Now_Physicality = $Script:DeviceHistoryDetail.physicality; $_.Now_Product = $Script:DeviceHistoryDetail.Product; $_.Now_timestamp= $Script:DeviceHistoryDetail.timestamp; $_.Now_Reference= $Script:DeviceHistoryDetail.Reference }
    } ## Install Join-Object PS module to merge statistics and Max Value Report

    Function Send-EnumerateAncestorPartners ($PartnerID) {

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
            TimeoutSec  = 180  # Set the timeout value in seconds
        }  

        $Script:EnumerateAncestorPartners = Invoke-RestMethod @params 

        $script:subroot_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "SubRoot"}
        $script:distributor_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Distributor"}
        $script:subdistributor_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "SubDistributor"}
        $script:reseller_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Reseller"}
        $script:Endcustomer_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "EndCustomer"}
        $script:Site_level = $EnumerateAncestorPartners.result.result | Where-object {$_.level -eq "Site"}

    } ## Get Parent Partners

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
    Write-Output ""
    
    Send-GetPartnerInfo $Script:cred0

    if (($null -eq $Script:period) -and ($null -eq $last)) {Get-Period}
    if (($null -eq $Script:period) -and ($last -eq 0)) { $Script:period = (get-date) }
    if (($null -eq $Script:period) -and ($last -ge 1)) { $Script:period = ((get-date).addmonths($last/-1)) }
    
    #$Script:Period
    #$Last

    if (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host "  Module ImportExcel Already Installed"
    } 
    else {
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
    
    if (Get-Module -ListAvailable -Name Join-Object) {
        Write-Host "  Module Join-Object Already Installed"
    }else{
        try {
            Install-Module -Name Join-Object -Confirm:$False -Force      ## link
        }
        catch [Exception] {
            $_.message 
            exit
        }
    }

    if ($AllPartners) {}else{Send-EnumeratePartners}
    
    foreach ($selectedpartner in $Script:Selection) {
    
    try {

        if($null -eq $Selection) {
            # Cancel was pressed
            # Run cancel script
            Write-Output    $Script:strLineSeparator
            Write-Output    "  No Partners Selected"
            Exit
        }else{
            # OK was pressed, $Selection contains what was chosen
            # Run OK script
            [int]$Script:PartnerId = $Script:selectedpartner.Id
            [String]$Script:PartnerName = $Script:selectedpartner.Name
            Write-Output "`n  Pulling data for $Partnername`n"
        }



            Send-GetDevices $script:partnerId
            
            $Script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,ComputerName,
                    Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,
                    SelectedGB,UsedGB,Location,OS,OSType,Physicality

            
            GetMVReport $partnerId
            join-reports
            
            if($null -eq $Script:SelectedDevices) {
                # Cancel was pressed
                # Run cancel script
                Write-Output    $Script:strLineSeparator
                Write-Warning    "No Devices Selected"
                #Exit
            }else{
                # OK was pressed, $Selection contains what was chosen
                # Run OK script
            }

            $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_MaxValuePlus_DPP_$($period.ToString('yyyy-MM'))_Statistics_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"

            $Script:MVPlus | foreach-object { if ($_.SelectedsizeGB -gt 0) {$_.SelectedsizeGB = [math]::round($_.SelectedsizeGB,2)}}
            $Script:MVPlus | foreach-object { if ($_.UsedStorageGB -gt 0) {$_.UsedStorageGB = [math]::round($_.UsedStorageGB,2)}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Physical")) {$_.PhysicalServer = "1"}else{$_.PhysicalServer = "0"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Virtual")) {$_.VirtualServer = "1"}else{$_.VirtualServer = "0"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -eq "Documents")) {$_.Documents = "1"}else{$_.Documents = "0"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -ne "Documents")) {$_.Workstation = "1"}else{$_.Workstation = "0"}}
            $Script:MVPlus | foreach-object { if (($_.Now_Physicality -eq "Undefined") -and ($_.O365Users)) {$_.Now_Physicality = "Cloud"; $_.OSType = "M365";  $_.ComputerName = "* M365 - $($_.DeviceName)"}}
            $Script:MVPlus | foreach-object { if ($_.PhysicalServer -eq 1) {$_.FairUseGB = $PhysServerFUGB }}
            $Script:MVPlus | foreach-object { if ($_.VirtualServer -eq 1) {$_.FairUseGB = $VirtServerFUGB }}
            $Script:MVPlus | foreach-object { if ($_.Workstation -eq 1) {$_.FairUseGB = $WorkstationFUGB }}
            $Script:MVPlus | foreach-object { if ($_.Documents -eq 1) {$_.FairUseGB = 0 ; $_.SelectedsizeGB = 0 }}
            $Script:MVPlus | foreach-object { if (($_.PhysicalServer -eq 1) -or ($_.VirtualServer -eq 1) -or ($_.Workstation -eq 1)) {$_.OverageGB = ($_.SelectedSizeGB - $_.FairuseGB) }}

            #PhysServerGBFU,VirtServerGBFU,WorkstationGBFU
            
            $Script:MVPlus | Where-object {$_.CustomerState -eq "InProduction"} | where-object {$_.CustomerName -notlike '*Recycle Bin'} | 
                Select-object Disti_Id,Disti,Disti_Legal,Disti_Ref,Disti_Vat,
                            SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,SubDisti_Vat,
                            Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Vat,
                            EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,EndCustomer_Vat,
                            Site_Id,Site,Site_Legal,Site_Ref,Site_Vat,
                            CustomerId,CustomerName,CustomerReference,ServiceType,
                            OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                            Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                            CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB|
                Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8

            $Script:xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")

            $ConditionalFormat =$(New-ConditionalText -ConditionalType DuplicateValues -Range 'AC:AC' -BackgroundColor CornflowerBlue -ConditionalTextColor Black)

            $Script:MVPlus | Where-object {$_.CustomerState -eq "InTrial"} | where-object {$_.CustomerName -notlike '*Recycle Bin'} | 
                            Select-object Disti_Id,Disti,Disti_Legal,Disti_Ref,Disti_Vat,
                            SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,SubDisti_Vat,
                            Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Vat,
                            EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,EndCustomer_Vat,
                            Site_Id,Site,Site_Legal,Site_Ref,Site_Vat,
                            CustomerId,CustomerName,CustomerReference,ServiceType,
                            OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                            Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                            CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB |
                            Export-Excel -Path "$xlsoutputfile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName TrialUsage -FreezeTopRowFirstColumn -WorksheetName "TRIAL $(get-date $period -UFormat `"%b-%Y`")" -tablestyle Medium6 -BoldTopRow

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
                            
            $Script:MVPlus | Where-object {$_.CustomerState -eq "InProduction"} | where-object {$_.CustomerName -notlike '*Recycle Bin'} | 
                            Select-object Disti_Id,Disti,Disti_Legal,Disti_Ref,Disti_Vat,
                            SubDisti_Id,SubDisti,SubDisti_Legal,SubDisti_Ref,SubDisti_Vat,
                            Reseller_Id,Reseller,Reseller_Legal,Reseller_Ref,Reseller_Vat,
                            EndCustomer_Id,EndCustomer,EndCustomer_Legal,EndCustomer_Ref,EndCustomer_Vat,
                            Site_Id,Site,Site_Legal,Site_Ref,Site_Vat,
                            CustomerId,CustomerName,CustomerReference,ServiceType,
                            OsType,CurrentMonthMvSKU,DeviceId,DeviceName,ComputerName,
                            Now_Product,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,
                            CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,Now_TimeStamp,SelectedSizeGB,UsedStorageGB,FairUseGB,OverageGB |
                            Export-Excel -Path "$xlsoutputfile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName ProductionUsage -FreezeTopRowFirstColumn -WorksheetName (get-date $period -UFormat "%b-%Y") -tablestyle Medium6 -BoldTopRow -IncludePivotTable -PivotRows Disti,SubDisti,Reseller,endcustomer -PivotDataToColumn -PivotData $pivotColumns
        
            if ($Launch) {
                If (test-path HKLM:SOFTWARE\Classes\Excel.Application) {                     Start-Process "$xlsoutputfile"
                    Write-Output $Script:strLineSeparator
                    Write-Output "  Opening XLS file"
                    }else{
                    Start-Process "$csvoutputfile"
                    Write-Output $Script:strLineSeparator
                    Write-Output "  Opening CSV file"
                    Write-Output $Script:strLineSeparator            
                    }
                } ## Launch CSV or XLS (if Excel is installed)  (Required -Launch Parameter)
                
            Write-Output $Script:strLineSeparator
            Write-Output "  Cove Usage - CSV Path = &`"$csvoutputfile`""
            Write-Output "  Cove Usage - XLS Path = &`"$xlsoutputfile`""
            Write-Output ""

        }catch {Write-Warning "An error occured getting data for <$Script:PartnerName>, this partner may be new or not have usage for the selected period" }
    }

