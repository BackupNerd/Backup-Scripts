# ----- About: ----
    # N-central Integrated Backup MaxValue Plus Usage Report  
    # Revision v21 - 2022-06-23
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
    # For use with the N-central Integrated edition of N-able Backup
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Requires N-central Backup API access credentials, contact @Backup_Nerd to request
    # Check/ Get/ Store secure credentials 
    # Authenticate to Backup infrastucture
    # Get Maximum Value Report for Period
    # Get Current Device Statistics
    # Merge Reports
    # Optionally export to XLS/CSV
    #
    # Use the -Period switch parameter to define Report Dates (yyyy-MM or MM-yyyy)
    # Use the -DeviceCount ## (default=20000) parameter to define the maximum number of devices returned
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherlands)
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path. Default is the execution path of the script.
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console/export.htm


# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Data Protection Plan Usage")]
    Param (
        
        [Parameter(Mandatory=$False)] [Switch]$DPP,                             ## Export Monthly Data Protection Plan Usage
        [Parameter(Mandatory=$False)][datetime]$Period,                         ## Lookup Date yyyy-MM or MM-yyyy
        [Parameter(Mandatory=$False)][ValidateRange(0,24)] $Last,               ## Count back # Last Months i.e. 0 current, 1 Prior Month
        [Parameter(Mandatory=$False)][int]$DeviceCount = 5000,                 ## Change Maximum Number of current devices results to return
        [Parameter(Mandatory=$False)][switch]$Launch = $true,                   ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)][string]$Delimiter = ',',                  ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",     ## Export Path
        [Parameter(Mandatory=$False)][switch]$ClearCredentials                  ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "N-central MSPBackup MaxValue Plus Usage Report"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-output "  $ConsoleTitle`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"

    Write-output "  Current Parameters:"
    Write-output "  -Period      = $Period"
    Write-output "  -Month Prior = $Last"
    Write-output "  -DeviceCount = $DeviceCount"
    Write-output "  -Launch      = $Launch"
    Write-output "  -ExportPath  = $ExportPath"
    Write-output "  -Delimiter   = $Delimiter"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
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
 
            Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($PartnerName.length -eq 0)
        $PartnerName | out-file $APIcredfile

        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
        
        Start-Sleep -milliseconds 300

        Send-APICredentials  ## Attempt API Authentication

    }  ## Set API credentials if not present

Function Get-APICredentials {

        $Script:True_path = "C:\ProgramData\MXB\"
        $Script:APIcredfile = join-path -Path $True_Path -ChildPath "$env:computername NC_API_Credentials.Secure.txt"
        $Script:APIcredpath = Split-path -path $APIcredfile
    
        if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
            Remove-Item -Path $Script:APIcredfile
            $ClearCredentials = $Null
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential File Cleared"
            Send-APICredentials  ## Retry Authentication
            
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
                    Write-output "  Stored N-central Backup API Partner  = $Script:cred0"
                    Write-output "  Stored N-central Backup API User     = $Script:cred1"
                    Write-output "  Stored N-central Backup API Password = Encrypted"
                    
                }else{
                    Write-Output    $Script:strLineSeparator 
                    Write-Output "  N-central Backup API Credential File Not Present"
    
                    Set-APICredentials  ## Create API Credential File if Not Found
                    }
                }
    
    }  ## Get API credentials if present

Function Send-APICredentials {

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

    #Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($authenticate.visa) { 

        $Script:visa = $authenticate.visa
        $Script:PartnerId = $authenticate.result.result.PartnerId
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate
    
Function Send-APICredentialsNC {
    #Remove-TypeData -ErrorAction Ignore System.Array

    #$url = "https://api.integrated.cloudbackup.management/jsonapi"
    $url = "https://api.backup.management/jsonapi"
    $script:data = @{}
    $data.jsonrpc = '2.0'
    $data.id = 'jsonrpc'
    $data.method = 'AuthenticateWithContext'
    $data.params = @{}
    $data.params.authenticationContext = @{}
    $data.params.authenticationContext.Labels = @()
    $data.params.authenticationContext.PartnerId = $($Script:PartnerId)
    $data.params.authenticationContext.Style = "" 
    $data.params.authenticationContext.ViewMode = ""
    $data.params.authenticationContext.VisiblePartnerEntityName = "Customer"
    $data.params.authenticationContext.nCentralVersion = "12345"
    $data.params.partner = $Script:cred0
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    $script:webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -Depth 4) `
        -Uri $url `
        -TimeoutSec 30 `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:NCAuthenticate = $webrequest | convertfrom-json

    #Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

    if ($Script:NCauthenticate.visa) { 

        $Script:ncvisa = $Script:NCAuthenticate.visa
        #Debug $Script:ncvisa
        #$Script:UserId = $Script:authenticate.result.result.id
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
           
        }

    }  ## Special N-central Authentication

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

    $DateForm = New-Object Windows.Forms.Form
    $DateForm.text = "Select Any Date for that Months Report"
    $DateForm.Font = [System.Drawing.Font]::new('Arial',15, [System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
    $DateForm.BackColor = [Drawing.Color]::SteelBlue
    $DateForm.AutoSize = $False
    $DateForm.MaximizeBox = $False
    $DateForm.Size = New-Object Drawing.Size(750,350)
    $DateForm.ControlBox = $True
    $DateForm.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    $DateForm.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(10,220)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $DateForm.AcceptButton = $okButton
    $DateForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(85,220)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $DateForm.CancelButton = $cancelButton
    $DateForm.Controls.Add($cancelButton)

    $Calendar = New-Object System.Windows.Forms.MonthCalendar 
    $Calendar.Location = "10,10"
    $Calendar.MaxSelectionCount = 1
    $Calendar.MinDate = (Get-Date).AddMonths(-24)   # Minimum Date Dispalyed
    $Calendar.MaxDate = (Get-Date)
    #$Calendar.MinDate = "01/01/2020"               # Minimum Date Displayed
    #$Calendar.MaxDate = "12/31/2021"               # Maximum Date Displayed
    $Calendar.SetCalendarDimensions([int]3,[int]1)  # 3x1 Grid
    $DateForm.Controls.Add($Calendar)

    $topmost = New-Object 'System.Windows.Forms.Form' -Property @{TopMost=$true}
    #$DateForm.ShowDialog($topmost)

    $DateForm.Add_Shown($DateForm.Activate())
    $result = $DateForm.showdialog($topmost)
    #$Calendar.SelectionEnd

    if ($result -eq [Windows.Forms.DialogResult]::OK) {
        $Script:Period = $calendar.SelectionEnd
       
        #Write-Output "Date selected: $($Period.ToShortDateString())"
        Write-Output "Date selected: $(get-date($period) -Format 'MMMM yyyy')"
    }

    if ($result -eq [Windows.Forms.DialogResult]::Cancel) {
        $Script:Period = get-date
    }
}


<# Depricated    
    Function Get-MVReport-old {

        $Script:TempMVReport = "c:\windows\temp\TempMVReport.xlsx"
        remove-item $Script:TempMVReport -Force -Recurse -ErrorAction SilentlyContinue

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36"

    Invoke-WebRequest -UseBasicParsing -Uri "https://api.integrated.cloudbackup.management/statistics-reporting/high-watermark-usage-report?_=0a7402d0c6f82&dateFrom=2022-06-01T00:00:00Z&dateTo=2022-06-15T17:05:00Z&exportOutput=OneXlsxFile&partnerId=$($Script:PartnerId)" `
    -WebSession $session `
    -OutFile $Script:TempMVReport `
    -Headers @{
    "authority"="api.integrated.cloudbackup.management"
    "method"="GET"
    #"path"="/statistics-reporting/high-watermark-usage-report?_=0a7402d0c6f82&dateFrom=2022-06-01T00:00:00Z&dateTo=2022-06-15T17:05:00Z&exportOutput=OneXlsxFile&partnerId=$($Script:PartnerId)"
    "scheme"="https"
    "accept"="application/json, text/plain, */*"
    "accept-encoding"="gzip, deflate, br"
    "accept-language"="en-US,en;q=0.9"
    "authorization"="Bearer $($script:ncvisa)"
    "origin"="https://integrated.cloudbackup.management"
    "referer"="https://integrated.cloudbackup.management/"
    "sec-ch-ua"="`" Not A;Brand`";v=`"99`", `"Chromium`";v=`"102`", `"Google Chrome`";v=`"102`""
    "sec-ch-ua-mobile"="?0"
    "sec-ch-ua-platform"="`"Windows`""
    "sec-fetch-dest"="empty"
    "sec-fetch-mode"="cors"
    #"sec-fetch-site"="same-site"
    }
    }
#>
    Function Get-MVReport {
        Param ([Parameter(Mandatory=$False)][Int]$PartnerId) #end param
        
        $Start = (Get-Date $period -Day 1).AddMonths(0).ToString('yyyy-MM-dd')
        $End = (Get-Date $period -Day 1).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
        Write-Output "  Requesting Maximum Value Report from $start to $end"
        
        $Script:TempMVReport = "c:\windows\temp\TempMVReport.xlsx"
        remove-item $Script:TempMVReport -Force -Recurse -ErrorAction SilentlyContinue
        
        $url2 = "https://api.integrated.cloudbackup.management/statistics-reporting/high-watermark-usage-report?_=0a7402d0c6f82&dateFrom=$($Start)T00:00:00Z&dateTo=$($end)T23%3A59%3A59Z&exportOutput=OneXlsxFile&partnerId=$($Script:PartnerId)"
        
        $method = 'GET'
        
        $params = @{
            Uri         = $url2
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:ncvisa" }
            ContentType = 'application/json; charset=utf-8'
            OutFile     = $Script:TempMVReport 
        }  

        #debug Write-output  "$url2"

        Invoke-RestMethod @params
        
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Host "  Module ImportExcel Already Installed"
        } 
        else {
            try {
                Install-Module -Name ImportExcel -Confirm:$False -Force      ## https://powershell.one/tricks/parsing/excel
            }
            catch [Exception] {
                $_.message 
                exit
            }
        }

        $Script:MVReportxls = Import-Excel -path "$Script:TempMVReport" -asDate "*Date" 

        $Script:MVReport = $Script:MVReportxls | select-object *,@{n='PhysicalServer';e='Null'},@{n='VirtualServer';e='Null'},@{n='Workstation';e='Null'},@{n='Documents';e='Null'}   

        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "0"){$_.RecoveryTesting = $null}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "True"){$_.RecoveryTesting = "1"}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "1"){$_.RecoveryTesting = "RecoveryTesting"}}
        $Script:MVReport | foreach-object { if ($_.RecoveryTesting -eq "2"){$_.RecoveryTesting = "StandbyImage"}}
        $Script:MVReport | foreach-object { if ($_.O365Users -eq "0"){$_.O365Users = $null}}

    } 

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
        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [Int]$DeviceResult.AccountId;
                                                                    PartnerID      = [string]$DeviceResult.PartnerId;
                                                                    DeviceName     = $DeviceResult.Settings.AN -join '' ;
                                                                    ComputerName   = $DeviceResult.Settings.MN -join '' ;
                                                                    DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                                                                    PartnerName    = $DeviceResult.Settings.PF -join '' ;
                                                                    #Reference      = $DeviceResult.Settings.AR -join '' ;
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

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
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
        $Script:DeviceHistoryDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [Int]$DeviceHistory.AccountId;
                                                                    PartnerID      = [string]$DeviceHistory.PartnerId;
                                                                    DeviceName     = $DeviceHistory.Settings.AN -join '' ;
                                                                    ComputerName   = $DeviceHistory.Settings.MN -join '' ;
                                                                    DeviceAlias    = $DeviceHistory.Settings.AL -join '' ;
                                                                    PartnerName    = $DeviceHistory.Settings.AR -join '' ;
                                                                    #Reference      = $DeviceHistory.Settings.PF -join '' ;
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
                                                                    Physicality    = $DeviceHistory.Settings.I81 -join'' ;
                                                                    ProfileID      = $DeviceHistory.Settings.OI -join '' }
        }

    } ## EnumerateAccountStatistics API Call

    Function Join-Reports {

        if (Get-Module -ListAvailable -Name Join-Object) {
            Write-Host "  Module Join-Object Already Installed"
        } 
        else {
            try {
                Install-Module -Name Join-Object -Confirm:$False -Force      ## link
            }
            catch [Exception] {
                $_.message 
                exit
            }
        }

        $Script:MVPlus = Join-Object -left $Script:MVReport -right $Script:SelectedDevices -LeftJoinProperty DeviceId -RightJoinProperty AccountID -Prefix 'Now_' -Type AllInBoth
        $Script:MVPlus | Where-Object {$_.DeviceDeletionDate} | foreach-object { Send-GetDeviceHistory $_.deviceid $_.DeviceDeletionDate; $_.Now_Physicality = $Script:DeviceHistoryDetail.physicality; $_.Now_Product = $Script:DeviceHistoryDetail.Product }
    }

    #endregion ----- Functions ----

            Send-APICredentials
            Send-APICredentialsNC

            Write-Output $Script:strLineSeparator
            Write-Output ""
            
            $Script:Period
            $Last

            if (($Script:period -eq $null) -and ($last -eq $null)) {Get-Period}
            if (($Script:period -eq $null) -and ($last -eq 0)) { $Script:period = (get-date) }
            if (($Script:period -eq $null) -and ($last -ge 1)) { $Script:period = ((get-date).addmonths($last/-1)) }

<#
        $Start = (Get-Date $period -Day 1).AddMonths(0).ToString('yyyy-MM-dd')
        $End = (Get-Date $period -Day 1).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
        #>

            Send-GetDevices $partnerId
            
            $Script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,AccountID,DeviceName,ComputerName,Creation,TimeStamp,LastSuccess,ProductId,Product,ProfileId,Profile,DataSources,SelectedGB,UsedGB,Location,OS,OSType,Physicality

            Get-MVReport $partnerId
            join-reports
            
            if($null -eq $Script:SelectedDevices) {
                # Cancel was pressed
                # Run cancel script
                Write-Output    $Script:strLineSeparator
                Write-Output    "  No Devices Selected"
                Exit
            }else{
                # OK was pressed, $Selection contains what was chosen
                # Run OK script
                
            }
        
            $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_MaxValuePlus_DPP_$($period.ToString('yyyy-MM'))_Statistics_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"

            $Script:MVPlus | foreach-object { if ($_.SelectedsizeGB -gt 0) {$_.SelectedsizeGB = [math]::round($_.SelectedsizeGB,2)}}
            $Script:MVPlus | foreach-object { if ($_.UsedStorageGB -gt 0) {$_.UsedStorageGB = [math]::round($_.UsedStorageGB,2)}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Physical")) {$_.PhysicalServer = "1"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Server") -and ($_.Now_Physicality -eq "Virtual")) {$_.VirtualServer = "1"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -eq "Documents")) {$_.Documents = "1"}}
            $Script:MVPlus | foreach-object { if (($_.OsType -eq "Workstation") -and ($_.Now_Product -ne "Documents")) {$_.Workstation = "1"}}
            $Script:MVPlus | foreach-object { if (($_.Now_Physicality -eq "Undefined") -and ($_.O365Users)) {$_.Now_Physicality = "Cloud"; $_.OSType = "M365";  $_.ComputerName = "* M365 - $($_.DeviceName)"}}
            
            # | Where-object {$_.CustomerState -eq "InProduction"} | 
            $Script:MVPlus | Where-object {$_.CustomerState -eq "InProduction"} | Select-object Parent3Name,Parent2Name,Parent1Name,CustomerName,ComputerName,DeviceName,DeviceId,Now_Product,OsType,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGB,UsedStorageGB | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8
        
            $Script:xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")

            $ConditionalFormat =$(New-ConditionalText -ConditionalType DuplicateValues -Range 'E:E' -BackgroundColor CornflowerBlue -ConditionalTextColor Black)
        
            $Script:MVPlus | Where-object {$_.CustomerState -eq "InTrial"} | Select-object Parent3Name,Parent2Name,Parent1Name,CustomerName,ComputerName,DeviceName,DeviceId,Now_Product,OsType,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGB,UsedStorageGB | Export-Excel -Path "$xlsoutputfile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName TrialUsage -FreezeTopRowFirstColumn -WorksheetName "TRIAL $(get-date $period -UFormat `"%b-%Y`")" -tablestyle Medium6 -BoldTopRow

            $Script:MVPlus | Where-object {$_.CustomerState -eq "InProduction"} | Select-object Parent3Name,Parent2Name,Parent1Name,CustomerName,ComputerName,DeviceName,DeviceId,Now_Product,OsType,Now_Physicality,PhysicalServer,VirtualServer,Workstation,Documents,RecoveryTesting,O365Users,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGB,UsedStorageGB | Export-Excel -Path "$xlsoutputfile" -ConditionalFormat $ConditionalFormat -AutoFilter -AutoSize -TableName ProductionUsage -FreezeTopRowFirstColumn -WorksheetName (get-date $period -UFormat "%b-%Y") -tablestyle Medium6 -BoldTopRow -IncludePivotTable -PivotRows Parent3Name,Parent2Name,Parent1Name,CustomerName -PivotDataToColumn -PivotData @{PhysicalServer='sum';VirtualServer='sum';Workstation='sum';Documents='sum';RecoveryTesting='count';O365Users='sum'}


                


   

    ## Launch CSV or XLS (if Excel is installed)  (Required -Launch Parameter)
        
    if ($Launch) {
        If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
            Start-Process "$xlsoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
            }else{
            Start-Process "$csvoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening CSV file"
            Write-output $Script:strLineSeparator            
            }
        }
        
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $csvoutputfile"
    Write-Output "  XLS Path = $xlsoutputfile"
    Write-Output ""
    
    Start-Sleep -seconds 10





#$Period = "2022-03-01" 
#Get-MVReport
#$MVReport | Out-GridView