# ----- About: ----
    # Cove Data Protection | Export Report  
    # Revision v23.12 - 2023-12-08
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
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
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----

    # Export Reports in Ascii, Csv, Html, OfficeOpenXml, Pdf or Xml formats
    #
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Optionally specify report month
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select devices
    # Generate Report for Period
    #
    #   Use the -Last parameter to specify the prior month count ( '0' current month, '1' month prior, etc)
    #       Or Use the -Period parameter to define report period (yyyy-MM or MM-yyyy)
    #       Or leave blank and be prompted for a month via pop-up GUI calendar
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path. Default is the execution path of the script.
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console/export.htm


# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)][ValidateRange(0,84)] [int]$Last = 1,                                  ## Count back # Last Months i.e. 0 current, 1 Prior Month
        [Parameter(Mandatory=$False)][datetime]$Period,                                                     ## Lookup Date yyyy-MM or MM-yyyy
        [Parameter(Mandatory=$False)][switch]$AllPartners,                                                  ## Skip GUI partner selection
        [Parameter(Mandatory=$False)][int]$DeviceCount = 5000,                                              ## Change Maximum Number of current devices results to return
        [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",                                 ## Export Path
        [Parameter(Mandatory=$False)]
            [ValidateSet("Ascii","Csv","Html","OfficeOpenXml","Pdf","Xml")]
            [String]$ReportFormat="Pdf",                                                                    ## Specify the report output format to use
        [Parameter(Mandatory=$False)]
            [ValidateSet("SingleReportPerPartner","AllReportsTogether")
            ][String]$Granularity="AllReportsTogether",                                                     ## 'AllReportsTogether' is currently the only supported option
        [Parameter(Mandatory=$False)]
            [ValidateSet("HumanReadable","InBytes","InGigabytes")]
            [String]$SizeFormat = "HumanReadable",                                                          ## Specify the number size format to use
        [Parameter(Mandatory=$False)][Switch]$ClearCredentials                                              ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----

Clear-Host
#Requires -Version 5.1
$ConsoleTitle = "Cove Data Protection | Export Report to $ReportFormat"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

Write-output "  $ConsoleTitle`n`n$ScriptPath"
$Syntax = Get-Command $PSCommandPath -Syntax 
Write-Output "  Script Parameter Syntax:`n`n$Syntax"

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Push-Location $dir
$CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
$ShortDate = Get-Date -format "yyyy-MM-dd"

if ($ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "Cove_Reports_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "Cove_Reports_$shortdate"}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------"
If ($exportpath) {mkdir -force -path $ExportPath | Out-Null}
$urlJSON = 'https://api.backup.management/jsonapi'

$culture = get-culture; $delimiter = $culture.TextInfo.ListSeparator

Write-output "  Current Parameters:"
Write-output "  -Period         = $Period"
Write-output "  -Month Prior    = $Last"
Write-output "  -AllPartners    = $AllPartners"
Write-output "  -DeviceCount    = $DeviceCount"
Write-output "  -ReportFormat   = $ReportFormat"
Write-output "  -SizeFormat     = $SizeFormat"
Write-output "  -Granularity    = $Granularity"
Write-output "  -Delimiter      = $Delimiter"
Write-output "  -ExportPath     = $ExportPath"

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
        
        }else{ 
            Write-Output $Script:strLineSeparator 
            Write-Output "  Getting Backup API Credentials" 
        
            if (Test-Path $APIcredfile) {
                Write-Output    $Script:strLineSeparator        
                Write-Output  "  Backup API Credential File Present"
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
    -TimeoutSec 30 `
    -SessionVariable Script:websession `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:Authenticate = $webrequest | convertfrom-json

#Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

if ($authenticate.visa) { 

    $Script:visa = $authenticate.visa
    $Script:UserId = $authenticate.result.result.id
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator 
       
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Visa-Check {
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
}

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

Function Send-GetPartnerInfo ([string]$PartnerName) { 
                
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

        [String]$Script:Uid = $Script:Partner.result.result.Uid
        $Script:PartnerId = $script:Partner.result.result.Id
        [String]$Script:Level = $Script:Partner.result.result.Level
        [String]$Script:PartnerName = $Script:Partner.result.result.Name

} ## get PartnerID and Partner Level    

Function Send-GetPartnerInfoByID ([int]$PartnerId) { 
                
    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfoById'
    $data.params = @{}
    $data.params.partnerId = [string]$PartnerId

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json



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
}

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
                $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                Write-Output    $Script:strLineSeparator
                Write-Output    "  All Partners Selected"
            }else{
                $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single
        
                if($null -eq $Selection) {
                    # Cancel was pressed
                    # Run cancel script
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  No Partners Selected"
                    Exit
                }
                else {
                    # OK was pressed, $Selection contains what was chosen
                    # Run OK script
                    [int]$Script:PartnerId = $Script:Selection.Id
                    [String]$Script:PartnerName = $Script:Selection.Name
                }
            }

    }
    
}  ## EnumeratePartners API Call

Function Send-GetReport ($Granularity,$ReportFormat,$SizeFormat,$partnerId,$timestamp) {
    #$now = get-date

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $visa
    $data.method = 'EnumerateReports'
    $data.params = @{}
    $data.params.granularity = "AllReportsTogether"             ## $Granularity
    $data.params.timestamp = Convert-DateTimeToUnixTime $end
    $data.params.sizeFormat = $SizeFormat
    $data.params.reportFormat = $ReportFormat
    $data.params.reportType = "Aggregated"
    $data.params.SelectionMode = "Merged"
    #$data.params.ArchiveType = "Zip"
    $data.params.query = @{}
    $data.params.query.Columns = @("AR","TB","MN","AN","AU","LN","OT","OS","AP","PN","PD","OP","OI","CD","TS","TL","T3","US","I80","I81")
    $data.params.query.Filter = ""
    $data.params.query.OrderBy = "CD DESC"
    $data.params.query.PartnerId = $partnerId
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = $DeviceCount
    $data.params.query.Totals = @("COUNT(AT==1)")

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
        
    }  

    $Script:ReportResponse = Invoke-RestMethod @params 
    $extension = $Script:ReportResponse.result.result.format[1]

    Switch ($ReportFormat) {
        'Pdf' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).$($extension)"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($Script:ReportResponse.result.result.content)))
            
            #[System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($Script:ReportResponse.result.result.content)) | Out-File "c:\data\test2.pdf" ## Alternate Method
        }
        'Html' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).$($extension)"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($ReportResponse.result.result.content)))
          }
        'Csv' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).$($extension)"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($ReportResponse.result.result.content)))
        }
        'OfficeOpenXml' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).Xlsx"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($ReportResponse.result.result.content)))
        }
        'Xml' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).$($extension)"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($ReportResponse.result.result.content)))
        }
        'Ascii' {
            if ($Granularity -eq "AllReportsTogether") { $Script:Reportpath = "$ExportPath\$($CurrentDate)_Cove_Report_$($period.year)-$($period.month)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).txt"}

            [IO.File]::WriteAllBytes("$reportpath",([Convert]::FromBase64String($ReportResponse.result.result.content)))
        }
    }
 
} ## EnumerateAccountStatistics API Call

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

$RestrictedPartnerLevel = @("Root","SubRoot","Distributor")
Send-APICredentialsCookie
Write-Output $Script:strLineSeparator
Write-Output ""

Send-GetPartnerInfo $Script:cred0

do {
        #Write-Output $Script:strLineSeparator
        #Write-output "  $Script:PartnerName - $partnerId - $Uid"
        #Write-Output $Script:strLineSeparator
                            
        if ($Script:Partner.result.result.Level -in $RestrictedPartnerLevel) {
            Write-Host "  Lookup for $($Script:Partner.result.result.Level) Partner Level Not Allowed"
            }

        if ($Script:Partner.error) {
            Write-Host "  $($Script:Partner.error.message)"
            }

        Write-Output $Script:strLineSeparator
        $Script:Partnernamelookup = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        
        Send-GetPartnerInfo $partnernamelookup
        
        #$Script:Uid            = $Script:Partner.result.result.Uid
        #$Script:PartnerId      = $script:Partner.result.result.Id
        #$Script:Level          = $Script:Partner.result.result.Level
        #$Script:PartnerName    = $Script:Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-output "  $Script:PartnerName - $Script:partnerId - $Script:Uid"
        Write-Output $Script:strLineSeparator

} while (
            ($Script:Partner.result.result.Level -in $RestrictedPartnerLevel) -or ($Script:Partner.error)
        )

if (($Script:period -eq $null) -and ($last -eq $null)) {Get-Period}
if (($Script:period -eq $null) -and ($last -eq 0)) { $Script:period = (get-date) }
if (($Script:period -eq $null) -and ($last -ge 1)) { $Script:period = ((get-date).addmonths($last/-1)) }

$Start = (Get-Date $period -Day 1).AddMonths(0).ToString('yyyy-MM-dd')
$End = (Get-Date $period -Day 1).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')

if ($AllPartners) {}else{Send-EnumeratePartners}

Send-GetReport $Granularity $ReportFormat $SizeFormat $PartnerId
        
Write-output $Script:strLineSeparator
Write-Output "  Output Path = $reportpath"

Start-Sleep -seconds 5