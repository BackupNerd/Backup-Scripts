<# ----- About: ----
    # N-able Backup Get Deleted Devices  
    # Revision v01 - 2021-04-02
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
    # For use with the Standalone edition of N-able Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Get Maximum Value Report for Period
    # Export to XLS/CSV
    # Optionally send email via SMTP
    # Optionally launch XLS/CSV
    #
    # Use the -MaxValue switch parameter to download MaxValue report
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -Export switch parameter to export statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/API-column-codes.htm 


# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [switch]$MaxValue = $true,                    ## Lookup Monthly Max Value Report
        [Parameter(Mandatory=$False)] [switch]$AllPartners = $true,                 ## Skip partner selection
        [Parameter(Mandatory=$False)] [datetime]$Period = (Get-Date),               ## Lookup Date yyyy-MM or MM-yyyy
        [Parameter(Mandatory=$False)] [switch]$Export = $true,                      ## Generate CSV / XLS Output Files
        [Parameter(Mandatory=$False)] [switch]$Launch,                              ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                     ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",                ## Export Path
        [Parameter(Mandatory=$False)] [switch]$SendMail,                            ## Send XLS output via Email
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                     ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Deleted Devices"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-output "  $ConsoleTitle`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n  $Syntax"

    If ($SendMail) {$Export = $true}    
    If ($Launch) {$Export = $true}   
    Write-output "  Current Parameters:"
    Write-output "  -MaxValue    = $MaxValue"
    Write-output "  -AllPartners = $AllPartners"
    Write-output "  -Export      = $Export"
    Write-output "  -Launch      = $Launch"
    Write-output "  -SendMail    = $SendMail"
    Write-output "  -ExportPath  = $ExportPath"
    Write-output "  -Delimiter   = $Delimiter"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
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

        Write-Output "  Enter Exact, Case Sensitive Partner Name for SolarWinds Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for SolarWinds Backup.Management API'
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
    $Script:UserId = $authenticate.result.result.id
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

#endregion ----- Authentication ----

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

        $RestrictedPartnerLevel = @("Root","SubRoot")

        if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$Script:Level = $Partner.result.result.Level
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
                        Break
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

    Function GetMVReport {
        Param ([Parameter(Mandatory=$False)][Int]$PartnerId) #end param
                
        $Script:end = (get-date).addDays(1).ToString('yyyy-MM-dd')
        $Script:start = (get-date $end).addDays(-15).ToString('yyyy-MM-dd')
        $Script:TempMVReport = "c:\windows\temp\TempMVReport.xlsx"
        
        Write-Output "  Requesting Deleted Devices from the Maximum Value Report for the period between $start and $end"

        $Script:url2 = "https://api.backup.management/statistics-reporting/high-watermark-usage-report?_=6e8d1e0fce68d&dateFrom=$($Start)T00%3A00%3A00Z&dateTo=$($end)T23%3A59%3A59Z&exportOutput=OneXlsxFile&partnerId=$($PartnerId)"
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Cookie", "__cfduid=d7cfe7579ba7716fe73703636ea50b1251593338423; visa=$Script:visa")

        Write-output  "  $url2"

        Invoke-RestMethod -Uri $url2 `
        -Method GET `
        -Headers $headers `
        -ContentType 'application/json' `
        -WebSession $websession `
        -OutFile $Script:TempMVReport 
    
        Install-module -name ImportExcel                                                     ## https://powershell.one/tricks/parsing/excel

        $Script:MVReportxls = Import-Excel -path "$Script:TempMVReport" -asDate "*Date" 
        $Script:MVReport = $Script:MVReportxls | select-object * 
        Remove-Item $Script:TempMVReport
    } 

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $switch = $PSCmdlet.ParameterSetName

    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Send-GetPartnerInfo $Script:cred0

    if ($AllPartners) {}else{Send-EnumeratePartners}
        
    if ($MaxValue) {
        GetMVReport $partnerId
    }

    if (($Export) -and ($MaxValue))  {

        $csvoutputfile = "$ExportPath\$($CurrentDate)_RecentlyDeletedDevices_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
        $Deleted = $Script:MVReport | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object Parent1Name,Parent1Id,CustomerName,CustomerId,ComputerName,DeviceName,DeviceId,OsType,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGb,UsedStorageGb,O365Users
        $Script:MVReport | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object Parent1Name,Parent1Id,CustomerName,CustomerId,ComputerName,DeviceName,DeviceId,OsType,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGb,UsedStorageGb,O365Users | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8
        $Script:MVReport | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object Parent1Name,Parent1Id,CustomerName,CustomerId,ComputerName,DeviceName,DeviceId,OsType,CustomerState,CreationDate,ProductionDate,DeviceDeletionDate,SelectedSizeGb,UsedStorageGb,O365Users | Export-Excel -path "$xlsoutputfile" -AutoFilter -AutoSize
    }

    if ($sendmail) {
        ##############################################################################
        $From = "nable.backup.nerd@gmail.com"
        $To = @("eric.harless@solarwinds.com")
        $Cc = @("eric.harless@solarwinds.com")
        $Attachment = "$xlsoutputfile"
        $Subject = "$($CurrentDate) | $($deleted.count) Recently Deleted Backup Devices for Partner $($Partnername) | Id $($PartnerId)"
        $Body = "<b>$($deleted.count)</b> recently deleted backup devices found for partner <b>$PartnerName</b> between <b>$($start)</b> and <b>$($end)</b>.<br /><br />"
        $Body += "You should contact backup technical support immediately if any listed devices are believed to have been accidentially, unintentionally or maliciously deleted.<br /><br />"
        $Body += $deleted | sort-object DeviceDeletionDate | Select-object Parent1Name,CustomerName,DeviceName,CreationDate,DeviceDeletiondate | ConvertTo-Html | Out-String
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = "587"
        Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject `
        -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
        -Credential (Get-Credential) -Attachments $Attachment -BodyAsHtml
        ##############################################################################

    } 

    Write-output $Script:strLineSeparator

    ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
        
    if ($Launch) {
        If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
            Start-Process -filepath $xlsoutputfile
            Write-output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
            }else{
                Start-Process -filepath $csvoutputfile
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