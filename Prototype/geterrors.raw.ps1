<# ----- About: ----
    # Bulk Get N-able Backup Device Errors
    # Revision v31 - 2023-10-20
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
    # For use with the Standalone edition of N-able Backup
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Get last session errors for selected devices
    # Optionally post/ clear error message to/ from a custom column in the https://backup.management console
    # Optionally export to XLS/CSV
    #
    # Use the -Days ## (default=7) parameter to define the age of devices with errors to query
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -GridView switch parameter to display output via Powershell Out-Gridview
    # Use the -Export switch parameter to export statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify alternate XLS/CSV file path
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -CustomColumn (default=$True) switch parameter to post last error to https://backup.management console column AA2045
    #   Note: Partner must add custom column AA2045 or specified custom column to their https://backup.management console to view last error message.
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/console/custom-columns.htm
# -----------------------------------------------------------#>  ## Behavior


#region ----- Environment, Variables, Names and Paths ----
    #Clear-Host
    #[int]$Days = 7
    #[int]$DeviceCount = 5000
    [String]$Action = "ActiveErrors"
    [switch]$CustomColumn = $true
    [string]$ColumnCode = "AA2045"
    [switch]$Export = $true
    #[string]$ExportPath = "C:\ProgramData\NerdScripts\GetCoveErrors"
    #[string]$Delimiter = ","
    #[string]$AdminLoginPartnerName = ""
    #[string]$AdminLoginUserName = ""
    #[string]$AdminPassword = ""

    #$scriptpath = $MyInvocation.MyCommand.Path
    #$dir = Split-Path $scriptpath
    #Write-output "  $ConsoleTitle`n`n$ScriptPath"
    #$Syntax = Get-Command $PSCommandPath -Syntax 
    #Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    #Push-Location $dir
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    $ExportPath = Join-path -path $ExportPath -childpath "DeviceErrors_$shortdate"
    mkdir -force -path $ExportPath | Out-Null
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'

    Write-output "  Current Parameters:"
    Write-output "  -Days          = $Days"
    Write-output "  -DeviceCount   = $DeviceCount"
    Write-output "  -Export        = $Export"
    Write-output "  -ExportPath    = $ExportPath"
    Write-output "  -Delimiter     = $Delimiter"
    Write-output "  -CustomColumn  = $CustomColumn"
    Write-output "  -ColumnCode    = $ColumnCode"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function Authenticate {

<#

    if ($AdminLoginPartnerName -eq $null) {
        Write-Output $script:strLineSeparator
        Write-Output "  Enter Your N-able | Cove https:\\backup.management Login Credentials"
        Write-Output $script:strLineSeparator
        
        $AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    }
    if ($AdminLoginUserName -eq $null) {$AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able | Cove Backup.Management API"}

    if ($PlainTextAdminPassword -eq $null) {
        $AdminPassword = Read-Host -AsSecureString "  Enter Password for N-able | Cove Backup.Management API"
        # (Convert SecureString Password to plain text)
        $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword))
        }
#>

    # (Show credentials for Debugging)
    Write-Output "  Logging on with the following Credentials`n"
    Write-Output "  PartnerName:  $AdminLoginPartnerName"
    Write-Output "  UserName:     $AdminLoginUserName"
    Write-Output "  Password:     It's secure..."

# (Create authentication JSON object using ConvertTo-JSON)
    $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
    Add-Member -PassThru NoteProperty method 'Login' |
    Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$AdminLoginUserName;password=$AdminPassword}|
    Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json

# (Call the JSON function with URL and authentication object)
    $script:session = CallJSON $urlJSON $objAuth
    #Start-Sleep -Milliseconds 100
    
# (Variable to hold current visa and reused in following routines)
    $script:visa = $session.visa
    $script:PartnerId = [int]$session.result.result.PartnerId
        
# (Get Result Status of Authentication)
    $AuthenticationErrorCode = $Session.error.code
    $AuthenticationErrorMsg = $Session.error.message

# (Check if ErrorCode has a value)
    If ($AuthenticationErrorCode) {
        Write-Output "Authentication Error Code:  $AuthenticationErrorCode"
        Write-Output "Authentication Error Message:  $AuthenticationErrorMsg"
        #Break Script
    }# (Exit Script if there is a problem)
    Else {

    } # (No error)

# (Print Visa to screen)
    #Write-Output $script:strLineSeparator
    #Write-Output "Current Visa is: $script:visa"
    #Write-Output $script:strLineSeparator

## Authenticate Routine
    }  ## Use Backup.Management credentials to Authenticate



    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Authenticate
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

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function CallJSON($url,$object) {

        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
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

    Function Send-GetPartnerInfo ($PartnerName) { 

        $RestrictedPartnerLevel = @("Root","SubRoot")
            
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'GetPartnerInfo'
        $data.params = @{}
        $data.params.name = [String]$AdminLoginPartnerName

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            $Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Partner = $webrequest | convertfrom-json

        if ($RestrictedPartnerLevel -notcontains $Partner.result.result.Level) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-output "  $Level - $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Output "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
            }
            write-output "  $($partner.error.message)"

    } ## get PartnerID and Partner Level

    Function GetDeviceErrors($DeviceId) {
    
        $url2 = "https://backup.management/web/accounts/properties/api/errors/recent?accounts.SelectedAccount.Id=$DeviceId"
        $method = 'GET'

            $params = @{
            Uri         = $url2
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:response = Invoke-RestMethod @params 

        if ($Script:response.error) {Send-APICredentialsCookie}

        If ($response.HasItemsWithCount -eq $False) {
        
        Write-output "  DeviceId# $DeviceId - Skipping, initial backup in progress"
        }else{
            Write-output "  DeviceId# $DeviceId - Requesting errors "   
            $response = $response.replace('[ESCAPE[','') 
            $response = $response.replace("]]","") 
            $response = $response.replace("\","\\")
            $response = $response.replace("&quot;","")  
            $response = $response.replace(",`n        `n    ]","`n     ]") | ConvertFrom-Json
        
            #$Script:response = $response.collection | select-object time,count,file,message,groupID,DeviceID | Format-Table
            #$Script:response.deviceid = $deviceid
            #$Script:response
            #$response.collection | select-object time,count,file,message | out-gridview
        }
        $Script:cleanresponse = $response.collection
        } 
    
    Function Send-UpdateCustomColumn($DeviceId,$ColumnId,$Message) {

        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Content-Type", "application/json")
        $headers.Add("Cookie", "__cfduid=d110201d75658c43f9730368d03320d0f1601993342")
        $headers.Add("Authorization", "Bearer $script:visa")
        
        $body = "{
        `n      `"jsonrpc`":`"2.0`",
        `n      `"id`":`"jsonrpc`",
        `n      `"method`":`"UpdateAccountCustomColumnValues`",
        `n      `"params`":{
        `n      `"accountId`": $DeviceId,
        `n      `"values`": [[$ColumnId,`"$Message`"]]
        `n      }
        `n  }
        `n"
        
        $script:updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body
        $script:updateCC.error.message
        }
    
    Function Send-GetErrorDevices {

        Switch ($Action) {
            'LastErrors'{
                $DeviceFilter = "(TS > $days.days().ago())"
            } ## Devices active in the last $Days
            'ActiveErrors'{
                $DeviceFilter = "(T7 >= 1) AND (TS > $days.days().ago())"
            } ## Devices active in the last $Days with a Current Error
            'Custom'{
                $DeviceFilter = "OT == 1 AND ( ANY =~ '*macos*' OR ANY =~ '*os X*')  AND (T7 >= 1) AND ( TS > 7.days().ago())"
            } ## Mac Workstations with Active

        }

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $DeviceFilter
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.OrderBy = "TS DESC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        $Script:ErrorDevices = $webrequest | convertfrom-json
    
        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-Output $Script:strLineSeparator
            Write-output "  Searching  $PartnerName "
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Output "PartnerName Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:ErrorDeviceDetail = @()

        if ($ErrorDevices.result.result.count -eq 0) {
            Write-Output "  No Errors Found for the Last $days Days`n  Exiting Script"
            #Start-Sleep -seconds 10
            #Break
        }
        
        Write-Output "  Requesting error details for $($ErrorDevices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $ErrorDevices.result.result ) {
            Get-VisaTime
            $Script:ErrorDeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [String]$DeviceResult.AccountId;
                                                                        PartnerID      = [string]$DeviceResult.PartnerId;
                                                                        DeviceName     = $DeviceResult.Settings.AN -join '' ;
                                                                        DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                                                                        PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                        Reference      = $DeviceResult.Settings.PF -join '' ;
                                                                        DataSources    = $DeviceResult.Settings.AP -join '' ;                                                                
                                                                        Account        = $DeviceResult.Settings.AU -join '' ;
                                                                        Location       = $DeviceResult.Settings.LN -join '' ;
                                                                        Notes          = $DeviceResult.Settings.AA843 -join '' ;
                                                                        TempInfo       = $DeviceResult.Settings.AA77 -join '' ;
                                                                        Errors         = $DeviceResult.Settings.T7 -join '' ;
                                                                        Errors_FS      = $DeviceResult.Settings.F7 -join '' ;                                                                    
                                                                        Product        = $DeviceResult.Settings.PN -join '' ;
                                                                        ProductID      = $DeviceResult.Settings.PD -join '' ;
                                                                        Profile        = $DeviceResult.Settings.OP -join '' ;
                                                                        OS             = $DeviceResult.Settings.OS -join '' ;                                                                
                                                                        ProfileID      = $DeviceResult.Settings.OI -join '' 
                                                                    }
                                                                

        GetDeviceErrors $DeviceResult.AccountId
                
        if (($CustomColumn) -and ($cleanresponse.message -like "*restarted*")) {
            $RestartedMSG = $cleanresponse.message | Where-Object {$_ -like "*restarted*"}
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) ($RestartedMSG + " " + $cleanresponse[-1].time)
            #$RestartedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*Operation Aborted*")) {
            $AbortedMSG = $cleanresponse.message | Where-Object {$_ -like "*Operation Aborted*"}
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) ($AbortedMSG + " " + $cleanresponse[-1].time)
            #$AbortedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*Operation not permitted*")) {
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) ("Operation not permitted - Enable Full Disk Access " + $cleanresponse[-1].time)
            #$AbortedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*significant time difference*")) {
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) ("The machine seems to have a significant time difference with the Cloud. " + $cleanresponse[-1].time)
            #$AbortedMSG
        }elseif ($CustomColumn) {
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) ($cleanresponse[-1].message + " " + $cleanresponse[-1].time)  
            #$cleanresponse[-1].message 

        }     
        
            ForEach ( $ErrorResult in $Script:cleanresponse ) {
  
            $Script:ResponseDetail += New-Object -TypeName PSObject -Property @{ DeviceId      = [String]$DeviceResult.AccountId;
                                                                     DeviceName     = $DeviceResult.Settings.AN -join '';
                                                                     PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                     Time           = [datetime]$ErrorResult.time;
                                                                     Count          = $ErrorResult.count;
                                                                     File           = $ErrorResult.file;
                                                                     Message        = $ErrorResult.message;
                                                                     GroupID        = $ErrorResult.groupid }
                                                                    }
        }   
    }

    Function Send-ClearErrorDevices {

        $filter1 = "($columncode AND T7 == 0) OR ($columncode AND TS < $days.days().ago())"

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        $Script:ZeroErrorDevices = $webrequest | convertfrom-json

        ForEach ( $DeviceResult in $ZeroErrorDevices.result.result ) {
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $null
            Write-output "  DeviceId# $($DeviceResult.AccountId) - Clearing Custom Columns for Resolved Errors or Errors older than $days Days"
            Get-VisaTime
        }  
    }
#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $Script:ResponseDetail = @()

    Authenticate

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Send-GetPartnerInfo $AdminLoginPartnerName

    Send-ClearErrorDevices
        
    Send-GetErrorDevices

    $filterDate = (Get-Date).AddDays(-$days)

    $ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | sort-object partnername,deviceid,groupid | format-table   

    if ($Script:Export) {
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_DeviceErrors_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $script:ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8 -append

        Write-Output "  CSV Path = $Script:csvoutputfile"
    }

