<# ----- About: ----
    # Bulk Get N-able Backup Device Errors
    # Revision v27 - 2020-07-22
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
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -CustomColumn (default=$True) switch parameter to post last error to https://backup.management console column AA2045
    #   Note: Partner must add custom column AA2045 or specified custom comlun to their https://backup.management console to view last error message.
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/console/custom-columns.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [int]$Days = 7,                                           ## Number of days to search for Errors
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 5000,                                 ## Change Maximum Number of devices results to return
        [Parameter(Mandatory=$False)] [switch]$GridView,                                        ## Display Output via Powershell Out-Gridview
        [Parameter(Mandatory=$False)] [switch]$Export,                                          ## Generate CSV / XLS Output Files
        [Parameter(Mandatory=$False)] [switch]$Launch,                                          ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                                 ## specify ',' or ';' Delimiter for XLS & CSV file
        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",                    ## Export Path
        [Parameter(Mandatory=$False)] [switch]$CustomColumn = $True,                            ## Update Backup.Manangement Custom Column
        [Parameter(Mandatory=$False)] [string]$ColumnCode = "AA2045",                           ## Column Code ShortID to Populate / Clear
        [Parameter(Mandatory=$False)] [string]$TaskPath = "$env:userprofile\BackupNerdScripts", ## Path to Store/Invoke Scheduled Backup Nerd Script Credentials and Tasks
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                 ## Remove Stored API Credentials at start of script
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Bulk Get Device Errors for the Last $days Days"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $dir = Split-Path $scriptpath
    Push-Location $dir
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    if ($launch) {[switch]$export = $true}
    if ($export -and $ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "DeviceErrors_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "DeviceErrors_$shortdate"}
    If ($export) {mkdir -force -path $ExportPath | Out-Null}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'

    Write-output "  Current Parameters:"
    Write-output "  -Days          = $Days"
    Write-output "  -DeviceCount   = $DeviceCount"
    Write-output "  -Export        = $Export"
    Write-output "  -ExportPath    = $ExportPath"
    Write-output "  -Delimiter     = $Delimiter"
    Write-output "  -Launch        = $Launch"
    Write-output "  -Gridview      = $GridView"
    Write-output "  -CustomColumn  = $CustomColumn"
    Write-output "  -ColumnCode    = $ColumnCode"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
        Write-Output $Script:strLineSeparator     
        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($partnerName.length -eq 0)
        $PartnerName | out-file $APIcredfile

        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
        
        Start-Sleep -milliseconds 300

        Send-APICredentialsCookie  ## Attempt API Authentication

    }  ## Set API credentials if not present

    Function Get-APICredentials {

        #$Script:True_path = "C:\ProgramData\MXB\"
        $Script:APIcredfile = join-path -Path $Taskpath -ChildPath "$env:computername API_Credentials.Secure.txt"
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
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate

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
} ## Save as output XLS Routine

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function CallJSON($url,$object) {

        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
        $web = [System.Net.WebRequest]::Create($url)
        $web.Method = “POST”
        $web.ContentLength = $bytes.Length
        $web.ContentType = “application/json”
        $stream = $web.GetRequestStream()
        $stream.Write($bytes,0,$bytes.Length)
        $stream.close()
        $reader = New-Object System.IO.Streamreader -ArgumentList $web.GetResponse().GetResponseStream()
        return $reader.ReadToEnd()| ConvertFrom-Json
        $reader.Close()
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
            $Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Partner = $webrequest | convertfrom-json

        if (($Partner.result.result.Level -ne "Root") -and ($Partner.result.result.Level -ne "Sub-root")) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
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

        $filter1 = "$columncode OR ((T7 >= 1) AND TS > $days.days().ago())"

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
        $data.params.query.OrderBy = "TS ASC"
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
            Write-Host "PartnerName Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:ErrorDeviceDetail = @()

        if ($ErrorDevices.result.result.count -eq 0) {
            Write-Output "  No Errors Found for the Last $days Days`n  Exiting Script"
            Start-Sleep -seconds 10
            Break
        }
        
        Write-Output "  Requesting error details for $($ErrorDevices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $ErrorDevices.result.result ) {

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
                                                                    ProfileID      = $DeviceResult.Settings.OI -join '' }

        GetDeviceErrors $DeviceResult.AccountId
                
        if (($CustomColumn) -and ($cleanresponse.message -like "*restarted*")) {
            $RestartedMSG = $cleanresponse.message | Where-Object {$_ -like "*restarted*"}
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $RestartedMSG
            #$RestartedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*Operation Aborted*")) {
            $AbortedMSG = $cleanresponse.message | Where-Object {$_ -like "*Operation Aborted*"}
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $AbortedMSG
            #$AbortedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*Operation Aborted*")) {
            $AbortedMSG = $cleanresponse.message | Where-Object {$_ -like "*Operation Aborted*"}
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $AbortedMSG
            #$AbortedMSG
            
        }elseif ($CustomColumn) {
            Send-UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $cleanresponse[-1].message 
            #$cleanresponse[-1].message 

        }     
        
            ForEach ( $ErrorResult in $Script:cleanresponse ) {
  
            $Script:ResponseDetail += New-Object -TypeName PSObject -Property @{ DeviceId      = [String]$DeviceResult.AccountId;
                                                                     DeviceName     = $DeviceResult.Settings.AN  -join '';
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
            Write-output "  DeviceId# $($DeviceResult.AccountId) - Clearing Custom Columns for Resolved Errors or Errors older than $days Days "
        }  
    }
#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $Script:ResponseDetail = @()

    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Send-GetPartnerInfo $Script:cred0

    Send-ClearErrorDevices
        
    Send-GetErrorDevices

    $filterDate = (Get-Date).AddDays(-$days)

    $ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | sort-object partnername,deviceid,groupid | format-table   

    if ($Script:GridView) {  

        $ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | sort-object partnername,deviceid,groupid | Out-GridView -Title " Devices with Errors Since $Filterdate"
    }

    if ($Script:Export) {
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_DeviceErrors_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $script:ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8 -append

    ## Generate XLS from CSV
    
        $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile
      
        Write-output $Script:strLineSeparator

    ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
        
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
        Write-Output "  CSV Path = $Script:csvoutputfile"
        Write-Output "  XLS Path = $Script:xlsoutputfile"
        Write-Output ""
        Start-Sleep -seconds 10
    }

