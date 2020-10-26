<## ----- About: ----
    # Get Device Errors
    # Revision v17 - 2020-09-19
    # Author: Eric Harless, Head Backup Nerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Behavior: ----
    # (Support for Standalone edition of SolarWinds Backup Only)
    #
    # Use with -Clear parameter to Remove Stored API Credentials at start of script
    # Use with -GridView parameter to Display output via Powershell Out-Gridview
    # Use with -CustomColumn parameter to pass last error to Backup.Management Column AA2045
    #
    # Authenticate to Backup.Management  (Supports the Standalone edition of SolarWinds Backup Only)
    # Check for \ Store Sercure Credentials 
    # Authenticate
    # Get Last Session Errors for DeviceId
    # Optionally post Error Message to Backup.Management
    #
    # Partner must Add Custom Column AA2045 to theri backup.Management console to view last error message.
    #    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm 
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/console/custom-columns.htm?Highlight=custom%20column
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [Alias("Clear")] [switch]$ClearCredentials,  ## Remove Stored API Credentials at start of script
        [Parameter(Mandatory=$False)] [switch]$GridView,   ## Display Output via Powershell Out-Gridview
        [Parameter(Mandatory=$False)] [switch]$CustomColumn   ## Update Backup.Manangement Custom Column

    )   

Clear-Host

# ----- Variables and Paths ----
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    Write-output "  Get Device Errors"
    Write-output ""
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n" $Syntax
    $urlJSON = 'https://api.backup.management/jsonapi'


# ----- End Variables and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
        Write-Output $Script:strLineSeparator     
        Write-Output "  Enter Exact, Case Sensitive Partner Name for SolarWinds Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($partnerName.length -eq 0)
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

        if (($Partner.result.result.Level -ne "Root") -and ($Partner.result.result.Level -ne "Sub-root") -and ($Partner.result.result.Level -ne "Distributor")) {
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
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Cookie", "__cfduid=d7cfe7579ba7716fe73703636ea50b1251593338423; visa=$Script:visa")
    
        $Script:response = Invoke-RestMethod -Uri $url2 `
            -Method GET `
            -Headers $headers `
            -ContentType 'application/json' `
            -WebSession $websession `

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
        
        $body = "{
        `n    `"jsonrpc`":`"2.0`",
        `n    `"id`":`"jsonrpc`",
        `n    `"visa`":`"$Visa`",
        `n    `"method`":`"UpdateAccountCustomColumnValues`",
        `n    `"params`":{
        `n      `"accountId`": $DeviceId,
        `n      `"values`": [[$ColumnId,`"$Message`"]]
        `n    	}
        `n    }
        `n"
        
        $updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body
        #$updateCC | ConvertTo-Json
        }
    
    Function Send-GetErrorDevices {

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.OrderBy = "T7 ASC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = 500
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        #Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"
    
        $Script:ErrorDevices = $webrequest | convertfrom-json
        #$Script:visa = $authenticate.visa
    
        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-Output $Script:strLineSeparator
            Write-output " Searching  $PartnerName "
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "PartnerName or UID Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:ErrorDeviceDetail = @()

        Write-Output "  Requesting error details for $($ErrorDevices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $ErrorDevices.result.result ) {
        #write-output "$($Deviceresult.accountId) Count $($Deviceresult.settings.T7)"
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
            Send-UpdateCustomColumn $DeviceResult.AccountId 2045 $RestartedMSG
        }elseif (($CustomColumn) -and ($cleanresponse.message -like "*Operation Aborted*")) {
                $AbortedMSG = $cleanresponse.message | Where-Object {$_ -like "*Operation Aborted*"}
                Send-UpdateCustomColumn $DeviceResult.AccountId 2045 $AbortedMSG
                $abortedmsg
        }elseif ($CustomColumn) {
            Send-UpdateCustomColumn $DeviceResult.AccountId 2045 $cleanresponse[-1].message 
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
        
        #write-output "$($Deviceresult.accountId) Time $($ErrorResult.time)"

        }   
    }

    Function Send-ClearErrorDevices {

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = 100
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        #Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"
    
        $Script:ZeroErrorDevices = $webrequest | convertfrom-json
        #$Script:visa = $authenticate.visa
    }
#region ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $Script:ResponseDetail = @()

    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Send-GetPartnerInfo $Script:cred0

    $filter1 = "(aa2045 AND T7 == 0) OR (aa2045 AND TS < 3.days().ago())"
    
    Send-ClearErrorDevices

    ForEach ( $DeviceResult in $ZeroErrorDevices.result.result ) {
    Send-UpdateCustomColumn $DeviceResult.AccountId 2045 ""
    Write-output "  DeviceId# $($DeviceResult.AccountId) - Error Resolved or older than 3 Days "   
    } 
 
    $filter1 = "aa2045 OR ((T7 >= 1) AND TS > 3.days().ago())"
        
    Send-GetErrorDevices

    $filterDate = (Get-Date).AddDays(-3)

    #$ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | sort-object partnername,deviceid,groupid | format-table
   
    $ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | sort-object partnername,deviceid,groupid | format-table   

    if ($GridView) {  
        #$ResponseDetail | select-object partnername,deviceid,DeviceName,groupid,time,count,message,file | sort-object partnername,deviceid,groupid | Out-GridView -Title " Devices with Errors"

        $ResponseDetail | select-object Partnername,deviceid,DeviceName,groupid,time,count,message,file | where-object {$_.time -ge $filterdate}  | sort-object partnername,deviceid,groupid | Out-GridView -Title " Devices with Errors Since $Filterdate"
    }
