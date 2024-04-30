<# ----- About: ----
    # Cove Update Device Alias from Custom Column
    # Revision v02 - 2023-09-07
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
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Get list of devices with data in a specific custom column field and write column data to the Alias Field
    #
    # Use the -Customer parameter to specify the case sensitive partner to lookup
    # Use the -DeviceCount ## (default=1000) parameter to define the maximum number of devices returned
    # Use the -ColumnCode parameter to specify which columncode to read data from
    #   Note: Partner must add the custom column specified to their https://backup.management console to view or edit it.
    # Use the -TaskPath parameter to assign a differet path to store credentials and output data
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/console/custom-columns.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [string]$Customer = "Webistix",                                       ## Customer name to process
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 1000,                                             ## Change Maximum Number of devices results to return
        [Parameter(Mandatory=$False)] [string]$ColumnCode = "AA2500",                                       ## Column Code ShortID to Lookup
        [Parameter(Mandatory=$False)] [string]$TaskPath = "$env:programdata\BackupNerdScripts",             ## Path to Store/Invoke Scheduled Backup Nerd Script, Credentials, Task and Logs
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                             ## Remove Stored API Credentials at start of script
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Update Device Alias from Custom Column"            # Set title varible for Console Window
    $Script:strLineSeparator = "  ---------`n"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle                          # Push title variable to Console Windows
    $scriptpath = $MyInvocation.MyCommand.Path                          # when ran, sets variable for the location of the current script
    Write-Output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    Split-Path $ScriptPath | Push-Location                              # split the file and directory of the script and Change directory to the script path.
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-output "  Current Parameters:"
    Write-output "  -Customer      = $Customer"
    Write-output "  -DeviceCount   = $DeviceCount"
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
        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($partnerName.length -eq 0)
        $PartnerName | out-file $APIcredfile

        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able | Cove Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
        
        Start-Sleep -milliseconds 300

        Send-APICredentialsCookie  ## Attempt API Authentication

    }  ## Set API credentials if not present

    Function Get-APICredentials {

        $Script:APIcredfile = join-path -Path $Taskpath -ChildPath "$env:computername API_Credentials.Secure.enc"
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
                    Write-Output    $Script:strLineSeparator 
                    
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
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.partner = $Script:cred0
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        WebSession  = $websession
        ContentType = 'application/json; charset=utf-8'
    }   

    $Script:Session = Invoke-RestMethod @params 


    if ($Session.visa) { 

        $Script:visa = $Session.visa
        }else{
            Write-Output    $Script:strLineSeparator
            $session.error.message 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            $visatime
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

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function CallJSONOLD($url,$object) {

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

        $RestrictedPartnerLevel = @("Root","SubRoot")
            
        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
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
            Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            if ($customer) {
                Send-GetPartnerInfo $customer
            }else{
                $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
                Send-GetPartnerInfo $Script:partnername
            }
        }

        if ($partner.error) {
            write-output "  $($partner.error.message)"
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername
        }

        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-output "  Searching  $PartnerName "
            Write-Output $Script:strLineSeparator
        }else{
           cls
            Write-Output "  PartnerName Not Found"
            Write-Output $Script:strLineSeparator
        }

    } ## get PartnerID and Partner Level

    Function Send-GetCustomColumnDevices {

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $ColumnCode
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","$ColumnCode")
        $data.params.query.OrderBy = "TS DESC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   
   
        $Script:Devices = Invoke-RestMethod @params 
    
        $Script:DeviceDetail = @()
        
        Write-Output "  Requesting Data from Custom Column $ColumnCode for $($Devices.result.result.count) devices.`n"
        Write-Output "  Please be patient, this could take a few moments."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $Devices.result.result ) {
            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ 
                AccountID      = [String]$DeviceResult.AccountId;
                PartnerID      = [string]$DeviceResult.PartnerId;
                DeviceName     = $DeviceResult.Settings.AN -join '' ;
                DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                PartnerName    = $DeviceResult.Settings.AR -join '' ;
                Reference      = $DeviceResult.Settings.PF -join '' ;
                DataSources    = $DeviceResult.Settings.AP -join '' ;                                                                
                Account        = $DeviceResult.Settings.AU -join '' ;
                Location       = $DeviceResult.Settings.LN -join '' ;
                Notes          = $DeviceResult.Settings."$ColumnCode" -join '' ;
                Product        = $DeviceResult.Settings.PN -join '' ;
                ProductID      = $DeviceResult.Settings.PD -join '' ;
                Profile        = $DeviceResult.Settings.OP -join '' ;
                OS             = $DeviceResult.Settings.OS -join '' ;                                                                
                ProfileID      = $DeviceResult.Settings.OI -join '' 
            }
        }     
    }   

    Function Send-UpdateAlias { 
            
        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'ModifyAccount'
        $data.params = @{}
        $data.params.accountInfo = @{}
        $data.params.accountInfo.Id = [int]$entry.AccountID
        $data.params.accountInfo.NameAlias = $entry.Notes

        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:ModifyAccountSession = Invoke-RestMethod @params 
        $Script:ModifyAccountSession.error.message
    }

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    Send-APICredentialsCookie
    Send-GetPartnerInfo $Script:cred0
    Send-GetCustomColumnDevices

    foreach ($entry in $Script:DeviceDetail) {
        Write-output "  Processing $($entry.accountid) | $($entry.Notes) | $($entry.DeviceName)`n"
        Send-UpdateAlias
    }

