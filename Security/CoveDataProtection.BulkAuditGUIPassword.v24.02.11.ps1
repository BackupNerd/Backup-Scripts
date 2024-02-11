<## ----- About: ----
    # Cove Data Protection | Bulk Audit GUI Password
    # Revision v24.02.11 - 2024-02-11
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com
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
    # For use with the standalone edition of N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Get | Set Secure Credentials & Authenticate
    # Get Last GUI Password change from Device Audit for Server and Workstation devices
    # Post password hash and audit date to a Custom Column
    # Note: M365 domains and documents devices do not support remote commands / GUI passwords
    #
    # Use the -Days ## parameter to select the number of days back to search for active devices
    # Use the -DeviceCount ## (default=2000) parameter to define the maximum number of devices returned
    # Use the -DefaultPartner parameter to specify a default partner name to execute the script against
    # Use the -GridView switch parameter to view data in Powershell Out-Gridview
    # Use the -Export switch parameter to generate CSV/XLS output files 
    # Use the -Launch switch parameter to launch the CSV/XLS file after completion
    # Use the -Delimiter parameter to specify ',' or ';' delimiter for CSV/XLS files
    # Use the -ExportPath (?:\Folder) parameter to specify CSV/XLS file path
    # Use the -CustomColumn switch parameter to update custom column data at https://Backup.Management 
    #   Note: Cove SuperUser credentials are required to update custom columns
    # Use the -ColumnCode parameter to specify the colume to update
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/devices.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/custom-columns.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/remote-commands.htm#SetBackupPassword
    
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [Int]$Days = 30,                              ## Number of days to search for active devices
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 2000,                     ## Specify the maximum number of devices results to return
        [Parameter(Mandatory=$False)] [string]$DefaultPartner = "",                 ## Specify a default partner name (case sensitive)
        [Parameter(Mandatory=$False)] [switch]$GridView = $true,                    ## Display output via Powershell Out-Gridview 
        [Parameter(Mandatory=$False)] [switch]$Export = $true,                      ## Generate CSV/XLS output files
        [Parameter(Mandatory=$False)] [switch]$Launch = $true,                      ## Launch CSV/XLS output file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                     ## Specify ',' or ';' delimiter for CSV/XLS files
        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",        ## Export Path (Script location is the default)
        [Parameter(Mandatory=$False)] [switch]$CustomColumn = $true,                ## Update custom column at https://Backup.Management 
        [Parameter(Mandatory=$False)] [string]$ColumnCode = "AA2048",               ## Column code to update
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                     ## Remove stored API credentials at script run
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    #Requires -Version 5.1
    $ConsoleTitle   = "Cove Data Protection | Bulk Audit GUI Password"              ## Update with full script name
    $ShortTitle     = "AuditGUIPW"                                                  ## Update with short script name
    $host.UI.RawUI.WindowTitle = $ConsoleTitle

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

    Write-Output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "`n  Script Parameter Syntax:`n`n  $Syntax"
       
    $CurrentDate    = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
    $ShortDate      = Get-Date -format "yyyy-MM-dd"
   
    if ($ExportPath) {
        $ExportPath = Join-path -path $ExportPath -childpath "$($ShortTitle)_$($ShortDate)" 
    }else{
        $ExportPath = Join-path -path $dir -childpath "$($ShortTitle)_$($ShortDate)"
    }

    If ($Export -and $ExportPath) {mkdir -force -path $ExportPath | Out-Null}

    Write-Output "  Current Parameters:"
    Write-Output "  -DefaultPartner = $DefaultPartner"
    Write-Output "  -Days           = $Days"
    Write-Output "  -DeviceCount    = $DeviceCount"
    Write-Output "  -ColumnCode     = $ColumnCode"
    Write-Output "  -CustomColumn   = $CustomColumn"
    Write-Output "  -Gridview       = $GridView"
    Write-Output "  -Export         = $Export"
    Write-Output "  -Launch         = $Launch"
    Write-Output "  -Delimiter      = $Delimiter"
    Write-Output "  -ExportPath     = $ExportPath"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'

    $Script:MXB_Path = "C:\ProgramData\MXB\"
    $Script:APIcredfile = join-path -Path $Script:MXB_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-path -path $APIcredfile

    #$filterDate = (Get-Date).AddDays(-$Days)
    #$counter = 1

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
        
        Authenticate  ## Attempt API Authentication

    } ## Set API credentials if not present

    Function Get-APICredentials {

        $Script:MXB_Path = "C:\ProgramData\MXB\"
        $Script:APIcredfile = join-path -Path $Script:MXB_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
        $Script:APIcredpath = Split-path -path $APIcredfile

        if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
            Remove-Item -Path $Script:APIcredfile
            $ClearCredentials = $Null
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential File Cleared"
            Authenticate  ## Retry Authentication
            
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

    Function Authenticate {
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

        #Debug Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

        if ($authenticate.visa) { 
            $Script:visa = $authenticate.visa
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
                Authenticate
            }
        }
    } ## Check Visa Time and Authenticate again if needed

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $epoch = $epoch.AddSeconds($inputUnixTime)
        return $epoch
        }else{ return ""}
    } ## Convert epoch time to date time 

    Function Get-LogTimeStamp {
        #Get-Date -format s
        return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    } ## Output proper timeStamp for log file

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
            Write-Output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
        }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
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

    Function AuditDeviceGUIPassword($DeviceId) {
        $url2 = "https://backup.management/web/accounts/properties/api/audit?accounts.SelectedAccount.Id=$deviceId&accounts.SelectedAccount.StorageNode.Audit.Shift=0&accounts.SelectedAccount.StorageNode.Audit.Count=1&accounts.SelectedAccount.StorageNode.Audit.Filter=execute%20command%3A%20set%20gui%20password"
        $method = 'GET'

        $params = @{
            Uri         = $url2
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $websession
            TimeoutSec  = 180
            ContentType = 'application/json; charset=utf-8'
        }   
    
        $Script:AuditResponse = Invoke-RestMethod @params 
  
        $response = [string]$AuditResponse -replace("[][]","") -replace("[{}]","") -creplace("ESCAPE","") -replace('    ','') -replace("`n`n`n","") -replace(",`n,","") -replace("rows: ","") -split("`n")
        $response = $response -replace (": "," = ") | ConvertFrom-StringData -ErrorAction SilentlyContinue 

        $CMDdate = ($response.Timestamp -replace ('",','') -replace ('"','') -split(','))[0]
        $details = $response.details -replace ('",','') -replace ('"','') -replace (' [(]hash[)]','') -replace ('password','hash') -replace (' restore_only = disallow ',' ') -replace (' restore_only = allow ',' AllowRestore ')            
        if ($details -like "*hash = 5A10*") {$details = "cleared "}

        Write-Output "  DeviceId# $DeviceId GUI Password $details$cmddate"   
        
        $Script:cleanresponse = $details + $cmddate
    }
    
    Function UpdateCustomColumn($DeviceId,$ColumnCode,$Message) {
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
        `n      `"values`": [[$ColumnCode,`"$Message`"]]
        `n    	}
        `n    }
        `n"
        
        $updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body

        $updatecc.error.message
    }

    Function Send-GetDevices {
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = "( AT==1 AND TS > $days.days().ago() ) OR !(AA2048 =='')"
        $data.params.query.Columns = @("AU","TS","AR","AN","MN","AL","CD","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.OrderBy = "AU ASC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable global:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        #Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"
    
        $Script:Devices = $webrequest | convertfrom-json
        $Script:visa = $Devices.visa
    
        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-Output $Script:strLineSeparator
            Write-Output "  Searching  $PartnerName "
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "PartnerName or UID Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:DeviceDetail = @()

        Write-Output "  Requesting audit details for $($Devices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $Devices.result.result ) {

            Get-VisaTime    
            AuditDeviceGUIPassword $DeviceResult.AccountId    

            if ($CustomColumn) {
                Write-Output "  Updating Custom Column $Columncode"
                UpdateCustomColumn $DeviceResult.AccountId ($Columncode.replace("AA","")) $cleanresponse
            }
    
            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{
                PartnerID      = [string]$DeviceResult.PartnerId ;
                PartnerName    = $DeviceResult.Settings.AR -join '' ;
                Reference      = $DeviceResult.Settings.PF -join '' ; 
                AccountID      = [String]$DeviceResult.AccountId;
                DeviceName     = $DeviceResult.Settings.AN -join '' ;
                ComputerName   = $DeviceResult.Settings.MN -join '' ;
                DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                DataSources    = $DeviceResult.Settings.AP -join '' ;
                OS             = $DeviceResult.Settings.OS -join '' ;
                Location       = $DeviceResult.Settings.LN -join '' ;
                Product        = $DeviceResult.Settings.PN -join '' ;
                ProductID      = $DeviceResult.Settings.PD -join '' ;
                Profile        = $DeviceResult.Settings.OP -join '' ;
                ProfileID      = $DeviceResult.Settings.OI -join '' ;
                Notes          = $DeviceResult.Settings.AA843 -join '' ; 
                TempInfo       = $DeviceResult.Settings.AA77 -join '' ;
                GUIPassword    = $cleanresponse 
            }
        }   
    }

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $Script:DeviceDetail = @()

    Authenticate

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    if ($defaultpartner) {
        Send-GetPartnerInfo $DefaultPartner
    }else{
        Send-GetPartnerInfo $Script:cred0
    }
    

    #$filter1 = ""

    Send-GetDevices $partnerId

    if ($Script:GridView) {  
        $Script:DeviceDetail | select-object PartnerID,PartnerName,Reference,AccountID,DeviceName,ComputerName,DeviceAlias,DataSources,OS,Location,Product,ProductID,Profile,ProfileID,Notes,TempInfo,GUIPassword | out-gridview -Title "GUI Password Device Audit for Server & Workstation Devices. ( Active in the last $days days or from a prior value in the device audit)"
    }

    if ($Script:Export) {
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_$($ShortTitle)_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $script:DeviceDetail | select-object PartnerID,PartnerName,Reference,AccountID,DeviceName,ComputerName,DeviceAlias,DataSources,OS,Location,Product,ProductID,Profile,ProfileID,Notes,TempInfo,GUIPassword | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8 -append

        $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile       
        Write-Output $Script:strLineSeparator
        
        if ($Launch) {
            If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
                Start-Process "$xlsoutputfile"
                Write-Output $Script:strLineSeparator
                Write-Output "  Opening XLS file"
                }else{
                Start-Process "$csvoutputfile"
                Write-Output $Script:strLineSeparator
                Write-Output "  Opening CSV file"
                Write-Output $Script:strLineSeparator            
                }
            } ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)

        Write-Output $Script:strLineSeparator
        Write-Output "  CSV Path = $Script:csvoutputfile"
        Write-Output "  XLS Path = $Script:xlsoutputfile"
        Write-Output ""
        Start-Sleep -seconds 10
    } ## Generate CSV / XLS and Optionally Launch

