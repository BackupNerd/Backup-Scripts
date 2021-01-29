<# ----- About: ----
    # Get SW Backup Session File Details
    # Revision v19 - 2021-01-28
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

<# ----- Compatibility: ----
    # For use with the Standalone edition of SolarWinds Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Enumerate Devices
    # Enumerate Sessions
    # Get File Detail from Most Recent Session 
    #
    # Use the -Action "FSSentSize" (default) parameter to get a list of the largest transfered files from the most recent FileSystem backup session
    # Use the -Action "ESXCount" parameter to get the VMs selected/ protected during the most recent VMware ESX Host-level backup session
    # Use the -Action "HVCount" parameter to get the VMs selected/ protected during the most recent Hyper-V Host-level backup session
    # Use the -Action "VHDinFS" parameter to get a list of VHD files that were selected/ protetected during the most recent FileSystem backup session
    # Use the -DeviceCount ## (default=20) parameter to define the maximum number of devices to process
    # Use the -FileCount ## (default=50) parameter to define how many files to return per device
    # Use the -GridView switch parameter to display output via Powershell Out-Gridview
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] $Action = "FSSent",                           ## Change default Action to (FSSent,FSSize,ESXCount,HVCount,VHDinFS)
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 20,                       ## Change default Number of devices to lookup
        [Parameter(Mandatory=$False)] [int]$FileCount = 50,                         ## Change default Number of entries to lookup per device
        [Parameter(Mandatory=$False)] [switch]$GridView,                            ## Display output via Powershell Out-Gridview
        [Parameter(Mandatory=$False)] [switch]$Launch,                              ## Launch XLS or CSV file 
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',                     ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",                ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                     ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Get Session File Details`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    Write-output "  Current Parameters:"
    Write-Output "  -Action      = $action"
    Write-Output "  -DeviceCount = $DeviceCount"
    Write-Output "  -FileCount   = $FileCount"
    Write-Output "  -Delimiter   = $Delimiter" 
    Write-Output "  -Grid-View   = $GridView"
    Write-Output "  -Launch      = $Launch"
    Write-output "  -ExportPath  = $ExportPath"
    Write-output "  -Delimiter   = $Delimiter"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $Script:True_path = "C:\ProgramData\MXB\"
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
    $enc = [System.Text.Encoding]::UTF8
    $Counter=0

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
    Write-Output "  Enter Exact, Case Sensitive Partner Name for SolarWinds Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($partnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -UserName "" -Message 'Enter Login UserName or Email and Password for SolarWinds Backup.Management API'
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
    -ContentType "application/json; charset=utf-8" `
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

Function Get-LogTimeStamp {
    #Get-Date -format s
    return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
}  ## Output proper timeStamp for log file 

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


#region ----- Backup.Management JSON Calls ----

Function CallJSON($url,$object) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($object)
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
    }  ## Call JSON function

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
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json

    $RestrictedPartnerLevel = @("Root","Sub-root","Distributor")

    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$script:Level = $Partner.result.result.Level
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-output "  $String:PartnerName - $partnerId - $Uid"
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

    }

} ## get PartnerID and Partner Level    


    Function Get-Devices {

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $DeviceFilter
        $data.params.query.Columns = @("AU","AR","AN","AL","LN","OP","OI","OS","PD","AP","PF","PN","AA843","AA77","T7")
        $data.params.query.OrderBy = "F5 DESC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -TimeoutSec 600 `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        #Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"
    
        $Script:Devices = $webrequest | convertfrom-json
        #$Script:visa = $authenticate.visa
    
        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-Output $Script:strLineSeparator
            Write-output "  Searching  $string:PartnerName "
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "PartnerName or UID Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:DeviceDetail = @()

        Write-Output "  Requesting Session $plugin details for $($Devices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $Devices.result.result ) {

            $Counter = ($counter + 1)
            Write-output "  # $counter"
            $Script:ProcessedFilesResponse = $null
            $Script:Cleanresponse = $null
            $Script:response2 = $null
            $Script:response = $null

            $Script:DeviceName = $DeviceResult.Settings.AN -join ''
            $Script:EndCustomername = $DeviceResult.Settings.AR -join '' ;


            Write-Output "  $(Get-LogTimeStamp) EndCustomer | $EndCustomerName | Device | $deviceName | $($DeviceResult.AccountId) | Getting Session IDs"
            Get-SessionIds $DeviceResult.AccountId


            if (($Script:Session.count -ge 1) -and($sessionId)) {

                Write-Output "  $(Get-LogTimeStamp) EndCustomer | $EndCustomerName | Device | $deviceName | $($DeviceResult.AccountId) | Requesting $filecount Files from Session ID $SessionId" 
           
                Get-ProcessedFiles $DeviceResult.AccountId
        
            }

        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [String]$DeviceResult.AccountId;
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
                                                                    ProfileID      = $DeviceResult.Settings.OI -join '' ;
                                                                    SessionId      = $sessionID }                                                                                                     
  

        }   




    }

    Function Get-SessionIds($DeviceId) {
    
        $url2 = "https://backup.management/web/accounts/properties/api/history-resent?accounts.SelectedAccount.Id=$DeviceId"
        
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Cookie", "__cfduid=d7cfe7579ba7716fe73703636ea50b1251593338423; visa=$Script:visa")
    
        $Script:SessionIdResponse = Invoke-RestMethod -Uri $url2 `
        -Method GET `
        -Headers $headers `
        -ContentType "application/json; charset=utf-8" `
        -TimeoutSec 600 `
        -WebSession $websession `
  
        $response = [string]$SessionIdResponse -replace("[][]","") -replace("[{}}]","") -replace('    ','') -replace("`n`n`n","") -replace("\+","") -replace('"SessionId" : ','') -split("`n,")
        $response = $response | convertfrom-string -delimiter ",`n"


        $cleanresponse = $response | Select-object @{N="SessionId";E={$_.P1}},@{N="Plugin";E={$_.P4}},@{N="SelectedSize";E={$_.P11}},@{N="TransferredSize";E={$_.P15}},@{N="TrsStatus";E={$_.P9}}
        
        $Script:Session = $cleanresponse | Where-Object {(($_.plugin -like "*$plugin*") -and ($_.TransferredSize -notlike "*0 B*"))}
        
        If ($Session.count) {
            Write-Output "  $(Get-LogTimeStamp) EndCustomer | $EndCustomerName | Device | $deviceName | $($DeviceResult.AccountId) | Found $($Session.count) $plugin Sessions"
            [int]$Script:SessionId = [int]$session.sessionid[0]
    
        }Else{
            Write-Output "  $(Get-LogTimeStamp) EndCustomer | $EndCustomerName | Device | $deviceName | $($DeviceResult.AccountId) | Found 0 Usable $plugin Sessions"
        }
        #$session[0]
        
    }

    Function Get-ProcessedFiles($DeviceId) {

        $Script:ProcessedFilesResponse = $null
        $Script:Cleanresponse = $null
        $Script:response2 = $null
        $Script:response = $null
    
        $url3 = "https://backup.management/web/accounts/properties/api/processed-files?accounts.SelectedAccount.Id=$DeviceId&accounts.SelectedAccount.StorageNode.BackupFiles.SessionId=$SessionId&accounts.SelectedAccount.StorageNode.BackupFiles.Shift=$fileshift&accounts.SelectedAccount.StorageNode.BackupFiles.Count=$FileCount&accounts.SelectedAccount.StorageNode.BackupFiles.Filter=$FileFilter&accounts.SelectedAccount.StorageNode.BackupFiles.Column=$filecolumn&accounts.SelectedAccount.StorageNode.BackupFiles.Order=$fileorder"
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Cookie", "__cfduid=d7cfe7579ba7716fe73703636ea50b1251593338423; visa=$Script:visa")
    
        $Script:ProcessedFilesResponse = Invoke-RestMethod -Uri $url3 `
        -Method GET `
        -Headers $headers `
        -ContentType "application/json; charset=utf-8" `
        -TimeoutSec 600 `
        -WebSession $websession `
  
        if ($ProcessedFilesResponse.error) {$Script:ProcessedFilesResponse = $null} else {

        $response = [string]$ProcessedFilesResponse -replace("[][]","") -replace("[}]","")  -creplace("ESCAPE","") -replace("&lt; ","")  -replace('    ','') -replace("`n"," ") -replace (", ,","") -split(" { ")
        $response = $response | convertfrom-string -delimiter '"'

        If ($response) {$response[0] = $null}
    
        if ($Action -eq "HVCount") {
            $Script:response2 = $response | Select-object @{N="EndCustomerName";E={$EndCustomerName}},@{N="DeviceName";E={$DeviceName}},@{N="DeviceID";E={$DeviceId}},@{N="SessionID";E={$SessionId}},@{N="Start";E={$_.P2}},@{N="Duration";E={$_.P4}},@{N="Size";E={[long]($_.P6.replace(" ","")/1)}},@{N="Sent";E={[long]($_.P8.replace(" ","")/1)}},@{N="Path";E={$_.P10.split('\')[3]}} | Where-object {$_.Start -ne $null} 
        }Else{
            $Script:response2 = $response | Select-object @{N="EndCustomerName";E={$EndCustomerName}},@{N="DeviceName";E={$DeviceName}},@{N="DeviceID";E={$DeviceId}},@{N="SessionID";E={$SessionId}},@{N="Start";E={$_.P2}},@{N="Duration";E={$_.P4}},@{N="Size";E={[long]($_.P6.replace(" ","")/1)}},@{N="Sent";E={[long]($_.P8.replace(" ","")/1)}},@{N="Path";E={$_.P10}} | Where-object {$_.Start -ne $null} 
        }


        $Script:Cleanresponse = $response2 | Select-object EndCustomerName,DeviceName,DeviceID,SessionID,Start,Duration,@{N="SizeMB";E={[math]::Round([Decimal](($_.Size) /1MB),0)}},@{N="SentMB";E={[math]::Round([Decimal](($_.Sent) /1MB),0)}},Path | Where-object {$_.Start -ne $null} 

        if ($GridView) { $Script:cleanresponse | Out-GridView -Title " $partnername | $DeviceName | $deviceid | Session $Sessionid | TOP $FileCount $FileColumn $FileOrder " }
        
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_$($action)_Files_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"

        If ($Script:Cleanresponse) {$Script:cleanresponse | Select-object * | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8 -append }
    }
    }
#endregion ----- Backup.Management JSON Calls ----    
#endregion ----- Functions ----

    $Script:DeviceDetail = @()

    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
  
    Switch ($Action) {
        'FSSent' {
            $plugin = "FileSystem"
            [int]$FileShift = 0
            $FileFilter = ""
            $FileColumn = "SentSize"     ## Path,Size,SentSize,StartTime,Duration
            $FileOrder = "Descending"     ## Ascending,Descending
            $DeviceFilter = "TS > 7.days().ago() AND FL > 7.days().ago() AND  I78 =~ '*D01*'"
        }
        'FSSize' {
            $plugin = "FileSystem"
            [int]$FileShift = 0
            $FileFilter = ""
            $FileColumn = "Size"     ## Path,Size,SentSize,StartTime,Duration
            $FileOrder = "Descending"     ## Ascending,Descending
            $DeviceFilter = "TS > 7.days().ago() AND FL > 7.days().ago() AND  I78 =~ '*D01*'"
        }
        'ESXCount' {
            $plugin = "VMWare"
            [int]$FileShift = 0
            $FileFilter = ""
            $FileColumn = "StartTime"     ## Path,Size,SentSize,StartTime,Duration
            $FileOrder = "Ascending"     ## Ascending,Descending
            $DeviceFilter = "I78 =~ '*D08*'"
        }
        'HVCount' {
            $plugin = "VssHyperV"
            [int]$FileShift = 0
            $FileFilter = ".vhd"
            $FileColumn = "StartTime"     ## Path,Size,SentSize,StartTime,Duration
            $FileOrder = "Ascending"     ## Ascending,Descending
            $DeviceFilter = "I78 =~ '*D14*'"
        }
        'VHDinFS' {
            $plugin = "FileSystem"
            [int]$FileShift = 0
            $FileFilter = ".vhd"
            $FileColumn = "Path"     ## Path,Size,SentSize,StartTime,Duration
            $FileOrder = "Descending"     ## Ascending,Descending
            $DeviceFilter = "TS > 7.days().ago() AND FL > 7.days().ago() AND  I78 =~ '*D01*' AND I78 =~ '*D14*'"
        }
    }

    Send-GetPartnerInfo $cred0

    Get-Devices $partnerId

    ## Generate XLS from CSV
    
    if ($csvoutputfile) {
        $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile
    }
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