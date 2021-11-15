<# ----- About: ----
    # N-able Backup Get M365 Device Stats
    # Revision v26 - 2021-11-07
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
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select M365 devices
    # Optionally export to XLS/CSV
    # Optionally import CSV to adjust Mail and OneDrive selections
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -AllDevices switch parameter to skip GUI device selection
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -ExportIndividual switch parameter to export M365 user statistics for individual end customers to XLS/CSV files
    # Use the -ExportCombined switch parameter to export combined M365 device and user statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # Use the -Import parameter to load mail and onedrive selections from a modified CSV export file
    # Use the -ExchangeAutoInclude parameter (On,Off) to set AutoInclude Backup for new users (Mandatory)
    # Use the -OneDriveAutoInclude parameter (On,Off) to set AutoInclude Backup for new users (Mandatory)
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm

# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Export")]
    Param (
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$AllPartners = $false,              ## $AllPartners = $true, to Skip GUI partner selection
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$AllDevices = $false,               ## $AllDevices = $true, to Skip GUI device selection
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [int]$DeviceCount = 5000,                   ## Set maximum number of device results to return
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$GridView,                          ## Display output via Powershell Out-Gridview
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$ExportIndividual,                  ## Generate individual End Customer XLS/CSV output files for M365 users
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$ExportCombined = $true,            ## Generate combined XLS/CSV output files for M365 devices and users
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$Launch = $true,                    ## Launch combined XLS/CSV outputfile if generated
        [Parameter(ParameterSetName="Import",Mandatory=$False)] [switch]$Import,                            ## Prompts for path to modified Export file (used to set/ adjsut user selection)
        [Parameter(ParameterSetName="Import",Mandatory=$False)] [String]$ImportPath,                        ## Full path and filename to selection modified .CSV file
        [Parameter(ParameterSetName="Import",Mandatory=$True)] 
            [ValidateSet('On','Off')] [String]$ExchangeAutoInclude,                                         ## (On,Off) to set AutoInclude Backup for new users (Mandatory)
        [Parameter(ParameterSetName="Import",Mandatory=$True)]
            [ValidateSet('On','Off')]  [String]$OneDriveAutoInclude,                                        ## (On,Off) to set AutoInclude Backup for new users (Mandatory)
        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",                                ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                             ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    #Requires -Version 5.1
    $ConsoleTitle = "Get M365 Device Statistics"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"

    $dir = Split-Path $scriptpath
    Push-Location $dir
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    if ($ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "M365_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "M365_$shortdate"}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    If ($exportindividual -or $exportcombined) {mkdir -force -path $ExportPath | Out-Null}
    $urlJSON = 'https://api.backup.management/jsonapi'

    Write-output "  Current Parameters:"
    Write-output "  -AllPartners     = $AllPartners"
    Write-output "  -AllDevices      = $AllDevices"
    Write-output "  -DeviceCount     = $DeviceCount"
    Write-output "  -GridView        = $GridView"
    Write-output "  -ExportCombined  = $ExportCombined"
    Write-output "  -ExportIndvidual = $ExportIndividual"
    Write-output "  -Launch          = $Launch"
    Write-output "  -ExportPath      = $ExportPath"
    Write-output "  -Delimiter       = Cultural Settings"
    
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
    $script:UserId = $authenticate.result.result.id
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

        <#
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }
        #>  ## Increment file name

        $xl.DisplayAlerts = $False  ## Block overwrite file warning
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

        $RestrictedPartnerLevel = @("Root","Sub-root","Distributor")

        if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
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

        }else{$script:visa = $Partner.visa}

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
                }else{
        # (No error)
        
                $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
                
                $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
                $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
                $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
            
                $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                        
                $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
                                
                if ($AllPartners) {
                    $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  All Partners Selected"
                }else{
                    $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single
            
                    if($null -eq $Selection) {
                        # Cancel was pressed
                        # Run cancel script
                        Write-Output    $Script:strLineSeparator
                        Write-Output    "  No Partners Selected"
                        Break
                    }else{
                        # OK was pressed, $Selection contains what was chosen
                        # Run OK script
                        [int]$script:PartnerId = $script:Selection.Id
                        [String]$script:PartnerName = $script:Selection.Name
                    }
                }

        }
        
    }  ## EnumeratePartners API Call

    Function Send-GetDevices {

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script.visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = "AT==2"
        $data.params.query.Columns = @("AR","PF","AN","MN","AL","AU","CD","TS","TL","T3","US","TB","T7","TM","D19F21","GM","JM","D5F20","D5F22","LN","AA843")
        $data.params.query.OrderBy = "CD DESC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $devicecount
        $data.params.query.Totals = @("COUNT(AT==2)","SUM(T3)","SUM(US)")

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
            
            $AccountID = [Int]$DeviceResult.AccountId
            Get-AccountInfoById $AccountID


        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID = [Int]$DeviceResult.AccountId;
                                                                    PartnerID        = [string]$DeviceResult.PartnerId;
                                                                    PartnerName      = $DeviceResult.Settings.AR -join '' ;
                                                                    Reference        = $DeviceResult.Settings.PF -join '' ;
                                                                    Account          = $DeviceResult.Settings.AU -join '' ;  
                                                                    DeviceName       = $DeviceResult.Settings.AN -join '' ;
                                                                    ComputerName     = $DeviceResult.Settings.MN -join '' ;
                                                                    DeviceAlias      = $DeviceResult.Settings.AL -join '' ;
                                                                    Creation         = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                                                                    TimeStamp        = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;  
                                                                    LastSuccess      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '') ;                                                                                                                                                                                                               
                                                                    SelectedGB       = [math]::Round([Decimal](($DeviceResult.Settings.T3 -join '') /1GB),2) ;  
                                                                    UsedGB           = [math]::Round([Decimal](($DeviceResult.Settings.US -join '') /1GB),2) ;
                                                                    Last28Days       = (($DeviceResult.Settings.TB -join '')[-1..-28] -join '') -replace("8",[char]0x26a0) -replace("7",[char]0x23f9) -replace("6",[char]0x23f9) -replace("5",[char]0x2611) -replace("2",[char]0x274e) -replace("1",[char]0x2BC8) -replace("0",[char]0x274c) ;
                                                                    Last28           = (($DeviceResult.Settings.TB -join '')[-1..-28] -join '') -replace("8","!") -replace("7","!") -replace("6","?") -replace("5","+") -replace("2","-") -replace("1",">") -replace("0","X") ;
                                                                    Errors           = $DeviceResult.Settings.T7 -join '' ;
                                                                    Billable         = $DeviceResult.Settings.TM -join '' ;
                                                                    Shared           = $DeviceResult.Settings.D19F21 -join '' ;
                                                                    MailBoxes        = $DeviceResult.Settings.GM -join '' ;

                                                                    OneDrive         = $DeviceResult.Settings.JM -join '' ;
                                                                    SPusers          = $DeviceResult.Settings.D5F20 -join '' ;
                                                                    SPsites          = $DeviceResult.Settings.D5F22 -join '' ;
                                                                    StorageLocation  = $DeviceResult.Settings.LN -join '' ;
                                                                    AccountToken     = $Script:AccountInfoResponse.result.result.Token;
                                                                    Notes            = $DeviceResult.Settings.AA843 -join '' }



        }

    } ## EnumerateAccountStatistics API Call

    Function GetM365Stats {
        Param([Parameter(Mandatory=$true)][Int]$DeviceId) #end param
    
        $url2 = "https://api.backup.management/c2c/statistics/devices/id/$deviceid"
        $method = 'GET'

        $params = @{
            Uri         = $url2
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:M365response = Invoke-RestMethod @params 

        Write-output  "$url2"

        Get-AccountInfoById $deviceID
 
        $script:devicestatistics = $M365response.deviceStatistics | Select-object @{N="Partner";E={$device.partnername}},
                                                                                    @{N="Account";E={$device.DeviceName}},
                                                                                    DisplayName,
                                                                                    EmailAddress,
                                                                                    @{N="Shared";E={$_.shared[0] -replace("TRUE","Shared") -replace("FALSE","") }},
                                                                                    @{N="MailBox";E={$_.datasources.status[0] -replace("unprotected","") }},
                                                                                    @{N="OneDrive";E={$_.datasources.status[1]  -replace("unprotected","")  }},
                                                                                    @{N="SharePoint";E={$_.datasources.status[2]  -replace("unprotected","")  }},
                                                                                    @{N="UserGuid";E={$_.UserId}},
                                                                                    @{N="AccountToken";E={$device.accountToken}} 
                                       

        $devicestatistics | Select-object * | format-table
        if ($gridview) {$devicestatistics | out-gridview -title "$($device.partnername) | $($device.DeviceName)" }
    } 

    Function Get-AccountInfoById ([int]$AccountId) {

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:Visa
        $data.method = 'GetAccountInfoById'
        $data.params = @{}
        $data.params.accountId = [int]$AccountId

        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:AccountInfoResponse = Invoke-RestMethod @params 
    }


    Function EnumerateM365Users ($AccountToken) {

        $url = "https://api.backup.management/management_api"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.method = 'EnumerateUsers'
        $data.params = @{}
        $data.params.accountToken = $accountToken

        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:EnumerateM365UsersResponse = Invoke-RestMethod @params 

        $script:M365UserStatistics = $EnumerateM365UsersResponse.result.result.Users | Select-object @{N="Partner";E={$device.partnername}},
                                                                                                    @{N="Account";E={$device.DeviceName}},
                                                                                                    DisplayName,
                                                                                                    EmailAddress,
                                                                                                    @{N="MailBoxSelection";E={$_.ExchangeInfo.Selection}},
                                                                                                    @{N="MailBoxType";E={$_.ExchangeInfo.MailboxType}},
                                                                                                    @{N="New";E={$_.IsNew[0] -replace("True","New") -replace("False","") }},
                                                                                                    @{N="Deleted";E={$_.IsDeleted[0] -replace("True","Deleted") -replace("False","") }},
                                                                                                    @{N="Shared";E={$_.IsShared[0] -replace("True","Shared") -replace("False","") }},
                                                                                                    @{N="MailBoxLastBackupStatus";E={$_.ExchangeInfo.LastBackupStatus}},
                                                                                                    @{N="MailBoxLastBackupTimestamp";E={Convert-UnixTimeToDateTime ($_.ExchangeInfo.LastBackupTimestamp)}},
                                                                                                    @{N="ExchangeAutoInclude";E={$Script:EnumerateM365UsersResponse.result.result.ExchangeAutoInclusionType}},
                                                                                                    @{N="OneDriveSelection";E={$_.OneDriveInfo.Selection}},
                                                                                                    @{N="OneDriveStatus";E={$_.OneDriveInfo.LicenseStatus}},
                                                                                                    @{N="OneDriveSelectedGib";E={[math]::Round([Decimal](($_.OneDriveInfo.SelectedSize) /1GB),3) }},
                                                                                                    @{N="OneDriveLastBackupStatus";E={$_.OneDriveInfo.LastBackupStatus}},
                                                                                                    @{N="OneDriveLastBackupTimestamp";E={Convert-UnixTimeToDateTime ($_.OneDriveInfo.LastBackupTimestamp)}},
                                                                                                    @{N="OneDriveAutoInclude";E={$Script:EnumerateM365UsersResponse.result.result.OneDriveAutoInclusionType}},
                                                                                                    UserGuid,
                                                                                                    @{N="AccountToken";E={$accounttoken}} 

        $Script:M365UserStatistics | Select-object * | format-table
        if ($gridview) {$Script:M365UserStatistics | out-gridview -title "$($device.partnername) | $($device.DeviceName)" }


    }

    Function Open-FileName($initialDirectory) {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Comma Seperated Value (*.csv)|*.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    } ## GUI Prompt for Filename to open

    Function Format-Hashtable {
        param(
          [Parameter(Mandatory,ValueFromPipeline)]
          [hashtable]$Hashtable,
    
          [ValidateNotNullOrEmpty()]
          [string]$KeyHeader = 'Name',
    
          [ValidateNotNullOrEmpty()]
          [string]$ValueHeader = 'Value'
        )
    
        $Hashtable.GetEnumerator() |Select-Object @{Label=$KeyHeader;Expression={$_.Key}},@{Label=$ValueHeader;Expression={$_.Value}}
    
    }

    Function UpdateDataSource ($AccountToken,$EntityId,$Exchange,$OneDrive) {
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Authorization","Bearer $Script:visa")
        $headers.Add("Content-Type","application/json")

        $body = "{
            `n  `"method`": `"UpdateDataSources`",
            `n  `"params`": {
            `n      `"accountToken`": `"$accountToken`",
            `n      `"dataSources`": {
            `n          `"DataSourceEntityAutoInclusions`": [
            `n              {
            `n                  `"DataSourceType`": `"Exchange`",
            `n                  `"EntityAutoInclusionType`": `"$ExchangeAutoInclude`"
            `n              },
            `n              {
            `n                  `"DataSourceType`": `"OneDrive`",
            `n                  `"EntityAutoInclusionType`": `"$OneDriveAutoInclude`"
            `n              }
            `n          ],
            `n          `"Selections`": [
            `n              {
            `n                  `"EntityId`": `"$entityId`",
            `n                  `"DataSourceSelections`": [
            `n                      {
            `n                          `"DataSourceType`": `"Exchange`",
            `n                          `"SelectionType`": `"$Exchange`"
            `n                      },
            `n                      {
            `n                         `"DataSourceType`": `"OneDrive`",
            `n                         `"SelectionType`": `"$OneDrive`"
            `n                      }
            `n                  ]
            `n              }
            `n          ]
            `n      }
            `n  },
            `n  `"jsonrpc`": `"2.0`",
            `n  `"id`": `"jsonrpc`"
            `n  }
            `n"

            $Script:UpdateDataSource = Invoke-RestMethod 'https://api.backup.management/management_api' -Method 'POST' -Headers $headers -Body $body
            if ($Script:UpdateDataSource.error) {Write-output $Script:UpdateDataSource.error.message}else{Write-output "Updating UserGuid $entityId"} 
        }

    Function UpdateDataSource2 ($AccountToken,$EntityId,$Exchange,$OneDrive) {

        $url = "https://api.backup.management/management_api"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = 'jsonrpc'
        $data.method = 'UpdateDataSources'
        $data.params = @{}
        $data.params.accountToken = $accountToken
        $data.params.dataSources = @{}
        #$data.params.dataSources.DataSourceEntityAutoInclusions = @{}
        #$data.params.dataSources.DataSourceEntityAutoInclusions 
        
        $selected = @{}
        $selected | Format-Hashtable -KeyHeader DataSourceType -ValueHeader SelectionType
        $selected += @{"Exchange"="$Exchange"}
        $selected += @{"OneDrive"="$OneDrive"}
        #[hashtable]$selected1 = $selected | Format-Hashtable -KeyHeader DataSourceType -ValueHeader SelectionType

        $data.params.dataSources.Selections = @()
        $data.params.dataSources.Selections += @{"$EntityId"=$selected}
        $data.params.dataSources.Selections = $data.params.dataSources.Selections | Format-Hashtable -KeyHeader EntityId -ValueHeader DataSourceSelections

        $script:jsondata = (ConvertTo-Json $data -depth 8)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:UpdateDataSourceResponse = Invoke-RestMethod @params 
    }
      
    Function Join-Reports {

        if (Get-Module -ListAvailable -Name Join-Object ) {
            Write-Host "  Module Join-Object Already Installed"
        } 
        else {
            try {
                Install-Module -Name Join-Object   -Confirm:$True -Force
            }
            catch [Exception] {
                $_.message
                Write-Warning "  PS Module 'Join-Object' not found, run with Administrator rights to install" 
                exit
            }
        }

        $Script:M365Users = Join-Object -left $script:M365UserStatistics -LeftJoinProperty UserGuid -right $script:devicestatistics -RightJoinProperty UserGuid -Prefix 'Prot_' -Type AllInLeft -RightProperties Mailbox,OneDrive,SharePoint


        $Script:M365UserStatistics = $Script:M365Users | Select-Object  Partner,
                                                                        Account,
                                                                        DisplayName,
                                                                        EmailAddress,
                                                                        Prot_Mailbox,
                                                                        MailBoxSelection,
                                                                        MailBoxType,
                                                                        New,
                                                                        Deleted,
                                                                        Shared,
                                                                        MailBoxLastBackupStatus,
                                                                        MailBoxLastBackupTimestamp,
                                                                        ExchangeAutoInclude,
                                                                        Prot_OneDrive,
                                                                        OneDriveSelection,
                                                                        OneDriveStatus,
                                                                        OneDriveSelectedGib,
                                                                        OneDriveLastBackupStatus,
                                                                        OneDriveLastBackupTimestamp,
                                                                        OneDriveAutoInclude,
                                                                        Prot_SharePoint,
                                                                        UserGuid,
                                                                        AccountToken	


    }

    Function ExitRoutine {
        Write-Output $Script:strLineSeparator
        Write-Output "  Secure credential file found here:"
        Write-Output $Script:strLineSeparator
        Write-Output "  & $APIcredfile"
        Write-Output ""
        Write-Output $Script:strLineSeparator
        #Start-Sleep -seconds 5
        }

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----


    Switch ($PSCmdlet.ParameterSetName) { 
        'Import' {
            Send-APICredentialsCookie
            Write-Output $Script:strLineSeparator
            Write-Output "" 
        
            if ($importpath) {

                $M365UserSelectionCSVFile = import-Csv -Path $ImportPath
                        
            }else{
                $OpenFileName = Open-FileName "$ExportPath" ; if ($null -eq $OpenFileName) {Break} 
                $M365UserSelectionCSVFile = import-Csv -delimiter $delimiter -Path $OpenFileName
            }
                $M365UserSelectionCSVFile | Out-GridView -Title "Summary of M365 user selections being made. Refresh your Backup.Management console to confirm."

                foreach ($M365user in $M365UserSelectionCSVFile) {
                    #Start-Sleep -Milliseconds 300
                    Visa-Check
                    if ((($M365User.MailBoxSelection -Ceq "Selected") -or ($M365User.MailBoxSelection -Ceq "Excluded")) -and (($M365User.OneDriveSelection -Ceq "Selected") -or ($M365User.OneDriveSelection -Ceq "Excluded"))) {
                        updatedatasource $M365User.AccountToken $M365User.UserGuid $M365User.MailBoxSelection $M365User.OneDriveSelection
                    }else{
                        Write-Output "CSV Row is Incorrectly Formatted, [MailBoxSelection] & [OneDriveSelection] cell values must be 'Selected' or 'Excluded'"
                    }
            
                }
            }
     
        'Export' { 
            Send-APICredentialsCookie

            Write-Output $Script:strLineSeparator
            Write-Output "" 

            Send-GetPartnerInfo $Script:cred0

            if ($AllPartners) {}else{Send-EnumeratePartners}

            Send-GetDevices $partnerId

            if ($AllDevices) {
                $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,AccountID,DeviceName,SelectedGB,UsedGB,Last28Days,Errors,Billable,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,AccountToken,Notes
            }else{
                $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,AccountID,DeviceName,SelectedGB,UsedGB,Last28Days,Errors,Billable,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,AccountToken,Notes  | Out-GridView -title "Current Partner | $partnername" -OutputMode Multiple}
                Visa-Check
            if ($null -eq $SelectedDevices) {
                # Cancel was pressed
                # Run cancel script
                Write-Output    $Script:strLineSeparator
                Write-Output    "  No Devices Selected"
                Break
            }else{
                # OK was pressed, $Selection contains what was chosen
                # Run OK script
                $DeviceDetail | Select-Object PartnerId,PartnerName,AccountID,DeviceName,SelectedGB,UsedGB,Last28,Errors,Billable,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,Notes  | Sort-object PartnerName,AccountId | format-table

                If ($Script:Exportcombined) {
                    $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_M365_Devices_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
                    $SelectedDevices | Select-object * | Export-CSV -useCulture -path "$csvoutputfile" -NoTypeInformation -Encoding UTF8}
                    
            }

            foreach ($device in $SelectedDevices) {
                GetM365Stats $Device.AccountID
                EnumerateM365Users $device.AccountToken
                join-reports
                
                If ($Script:Exportcombined) {
                    #$Script:csvoutputfile2a = "$ExportPath\$($CurrentDate)_M365_Protected_Users_Summary_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
                    #$Script:DeviceStatistics | Select-object * | Export-CSV -useCulture -path "$csvoutputfile2a" -NoTypeInformation -Encoding UTF8 -append

                    $Script:csvoutputfile2b = "$ExportPath\$($CurrentDate)_M365_User_Export_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
                    $Script:M365UserStatistics | Select-object * | Export-CSV -useCulture -path "$csvoutputfile2b" -NoTypeInformation -Encoding UTF8 -append
                }

                If ($Script:ExportIndividual) {
                    #$Script:csvoutputfile3a = "$ExportPath\$($CurrentDate)_M365_Protected_Users_Summary_$($device.Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($device.PartnerId).csv"
                    #$Script:DeviceStatistics | Select-object * | Export-CSV -useCulture -path "$csvoutputfile3a" -NoTypeInformation -Encoding UTF8 -append

                    $Script:csvoutputfile3b = "$ExportPath\$($CurrentDate)_M365_User_Export_$($device.Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($device.PartnerId).csv"
                    $Script:M365UserStatistics | Select-object * | Export-CSV -useCulture -path "$csvoutputfile3b" -NoTypeInformation -Encoding UTF8 -append
                }     

                #If ($csvoutputfile3a) {
                #    $xlsoutputfile3a = $csvoutputfile3a.Replace("csv","xlsx")
                #    Save-CSVasExcel $csvoutputfile3a
                #}

                If ($csvoutputfile3b) {
                    $xlsoutputfile3b = $csvoutputfile3b.Replace("csv","xlsx")
                    Save-CSVasExcel $csvoutputfile3b
                }

            }

            ## Generate XLS from CSV
            
            if ($csvoutputfile) {
                $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
                Save-CSVasExcel $csvoutputfile
            }

            #If ($csvoutputfile2a) {
            #    $xlsoutputfile2a = $csvoutputfile2a.Replace("csv","xlsx")
            #    Save-CSVasExcel $csvoutputfile2a
            #}

            If ($csvoutputfile2b) {
                $xlsoutputfile2b = $csvoutputfile2b.Replace("csv","xlsx")
                Save-CSVasExcel $csvoutputfile2b
            }

            Write-output $Script:strLineSeparator

            ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
                
            if ($Launch -and $ExportCombined) {
                If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
                    Start-Process "$xlsoutputfile"
                    #Start-Process "$xlsoutputfile2a"
                    Start-Process "$xlsoutputfile2b"
                    Write-output $Script:strLineSeparator
                    Write-Output "  Opening XLS file"
                    }else{
                    Start-Process "$csvoutputfile"
                    #Start-Process "$csvoutputfile2a"
                    Start-Process "$csvoutputfile2b"
                    Write-output $Script:strLineSeparator
                    Write-Output "  Opening CSV file"
                    Write-output $Script:strLineSeparator            
                    }
            }

            If ($Exportcombined) {
                Write-output $Script:strLineSeparator
                Write-Output "  CSV Path = $Script:csvoutputfile"
                #Write-Output "  CSV Path = $Script:csvoutputfile2a"
                Write-Output "  CSV Path = $Script:csvoutputfile2b"
                Write-Output "  XLS Path = $Script:xlsoutputfile"
                #Write-Output "  XLS Path = $Script:xlsoutputfile2a"
                Write-Output "  XLS Path = $Script:xlsoutputfile2b"
                Write-Output ""
            }
            
            If ($ExportIndividual) {
                Write-output $Script:strLineSeparator
                Write-Output "  Export Path = $Script:ExportPath"

            }
    

        
        
        }
    }


    Start-Sleep -seconds 10
    ExitRoutine
  