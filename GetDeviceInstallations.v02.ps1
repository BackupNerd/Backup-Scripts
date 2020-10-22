Clear-Host
<## ----- About: ----
    # Get All Device Installations
    # Revision v02 - 2020-10-21
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
    # Authenticate to Backup.Management  (Supports the Standalone edition of SolarWinds Backup Only)
    # Get / Prompt for Partner Id
    # Enumerate Device Installations
    #   Incudes all Historic installation instances of a Device, including Backup, Restore Only, Bare-Metal Recovery, Recovery Console and Recovery Testing Instances
    #   Useful for Auding the last activity date for a specific installation Id 
    # Display as Grid-View
    # Output as CSV / XLS / AutoLaunch
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm 
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/API-column-codes.htm
# -----------------------------------------------------------#>  ## Behavior

# ----- Variables and Paths ----
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    Write-output "  Get All Device Installations"

# ----- End Variables and Paths ----

# ----- Functions ----

    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator 
            "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 

        $PartnerName = Read-Host -assecurestring "Enter EXACT Login Customer Name for SolarWinds Backup.Management API  " | convertfrom-securestring | out-file $APIcredfile
        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for SolarWinds Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
     
        Authenticate-Cookie 
    }  ## Set API credentials if not present

    Function Get-APICredentials {

        $Script:True_path = "C:\ProgramData\MXB\"
        $Script:APIcredfile = join-path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
        $Script:APIcredpath = Split-path -path $APIcredfile
        
        Write-Output $Script:strLineSeparator 
        Write-Output "  Getting Backup API Credentials" 
        if (Test-Path $APIcredfile) {
            Write-Output    $Script:strLineSeparator        
            "  Backup API Credential File Present"
            $APIcredentials = get-content $APIcredfile
            $Script:cred0 = $APIcredentials[0] | Convertto-SecureString
            $Script:cred0 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred0))
            $Script:cred1 = [string]$APIcredentials[1] 
            $Script:cred2 = $APIcredentials[2] | Convertto-SecureString
            $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred2))

            Write-Output    $Script:strLineSeparator 
            Write-output "  Stored Backup API Partner  = $cred0"
            Write-output "  Stored Backup API User     = $Script:cred1"
            Write-output "  Stored Backup API Password = Encrypted"
            
            Authenticate-Cookie

            }else{
            Write-Output    $Script:strLineSeparator 
            "  Backup API Credential File Not Present"
            Set-APICredentials
            
            } 
    }  ## Get API credentials if present
           
    Function Authenticate-Cookie {

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
    
    #Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"

    $Authenticate = $webrequest | convertfrom-json
    $Script:visa = $authenticate.visa
    }  ## Use Backup.Management credentials to Authenticate

    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $epoch = $epoch.AddSeconds($inputUnixTime)
        return $epoch
        }else{ return ""}
    }  ## Convert epoch time to date time 

    Function Get-PartnerInfo ($PartnerName) {      

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $visa
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
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
	        Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            $PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Get-PartnerInfo $partnername

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
        $data.params.query.SelectionMode = "PerInstallation"
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","LN","OP","OI","OS","PD","AP","PN","AA843","MN","TS","EI","IP","MO","MF","CD","VN","II","IM","RTG","RP")
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
    
    
        $Script:Devices = $webrequest | convertfrom-json
             
        $Script:DeviceDetail = @()

        Write-Output "  Requesting details for $($Devices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $Devices.result.result ) {

        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [String]$DeviceResult.AccountId;
                                                                    PartnerID      = [string]$DeviceResult.PartnerId;
                                                                    DeviceName     = $DeviceResult.Settings.AN -join '' ;
                                                                    PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                    DataSources    = $DeviceResult.Settings.AP -join '' ;                                                                
                                                                    Location       = $DeviceResult.Settings.LN -join '' ;
                                                                    Notes          = $DeviceResult.Settings.AA843 -join '' ;
                                                                    Product        = $DeviceResult.Settings.PN -join '' ;
                                                                    ProductID      = $DeviceResult.Settings.PD -join '' ;
                                                                    Profile        = $DeviceResult.Settings.OP -join '' ;
                                                                    OS             = $DeviceResult.Settings.OS -join '' ;
                                                                    MachineName    = $DeviceResult.Settings.MN -join '' ;  
                                                                    MFG_Name       = $DeviceResult.Settings.MF -join '' ;
                                                                    MFG_Model      = $DeviceResult.Settings.MO -join '' ;
                                                                    IP             = $DeviceResult.Settings.IP -join '' ;  
                                                                    Ext_IP         = $DeviceResult.Settings.EI -join '' ;
                                                                    Creation       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                                                                    TimeStamp      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;  
                                                                    ClientVersion  = $DeviceResult.Settings.VN -join '' ;
                                                                    InstallID      = $DeviceResult.Settings.II -join '' ;  
                                                                    InstallMode    = $DeviceResult.Settings.IM -join '' ;
                                                                    LastRestore    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.RTG -join '') ;  
                                                                    ProfileID      = $DeviceResult.Settings.OI -join '' }
        }     


    }  ## Enumerate devices under specified Sub-distributor or lower partner

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

# ----- End Functions ----

    Get-APICredentials

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Get-PartnerInfo $cred0

    $filter1 = "AT == 1"  ## Exclude M365 devices
   
    Get-Devices

    $DeviceDetail = $DeviceDetail | select-object PartnerID,PartnerName,AccountId,DeviceName,MachineName,Creation,LastRestore,TimeStamp,DataSources,ClientVersion,MFG_Name,MFG_Model,OS,InstallMode,InstallID,Ext_IP,IP,location,productID,Profile,Notes | Sort-Object Partnername,devicename,timestamp 

    ## Display Grid View

    $DeviceDetail | out-gridview -Title "Device Installation Audit"

    ## Export CSV

    $csvoutputfile = "$PSScriptRoot\$($CurrentDate)_Partner_$($PartnerId)_Device_Audit.csv"
    $DeviceDetail | Export-Csv -Path $csvoutputfile -NoTypeInformation

    ## Generate XLS

    $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
    Save-CSVasExcel $csvoutputfile
        
    Write-output $Global:strLineSeparator

    ## Launch CSV or XLS if Excel is installed
        
    If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
        Start-Process "$xlsoutputfile"
        }else{
        Start-Process "$csvoutputfile"
        }
    start-sleep -seconds 15

    