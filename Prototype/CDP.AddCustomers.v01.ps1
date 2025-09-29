<# ----- About: ----
    # N-able Cove Data Protection | Add Customers
    # Revision v11 - 2025-09-29
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
    # For use with Cove Data Protection from N-able
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Manage Cove Data Protection Customers via JSON API
    # Use the -Import switch parameter to add users via CLI parameters
    # Use the -ImportPath (?:\Folder) parameter to specify CSV path and filename
    # Sample CSV file format: 
        companyname,ParentId,Level,State,LocationId,DeviceCountry
        Acme Corp 1,229433,EndCustomer,InTrial,18,US
        Beta LLC 2,229433,EndCustomer,InTrial,18,US
        Gamma Inc 3,229433,EndCustomer,InTrial,18,US
        #
    # Use the -Delimiter (default=',') parameter to set the delimiter for log output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm

    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm
    #
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)] [Switch]$Import,                  ## Add user via CLI parameters
    [Parameter(Mandatory=$False)] [String]$ImportPath = "$PSScriptRoot\ImportUsers.csv",  ## Import Users from CSV file
    [Parameter(Mandatory=$False)] [String]$Delimiter = ',',         ## specify ',' or ';' Delimiter for XLS & CSV file   
    [Parameter(Mandatory=$False)] [Switch]$ClearCredentials         ## Remove Stored API Credentials
)

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    
    Write-output "  Import Customers`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
    Write-output "  Current Parameters:"
   
    
    Write-output "  -Import        = $Import"
    Write-output "  -ImportPath    = $ImportPath"
    
    Write-output "  -Delimiter     = $Delimiter"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $urljson = "https://api.backup.management/jsonapi"
    $host.UI.RawUI.WindowTitle = "Manage Cove Users - $($PsCmdlet.ParameterSetName)"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 

        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove Backup.Management API'
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
$data.params.username = $Script:cred1
$data.params.password = $Script:cred2

$webrequest = Invoke-RestMethod -Method POST `
    -ContentType "application/json; charset=utf-8" `
    -Body (ConvertTo-Json $data) `
    -Uri $url `
    -SessionVariable Script:websession `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:Authenticate = $webrequest

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

if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
    Remove-Item -Path $Script:APIcredfile
    $ClearCredentials = $Null
    Write-Output $Script:strLineSeparator 
    Write-Output "  Backup API Credential File Cleared"
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

    $RestrictedPartnerLevel = @("Root","SubRoot")

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
        Write-Host "  Lookup for $($Partner.result.result.Level) level partner not allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive N-able | Cove Backup.Management Partner Name i.e. 'AcmeIT (bob@acmeit.org)' / N-central Activation Id i.e. '0015000000YXXXXAAX'"
        Send-GetPartnerInfo $Script:partnername
        }

    if ($partner.error) {
        write-output "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive N-able | Cove Backup.Management Partner Name i.e. 'AcmeIT (bob@acmeit.org)' / N-central Activation Id i.e. '0015000000YXXXXAAX'"
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
        
            $Script:SelectedPartners = @()

            $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'} #| Where-object {$_.level -eq "Reseller"}
            
            $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
            
            if ($SubPartners) {

                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $Script:partnername | Please select a partner" -OutputMode Single

                if($null -eq $Selection) {
                    # Cancel was pressed
                    # Run cancel script
                    #Write-Output    $Script:strLineSeparator
                    Write-Output    "  No Partner/s selected"
                    Break
                }
                else {
                    # OK was pressed, $Selection contains what was chosen
                    # Run OK script
                    [int]$script:PartnerId = $script:Selection.Id
                    [String]$script:PartnerName = $script:Selection.Name
                    Write-output "  Selected Partner = $PartnerName | $PartnerId"
                }

            }else{

                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                Write-Output    $Script:strLineSeparator
                Write-Output    "  Top Partner Selected"

            }
    }
    
}  ## EnumeratePartners API Call


Function Send-AddPartner {
    param (
        [string]$companyname,
        [int]$ParentId = 229433,
        [string]$Level = "EndCustomer",
        [ValidateSet("InProduction", "InTrial")] [string]$State = "InTrial",
        [int]$LocationId = 18,
        [string]$DeviceCountry = "US"
    )

    $url = "https://api.backup.management/jsonapi"
    $body = @{
        jsonrpc = "2.0"
        id      = "1"
        visa    = $script:visa
        method  = "AddPartner"
        params  = @{
            partnerInfo = @{
                ParentId          = $ParentId
                Level             = $Level
                ServiceType       = "AllInclusive"
                ChildServiceTypes = @("AllInclusive")
                Name              = $companyname
                State             = $State
                LocationId        = $LocationId
                DeviceCountry     = $DeviceCountry
                Company           = @{
                    LegalCompanyName = $companyname
                    PostAddress      = @{
                        Country = $DeviceCountry
                    }
                }
            }
        }
    }

    $response = Invoke-RestMethod -Uri $url -Method POST -ContentType "application/json" -Body ($body | ConvertTo-Json -Depth 5)
    return $response
}

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

Send-APICredentialsCookie
Send-GetPartnerInfo $Script:cred0
Send-EnumeratePartners 

    $Script:NewCustomers = Import-Csv -Path $ImportPath -Delimiter $Delimiter

    # Validate required columns in the CSV
    $requiredColumns = @('companyname', 'ParentId', 'Level', 'State', 'LocationId', 'DeviceCountry')
    $missingColumns = @()

    foreach ($column in $requiredColumns) {
        if (-not ($Script:NewCustomers | Get-Member -Name $column -MemberType NoteProperty)) {
            $missingColumns += $column
        }
    }

    if ($missingColumns.Count -gt 0) {
        Write-Warning "The following required columns are missing in the CSV file: $($missingColumns -join ', ')"
        Break
    }

    foreach ($Script:NewCustomer in $Script:NewCustomers) {
        $Script:companyname = $Script:NewCustomer.companyname
        $Script:ParentId = $Script:NewCustomer.ParentId
        $Script:Level = $Script:NewCustomer.Level
        $Script:State = $Script:NewCustomer.State
        $Script:LocationId = $Script:NewCustomer.LocationId
        $Script:DeviceCountry = $Script:NewCustomer.DeviceCountry

        Send-AddPartner -companyname $NewCustomer.companyname `
                        -ParentId $NewCustomer.ParentId `
                        -Level $NewCustomer.Level `
                        -State $NewCustomer.State `
                        -LocationId $NewCustomer.LocationId `
                        -DeviceCountry $NewCustomer.DeviceCountry
    }

Start-Sleep -seconds 3
Read-Host "  Press ENTER to exit..."
Exit

