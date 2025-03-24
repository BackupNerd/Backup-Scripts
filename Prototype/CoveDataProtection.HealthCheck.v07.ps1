
[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 1000,                                ## Set maximum number of device / domain results to return
        [Parameter(Mandatory=$False)] [string]$ExportPath = $PSScriptRoot,                     ## Export base path, dated sub folders will be created
        [Parameter(Mandatory=$False)] [switch]$ExportCSV,                                      ## if $true, export CSV files
        [Parameter(Mandatory=$False)] [switch]$ExportXLSX = $true,                             ## if $true, export XLSX files
        [Parameter(Mandatory=$False)] [string]$delimiter = ",",                                ## Delimiter
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
Clear-Host

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Output "  This script requires PowerShell 7 or higher. Relaunching in PowerShell 7..."
    $pwshPath = "C:\Program Files\PowerShell\7\pwsh.exe"
    if (Test-Path $pwshPath) {
        & $pwshPath $MyInvocation.MyCommand.Path @args
        exit
    } else {
        Write-Error "PowerShell 7 is not installed. Please install PowerShell 7 and try again."
        exit
    }
}

#Requires -Version 5.1
$ConsoleTitle = "Cove Data Protection | The N-able Way Health Check"
$host.UI.RawUI.WindowTitle = $ConsoleTitle
if (-not $ExportPath) { $ExportPath = $PSScriptRoot }
$scriptpath = $MyInvocation.MyCommand.Path
Write-output "  $ConsoleTitle`n`n  $ScriptPath`n`n  Script Parameter Syntax:`n`n  $(Get-Command $PSCommandPath -Syntax)"
Push-Location (Split-Path $scriptpath)

$CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
#$ShortDate = Get-Date -format "yyyy-MM-dd"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------"

$urlJSON = 'https://api.backup.management/jsonapi'

$delimiter = (Get-Culture).TextInfo.ListSeparator

Write-output "  Current Parameters:"
Write-output "  -PowerShell Version             = $($PSVersionTable.PSVersion)"
Write-output "  -DeviceCount                    = $DeviceCount"
Write-output "  -ExportCSV                      = $ExportCSV"
Write-output "  -ExportXLSX                     = $ExportXLSX"
Write-output "  -ExportPath                     = $ExportPath"
Write-output "  -Delimiter                      = $delimiter"
Write-output "  -ClearCredentials               = $ClearCredentials"
Write-output $Script:strLineSeparator
#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove | Backup.Management API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

    $BackupCred.UserName | Out-file -append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
    
    Start-Sleep -milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    if (-not (Test-Path -Path $Script:True_path)) {
        New-Item -ItemType Directory -Path $Script:True_path -Force
    }
    $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-Path -Path $APIcredfile

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
                Write-output "  Stored Backup API Password = Encrypted`n"
                
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
    #$data.params.partner = $Script:cred0
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
        Write-Warning "`n  $($authenticate.error.message)"
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"

        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Test-CDPVisa{
     if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){   
            Send-APICredentialsCookie
        }
    }
}  ## Recheck remaining Visa time and reauthenticate

#endregion ----- Authentication ----

#region ----- Data Conversion ----

Function Measure-ElapsedTime {
    if (-not $script:StartTime) {
        $script:StartTime = Get-Date
    }else{
        $script:EndTime = Get-Date
        $elapsedTime = $script:EndTime - $script:StartTime
        $elapsedTime = "{0:hh\:mm\:ss}" -f $elapsedTime
        Write-Output "  Elapsed time: $elapsedTime`n"
    }
}

Function Convert-UnixTimeToDateTime($UnixToConvert) {
    if ($UnixToConvert -gt 0 ) { $Epoch2Date = ((Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()).AddSeconds($UnixToConvert)
    return $Epoch2Date }else{ return ""}
}  ## Convert epochtime to datetime #Rev.03

Function Convert-DateTimeToUnixTime($DateToConvert) {
    $Date2Epoch = (New-TimeSpan -Start (Get-Date -Date "1970-01-01 00:00:00Z")-End (Get-Date -Date $DateToConvert)).TotalSeconds
    Return $Date2Epoch
}  ## Convert datetime to epochtime #Rev.03

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

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

function Invoke-EnumerateColumns {
    param (
        [int]$partnerId
    )

    $headers = @{
        "Content-Type" = "application/json"
    }

    $body = @{
        jsonrpc = "2.0"
        id = "1"
        visa = $SCript:visa
        method = "EnumerateColumns"
        params = @{
            partnerId = $partnerId
        }
    } | ConvertTo-Json

    $Script:ColumnsResponse = Invoke-RestMethod 'https://api.backup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body
    return $ColumnsResponse | ConvertTo-Json
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
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json

    $RestrictedPartnerLevel = @("Root","SubRoot","Distributor")
    
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
        write-warning "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername

    }else{$script:visa = $Partner.visa}

} ## get PartnerID and Partner Level    

Function Send-EnumeratePartners ([int]$parentPartnerId) {
    Write-Output $Script:strLineSeparator
    Write-Output "  Enumerating Partners"
    $url = "https://api.backup.management/jsonapi"
    $data = @{
        jsonrpc = '2.0'
        visa = $Script:visa
        method = 'EnumeratePartners'
        params = @{
            parentPartnerId = $parentPartnerId
            fetchRecursively = $true
            fields = (0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25)
        }
        id = '1'
    } | ConvertTo-Json -Depth 5


    $headers = @{
        'Content-Type' = 'application/json; charset=utf-8'
        'Accept' = 'application/json; charset=utf-8'
    }


    $Script:EnumeratePartnersSession = Invoke-RestMethod -Method POST `
        -Uri $url `
        -Headers $headers `
        -Body $data `
        -ContentType 'application/json; charset=utf-8' `
        -ResponseHeadersVariable responseHeaders

    # Inspect the response headers
    $contentType = $responseHeaders.'Content-Type'
    Write-Output "  Content-Type: $contentType"


    $Script:visa = $EnumeratePartnersSession.visa
    if ($EnumeratePartnersSession.error) {
        Write-Output $Script:strLineSeparator
        Write-Output "  EnumeratePartnersSession Error Code:  $($EnumeratePartnersSession.error.code)"
        Write-Output "  EnumeratePartnersSession Message:  $($EnumeratePartnersSession.error.message)"
        Write-Output $Script:strLineSeparator
        Write-Output "  Exiting Script"
        Exit
    } else { # (No error)
        
        #$Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Parentid,@{l='PartnerId';e={($_.Id).tostring()}},Name,ExternalCode,Level,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
        
        $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | ForEach-Object {
            $obj = $_ | Select-Object -Property Parentid,@{l='PartnerId';e={($_.Id).tostring()}},@{l='PartnerName';e={($_.name).tostring()}},ExternalCode,Level,LocationId,* -ExcludeProperty Company,advancedpartnerproperties,ExternalPartnerProperties,ChildServiceTypes,Flags,name -ErrorAction Ignore
            #$obj | Add-Member -MemberType NoteProperty -Name 'Company' -Value ($_.Company | Out-String)
            $obj | Add-Member -MemberType NoteProperty -Name 'AdvancedPartnerProperties' -Value ($_.AdvancedPartnerProperties | Out-String)
            $obj | Add-Member -MemberType NoteProperty -Name 'ExternalPartnerProperties' -Value ($_.ExternalPartnerProperties | Out-String)
            $obj | Add-Member -MemberType NoteProperty -Name 'ChildServiceTypes' -Value ($_.ChildServiceTypes | Out-String)
            $obj | Add-Member -MemberType NoteProperty -Name 'Flags' -Value ($_.Flags | Out-String)
            $obj
        }
        
        $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
        $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
        $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
    
        $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                
        $Script:SelectedPartners = $Script:SelectedPartners += @( [pscustomobject]@{PartnerName=$PartnerName;Id=[string]$PartnerId;Level="<TOPLEVEL>";PartnerId=[string]$PartnerId} ) 
                        
        #$script:Selection = $Script:SelectedPartner | Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid 


        $Script:cleanpartnername = $Script:partnername -replace '[^a-zA-Z0-9]', '_' ## Clean Partner Name for Export File Name

        if ($ExportPath) {
            $Script:ExportPath = Join-Path -Path $ExportPath -ChildPath "HealthCheck_$($cleanpartnername)_$($partnerid)"
        } else {
            $Script:ExportPath = Join-Path -Path $PSScriptRoot -ChildPath "HealthCheck_$($cleanpartnername)_$($partnerid)"
        } ## Set the export path based on the provided ExportPath parameter or default to script directory
        
        if ($Script:ExportPath) {
            mkdir -Force -Path $Script:ExportPath | Out-Null 
        } ## Create the export directory if it doesn't exist
        
        $ConsoleTitle = "Cove Data Protection | The N-able Way Health Check | $Partnername" ## Update console title with partner name
        $host.UI.RawUI.WindowTitle = $ConsoleTitle
        
        if ($ExportCSV) {
            write-output $Script:strLineSeparator
            write-output "  Exporting Partner Data to CSV"
            $Script:SelectedPartners | export-csv -Path "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Partners.csv" -NoTypeInformation -Delimiter $delimiter -Encoding utf8
        } ## Export the data to a CSV file

        if ($ExportXLSX) {
            write-output $Script:strLineSeparator
            write-output "  Exporting Partner Data to Excel"
            $Script:SelectedPartners | Export-Excel -Path "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx" -WorksheetName "Partners" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
        } ## Export the data to an Excel file
        
    }
}  ## EnumeratePartners API Call

Function Send-EnumerateUsers { 

    Write-Output $Script:strLineSeparator
    Write-Output "  Enumerating Users"

    [int32[]]$CID = @($selectedpartners.id)

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'EnumerateUsers'
    $data.params = @{}
    $data.params.partnerIds = $cid

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:websession = $websession
        $Script:EnumerateUsers = $webrequest | convertfrom-json

        #$Script:Users = $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,*,phonenumber -ErrorAction SilentlyContinue

        $Script:Users = $Script:EnumerateUsers.result.result | Select-Object @{l='ParentId';e={""}},PartnerId,@{l='PartnerName';e={""}},Name,emailAddress,FirstName,FullName,id,Roleid,Title,PhoneNumber,firstlogintime,lastlogintime,* -ExcludeProperty password -ErrorAction SilentlyContinue | ForEach-Object {
            $_.firstlogintime = Convert-UnixTimeToDateTime $_.firstlogintime
            $_.lastlogintime = Convert-UnixTimeToDateTime $_.lastlogintime
            $_ | Add-Member -MemberType NoteProperty -Name SecurityOfficer -Value ($_.Flags -contains "SecurityOfficer")
            $_ | Add-Member -MemberType NoteProperty -Name APIAuthentication -Value ($_.Flags -contains "AllowApiAuthentication")
            $_ | Add-Member -MemberType NoteProperty -Name APIOnlyAccess -Value ($_.Flags -contains "NonInteractive")
            $_.Roleid = switch ($_.Roleid) {
            1 { "Superuser" }
            2 { "Administrator" }
            3 { "Manager" }
            4 { "Operator" }
            5 { "Supporter" }
            6 { "Reporter" }
            7 { "Notifier" }
            default { $_.Roleid }
            }
            $_
        }
        if ($ExportCSV) {
            write-output $Script:strLineSeparator
            write-output "  Exporting User Data to CSV"
            $Script:users | Export-Csv -Path "$ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Users.csv" -NoTypeInformation -Delimiter $delimiter
        } ## Export the data to a CSV file

        if ($ExportXLSX) {
            write-output $Script:strLineSeparator
            write-output "  Exporting User Data to Excel"
            $Script:users | Export-Excel -Path "$ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx" -WorksheetName "Users" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
        } ## Export the data to an Excel file
} ## EnumerateUsers API Call

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


Function Send-EnumerateAccountProfiles {

    Write-Output $Script:strLineSeparator
    Write-Output "  Enumerating Profiles"

    $Script:FlattenedProfiles = @()

    foreach ($SelectedPartner in $Script:SelectedPartners | sort-object PartnerName ) {
        Write-output " Enumerating Profiles for $($SelectedPartner.PartnerName)"

        # ----- Get Profiles via EnumerateAccountProfiles -----
        $objEnumerateAccountProfiles = @{
            jsonrpc = '2.0'
            visa = $Script:visa
            method = 'EnumerateAccountProfiles'
            params = @{
                partnerId = [int]$SelectedPartner.id
                fetchRecursively = $false
             }
            id = '1'
        } | ConvertTo-Json -Depth 3

        $script:EnumerateAccountProfilesSession = CallJSON $urlJSON $objEnumerateAccountProfiles

        #Start-Sleep -Milliseconds 100

        if ($EnumerateAccountProfilesSession.error) {
            Write-Output $Script:strLineSeparator
            Write-Output "  EnumerateAccountProfilesSession Error Code: $($EnumerateAccountProfilesSession.error.code)"
            Write-Output "  EnumerateAccountProfilesSession Message: $($EnumerateAccountProfilesSession.error.message)"
            Write-Output $Script:strLineSeparator
            Write-Output "  Exiting Script"
            return
        }
            Else {
        # (No error)
                $EnumerateAccountProfiles = $EnumerateAccountProfilesSession.result.result | Sort-Object Partnerid -Descending | Where-Object { $_.id -notin $Script:FlattenedProfiles.id }

                if ($EnumerateAccountProfiles.id.count -ge 1) {
                    Write-Host " -->  $($EnumerateAccountProfiles.id.count) unique profiles found" -ForegroundColor Green
                }

                foreach ($profile in $EnumerateAccountProfiles) {
         

                    $flattenedProfile = [PSCustomObject]@{
                        ParentId            = ""
                        PartnerId           = $profile.PartnerId
                        PartnerName         = ""
                        ID                  = $profile.ID
                        Name                = $profile.Name
                        Version             = $profile.version
                        AutoUpdatePolicy    = $profile.profiledata.AutoUpdatePolicy
                        AutomaticSelection  = $profile.profiledata.AutomaticSelection
                        BackupSchedule      = $profile.profiledata.BackupSchedule
                        WorkingHours        = $profile.profiledata.HighFrequentBackupSchedule.WorkingHours.days
                        HighFrequentBackupSchedule = $profile.profiledata.HighFrequentBackupSchedule
                        BackupScheduleItems = $profile.profiledata.HighFrequentBackupSchedule.BackupScheduleItems
                        Language            = $profile.profiledata.Language
                        TemporaryFolderPath = $profile.profiledata.TemporaryFolderPath
                    }



                    foreach ($dataSource in $profile.profiledata.BackupDataSourceSettings) {
                        $flattenedProfile | Add-Member -MemberType NoteProperty -Name ($dataSource.DataSource) -Value $dataSource.DataSource
                        $flattenedProfile | Add-Member -MemberType NoteProperty -Name ($dataSource.DataSource + " Selections") -Value $dataSource.SelectionCollection.selection
                        $flattenedProfile | Add-Member -MemberType NoteProperty -Name ($dataSource.DataSource + " Exclusions") -Value $dataSource.ExclusionFilter
                        $flattenedProfile | Add-Member -MemberType NoteProperty -Name ($dataSource.DataSource + " Policy") -Value $dataSource.Policy
                    }
                    
         

                    $Script:FlattenedProfiles += $flattenedProfile
                }
            }
        }
        # Create a template object with all columns in $Script:FlattenedProfiles dynamically
        $TemplateObject = [PSCustomObject]@{}
        $Script:FlattenedProfiles | ForEach-Object {
            $_.PSObject.Properties.Name | ForEach-Object {
                if (-not $TemplateObject.PSObject.Properties[$_]) {
                    $TemplateObject | Add-Member -MemberType NoteProperty -Name $_ -Value $null
                }
            }
        }

        # Add the template object to the flattened profiles to ensure all columns are present
        $Script:FlattenedProfiles += $TemplateObject

        if ($ExportCSV) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Exporting Profile Data to CSV"            
            $Script:FlattenedProfiles | Sort-Object id -Unique |  Select-Object * -ExcludeProperty BackupSchedule,WorkingHours,HighFrequentBackupSchedule,BackupScheduleItems  | Export-Csv -Path "$ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Profiles.csv" -NoTypeInformation -Delimiter $delimiter
        } ## Export the data to a CSV file

        if ($ExportXLSX) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Exporting Profile Data to Excel"
            $Script:FlattenedProfiles | Sort-Object id -Unique |  Select-Object * -ExcludeProperty BackupSchedule,WorkingHours,HighFrequentBackupSchedule,BackupScheduleItems | Export-Excel -Path "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx" -WorksheetName "Profiles" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
        } ## Export the data to an Excel file
    

} ## EnumerateAccountProfiles API Call

Function Send-EnumerateProducts {
    # ----- Get Products via EnumerateProducts -----

    Write-Output $Script:strLineSeparator
    Write-Output "  Enumerating Products"

    $Script:FlattenedProducts = @()
    Foreach ($SelectedPartner in $selectedPartners | sort-object PartnerName ) {
        Write-output " Enumerating Products for $($SelectedPartner.PartnerName)"

        $objEnumerateProducts = @{
            jsonrpc = '2.0'
            visa = $Script:visa
            method = 'EnumerateProducts'
            params = @{ 
                partnerId = [int]$SelectedPartner.id 
            }
            id = '1'
        } | ConvertTo-Json -Depth 3

        $Script:EnumerateProductsSession = CallJSON $urlJSON $objEnumerateProducts
        $Script:visa = $EnumerateProductsSession.visa

        #Start-Sleep -Milliseconds 100

        if ($EnumerateProductsSession.error) {
            Write-Output $Script:strLineSeparator
            Write-Output "  EnumerateProductsSession Error Code: $($EnumerateProductsSession.error.code)"
            Write-Output "  EnumerateProductsSession Message: $($EnumerateProductsSession.error.message)"
            Write-Output $Script:strLineSeparator
            Write-Output "  Exiting Script"
            return
        } else {
            # (No error)
            $Script:EnumerateProductsSessionResults = $EnumerateProductsSession.result.result | Sort-Object id | Where-Object { $_.id -notin $Script:FlattenedProducts.productid }
   
            if ($Script:EnumerateProductsSessionResults.id.count -ge 1) {
                Write-Host " -->  $($Script:EnumerateProductsSessionResults.id.count) unique products found" -ForegroundColor Green
            }

            foreach ($product in $Script:EnumerateProductsSessionResults) {
                $flattenedProduct = [PSCustomObject]@{
                    ParentId    = ""
                    PartnerId   = $product.PartnerId
                    PartnerName = ""
                    ProductId   = $product.Id
                    ProductName = $product.Name
                }

                foreach ($feature in $product.features) {
                    $flattenedProduct | Add-Member -MemberType NoteProperty -Name $feature[0] -Value $feature[1]
                }

                $Script:FlattenedProducts += $flattenedProduct
            }
            # Create a template object with all columns in $Script:FlattenedProducts dynamically
            $TemplateObject = [PSCustomObject]@{}
            $Script:FlattenedProducts | ForEach-Object {
                $_.PSObject.Properties.Name | ForEach-Object {
                    if (-not $TemplateObject.PSObject.Properties[$_]) {
                        $TemplateObject | Add-Member -MemberType NoteProperty -Name $_ -Value $null
                    }
                }
            }

            # Add the template object to the flattened products to ensure all columns are present
           
        }

    }
            $Script:FlattenedProducttemplate = @()
            $Script:FlattenedProducttemplate += $TemplateObject
            $Script:FlattenedProducttemplate += $Script:FlattenedProducts | Sort-Object productid -Unique

            if ($ExportCSV) {
                Write-Output $Script:strLineSeparator
                Write-Output "  Exporting Product Data to CSV"
                $Script:FlattenedProducts | Export-Csv -Path "$ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Products.csv" -NoTypeInformation -Delimiter $delimiter
            } ## Export the data to a CSV file
            
            if ($ExportXLSX) {
                Write-Output $Script:strLineSeparator
                Write-Output "  Exporting Product Data to Excel"
                $Script:FlattenedProducts |  Export-Excel -Path "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx" -WorksheetName "Products" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
            } ## Export the data to an Excel file
            
}
    
Function Send-EnumerateAccountStatistics {
# ----- Get Devices via EnumerateAccountStatistics -----

Write-Output $Script:strLineSeparator
Write-Output "  Enumerating Statistics"

$Filter = ""
$Columns = ('AT','AU','AN','AR','MN','I78','OT','I81','CD','I82','RU','PN','PD','OP','OI','T0','TL','T7','T3','US','I80','AS','VN','F5','FA','F3','T5','YV','OS','AA843','TS','HN','HI','H3','H5','HA','EN','W3','W5','WA','WI','TM','D9F27','GM','JM','D5F20','D23F20','AA3147','AA3148','AA3149','AA3150')

$Properties = @("ParentID","PartnerID","PartnerName","Reference","AccountID","DeviceName","MachineName","AccountType","OsType","OS","Version","Physicality","Datasources","Passphrase","Recovery","Retention","Product","ProductID","Profile","ProfileID","Status","Errors","LSVstatus","Creation","TotalLast","TimeStamp","Notes","UsedGB","ArchivedGB","SelectedGB","SentGB","FileSelGB","FileSentGB","FileDuration","HVvms","HVlic","HVSelGB","HVSentGB","HVDuration","VMvms","VMlic","VMWSelGB","VMWSentGB","VMWDuration","M365Bill","M365Off","EXUsers","ODUsers","SPUsers","TMOwners","EXAutoAdd","ODAutoAdd","SPAutoAdd","TMAutoAdd")

$ReplaceStatus = @{
    'Never' = 'NoBackup (o)'
    '' = 'NoBackup (o)'
    '0' = 'NoBackup (o)'
    '1' = 'InProcess (>)'
    '2' = 'Failed (-)'
    '3' = 'Aborted (x)'
    '4' = 'Unknown (?)'
    '5' = 'Completed (+)'
    '6' = 'Interrupted (&)'
    '7' = 'NotStarted (!)'
    '8' = 'CompletedWithErrors (#)'
    '9' = 'InProgressWithFaults (%)'
    '10' = 'OverQuota ($)'
    '11' = 'NoSelection (0)'
    '12' = 'Restarted (*)'
} ## Replace numeric status codes with description and Ascii char.

# Create the JSON object to call the EnumerateAccountStatistics function
$objEnumerateAccountStatistics = @{
    jsonrpc = '2.0'
    visa = $Script:visa
    method = 'EnumerateAccountStatistics'
    params = @{
        query = @{
            PartnerId = $partnerid
            Filter = $filter
            Columns = $columns
            StartRecordNumber = 0
            RecordsCount = $DeviceCount
            Totals = @("COUNT(AT==1)", "SUM(T3)", "SUM(US)")
        }
    }
    id = '1'
} | ConvertTo-Json -Depth 4

# Call the JSON Web Request Function to get the EnumerateAccountStatistics Object
$script:response = Invoke-RestMethod -Uri $urlJSON -Method 'POST' -ContentType 'application/json' -Body $objEnumerateAccountStatistics

$Script:visa = $response.visa

# Get Result Status of EnumerateAccountStatistics
$SessionErrorCode = $response.error.code
$SessionErrorMsg = $response.error.message

# Check for Errors with EnumerateAccountStatistics - Check if ErrorCode has a value
if ($SessionErrorCode) {
    Write-Output $Script:strLineSeparator
    Write-Output "  EnumerateAccountStatistics Error Code:  $SessionErrorCode"
    Write-Output "  EnumerateAccountStatistics Message:  $SessionErrorMsg"
    Write-Output $Script:strLineSeparator
    Write-Output "  Exiting Script"
    return
}
Else {
# (No error)

}
    
$Script:DeviceDetail = @()
ForEach ( $Result in $Script:response.result.result ) {
    $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ 
        ParentId    = "" ;
        PartnerId   = [string]$Result.PartnerId ;
        PartnerName = $Result.Settings.AR -join '' ;
        Reference   = $Result.Settings.PF -join '' ;
        DeviceId    = $Result.Settings.AU -join '' ;
        DeviceName  = $Result.Settings.AN -join '' ;
        MachineName = $Result.Settings.MN -join '' ; 
        AccountType = ($Result.Settings.AT -join '' ).replace("1","Backup Manager").replace("2","Cloud2Cloud") ; 
        OsType      = ($Result.Settings.OT -join '' ).replace("1","Workstation").replace("2","Server")  ; 
        OS          = $Result.Settings.OS -join '' ;   
        Version     = $Result.Settings.VN -join '' ;
        Physicality = $Result.Settings.I81 -join '' ;
        Datasources = ($Result.Settings.I78 -join '' ).replace("D01","Fs ").replace("D02","Ss ").replace("D10","Mssql ").replace("D04","Exch ").replace("D06","Network ").replace("D08","VMware ").replace("D14","Hyperv ").replace("D15","Mysql ").replace("D19","365Ex ").replace("D20","365OD ").replace("D23","365TM ").replace("D05","365SP ") ;
        Passphrase  = $Result.Settings.I82 -join '' ;
        Recovery    = ($Result.Settings.I80 -join '' ).replace("1","RecoveryTesting").replace("2","StandbyImage") ;
        Retention   = $Result.Settings.RU -join '' ;
        Product     = $Result.Settings.PN -join '' ; 
        ProductID   = $Result.Settings.PD -join '' ;     
        Profile     = $Result.Settings.OP -join '' ;
        ProfileID   = $Result.Settings.OI -join '' ;
        Status      = $Result.Settings.T0 -join '' ;
        Errors      = $Result.Settings.T7 -join '' ;
        LSVstatus   = $Result.Settings.YV -join '' ;
        Creation    = (Convert-UnixTimeToDateTime ($Result.Settings.CD -join '')) ;
        TotalLast   = (Convert-UnixTimeToDateTime ($Result.Settings.TL -join '')) ;
        TimeStamp   = (Convert-UnixTimeToDateTime ($Result.Settings.TS -join '')) ;
        Notes       = $Result.Settings.AA843 -join '' ;
        UsedGB      = [math]::Round([Decimal]($Result.Settings.US -join '') / 1GB, 2) ;
        ArchivedGB  = [math]::Round([Decimal]($Result.Settings.AS -join '') / 1GB, 2) ;
        SelectedGB  = [math]::Round([Decimal]($Result.Settings.T3 -join '') / 1GB, 2) ;
        SentGB      = [math]::Round([Decimal]($Result.Settings.T5 -join '') / 1GB, 2) ;
        FileSelGB   = [math]::Round([Decimal]($Result.Settings.F3 -join '') / 1GB, 2) ;
        FileSentGB  = [math]::Round([Decimal]($Result.Settings.F5 -join '') / 1GB, 2) ;
        FileDuration= $Result.Settings.FA -join '' ;
        HVvms       = $Result.Settings.HN -join '' ;
        HVlic       = $Result.Settings.HI -join '' ;
        HVSelGB     = [math]::Round([Decimal]($Result.Settings.H3 -join '') / 1GB, 2) ;
        HVSentGB    = [math]::Round([Decimal]($Result.Settings.H5 -join '') / 1GB, 2) ;
        HVDuration  = $Result.Settings.HA -join '' ;
        VMvms       = $Result.Settings.EN -join '' ;
        VMlic       = $Result.Settings.WI -join '' ;
        VMWSelGB    = [math]::Round([Decimal]($Result.Settings.W3 -join '') / 1GB, 2) ;
        VMWSentGB   = [math]::Round([Decimal]($Result.Settings.W5 -join '') / 1GB, 2) ;
        VMWDuration = $Result.Settings.WA -join '' ;
        M365Bill    = $Result.Settings.TM -join '' ;
        M365Off     = $Result.Settings.D9F27 -join '' ;
        EXUsers     = $Result.Settings.GM -join '' ;
        ODUsers     = $Result.Settings.JM -join '' ;
        SPUsers     = $Result.Settings.D5F20 -join '' ;
        TMOwners    = $Result.Settings.D23F20 -join '' ;
        EXAutoAdd   = $Result.Settings.AA3147 -join '' ;
        ODAutoAdd   = $Result.Settings.AA3148 -join '' ;
        SPAutoAdd   = $Result.Settings.AA3149 -join '' ;
        TMAutoAdd   = $Result.Settings.AA3150 -join '' ;
       }
}
 
# Replace status codes in the DeviceDetail objects
$Script:DeviceDetail = $Script:DeviceDetail | ForEach-Object {
    $_.Status = $ReplaceStatus[$_.Status]
    $_
}

# (Summarize DeviceDetail)

if ($ExportCSV) {
    Write-Output $Script:strLineSeparator
    Write-Output "  Exporting Device Data to CSV"
    $Script:DeviceDetail | Select-Object $properties | Export-Csv -Path "$ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Devices.csv" -NoTypeInformation -Delimiter $delimiter
} ## Export the data to a CSV file

if ($ExportXLSX) {
    Write-Output $Script:strLineSeparator
    Write-Output "  Exporting Device Data to Excel" 
    $Script:DeviceDetail | Select-Object $properties | Export-Excel -Path "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx" -WorksheetName "Devices" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    
} ## Export the data to an Excel file

} ## EnumerateAccountStatistics API Call

Function Install-RequiredModules {
    param (
        [string[]]$Modules
    )
    foreach ($Module in $Modules) {
        if (Get-Module -ListAvailable -Name $Module) { ## Check if the module is already installed
            Write-Host "  Module $Module Already Installed" ## Inform the user that the module is already installed
        } 
        else {
            try {
                Install-Module -Name $Module -Confirm:$True -Force ## Attempt to install the module
            }
            catch [Exception] {
                $_.message ## Output the exception message
                Write-Warning "  Module '$Module' not found, run with Administrator rights to install" ## Warn the user if the module is not found and suggest running with admin rights
                exit ## Exit the script
            }
        }
    }
}

Function ExitRoutine {
    Write-Output $Script:strLineSeparator
    Write-Output "  Secure credential file found here:"
    Write-Output $Script:strLineSeparator
    Write-Output "  & $APIcredfile"
    Write-Output ""
    Write-Output $Script:strLineSeparator
    Start-Sleep -Seconds 3
}

Function Add-Summary {

    $xlSourcefile = "$Script:ExportPath\$($cleanpartnername)_$($partnerid)_$($CurrentDate)_Combined.xlsx"

    $shortDate | Export-Excel -Path $xlSourcefile -WorksheetName "Cove360" -AutoSize -FreezeTopRow -BoldTopRow
     
    $excel = Open-ExcelPackage $xlSourcefile
     
    # Access the Cove360 sheet
    $sheet = $excel.Workbook.Worksheets["Cove360"]
     
    $sheet.View.ShowGridLines = $true
    $sheet.View.ShowHeaders = $true
     
    # Set light grey background for column K
    Set-ExcelRange -Address $sheet.Cells["K:K"] -BackgroundColor LightGray
     
    # Set formats of cells
    Set-ExcelRange -Address $sheet.Cells["C:C"] -WrapText -HorizontalAlignment Left
    Set-ExcelRange -Address $sheet.Cells["D:D"] -NumberFormat "#.#0%"  -WrapText -HorizontalAlignment Center
     
    Set-ExcelRange -Address $sheet.Cells["E:E"] -NumberFormat "$#,##0" -WrapText -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["F:F"] -NumberFormat "#.#0%"  -WrapText -HorizontalAlignment Center
     
    Set-ExcelRange -Address $sheet.Cells["G:H"] -WrapText -HorizontalAlignment Center
     
    Set-ExcelRange -Address $sheet.Cells["A:A"] -Width 10
    Set-ExcelRange -Address $sheet.Cells["B:B"] -Width 6
    Set-ExcelRange -Address $sheet.Cells["C:C"] -Width 30
    Set-ExcelRange -Address $sheet.Cells["D:D"] -Width 20
    Set-ExcelRange -Address $sheet.Cells["E:E"] -Width 20
    Set-ExcelRange -Address $sheet.Cells["F:F"] -Width 10
    Set-ExcelRange -Address $sheet.Cells["G:G"] -Width 10
    Set-ExcelRange -Address $sheet.Cells["H:H"] -Width 10
    Set-ExcelRange -Address $sheet.Cells["I:I"] -Width 12
    Set-ExcelRange -Address $sheet.Cells["J:J"] -Width 12
    Set-ExcelRange -Address $sheet.Cells["K:K"] -Width 16
    Set-ExcelRange -Address $sheet.Cells["L:L"] -Width 16
     
    # $BorderBottom = "Thick"
    # $BorderColor = "Black"
     
    # Set the third row to dark grey background with white font
    Set-ExcelRange -Address $sheet.Cells["3:3"] -BackgroundColor DarkGray -FontColor White -Height 20
     
    # Insert current timestamp in cell A1 with red font
    Set-ExcelRange -Address $sheet.Cells["A1"] -Value 'Generated:' -FontColor DarkGray
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Set-ExcelRange -Address $sheet.Cells["B1"] -Value $timestamp -FontColor DarkGray
     
    # Insert a sum formula in cell L1 to sum all values in column J
    Set-ExcelRange -Address $sheet.Cells["L1"] -Formula '=COUNTIF(J:J,"yes")' -Bold -HorizontalAlignment Center
     
    # Insert formula in J8 to check if sum of K9:K12 > 0
    #Set-ExcelRange -Address $sheet.Cells["J8"] -Formula "=IF(SUM(K9:K12)>0; \"yes\"; \"no\")" -Bold -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["J8"] -Formula '=IF(SUM(K9:K12)>0,"yes","no")' -Bold -HorizontalAlignment Center
   
     
    Set-ExcelRange -Address $sheet.Cells["K1"] -HorizontalAlignment Center -Bold -Value "Advisory Topics"
    Set-ExcelRange -Address $sheet.Cells["K2"] -HorizontalAlignment Center -Bold -Value "Advisory Items"
     
    Set-ExcelRange -Address $sheet.Cells["C5"] -Value 'Total users'
    Set-ExcelRange -Address $sheet.Cells["C6"] -Value 'Total users top level'
     
    Set-ExcelRange -Address $sheet.Cells["D8"] -Value 'Superuser' -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["E8"] -Value 'Administrator' -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["F8"] -Value 'Manager' -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["G8"] -Value 'Operator' -HorizontalAlignment Center
    Set-ExcelRange -Address $sheet.Cells["H8"] -Value 'Supporter' -HorizontalAlignment Center
     
     
    Set-ExcelRange -Address $sheet.Cells["C9"] -Value 'Top level users split by role'
    Set-ExcelRange -Address $sheet.Cells["C10"] -Value 'Security Officers'
    Set-ExcelRange -Address $sheet.Cells["C11"] -Value '2FA enabled'
    Set-ExcelRange -Address $sheet.Cells["C12"] -Value 'API Access'
     
     
    #Set-ExcelRange -Address $sheet.Cells["E10"] -Formula "=Sum(E3:E8)" -Bold
    #Set-ExcelRange -Address $sheet.Cells["I10"] -Formula "=Sum(I3:I8)" -Bold
    #Set-ExcelRange -Address $sheet.Cells["M10"] -Formula "=Sum(M3:M8)" -Bold
    #Set-ExcelRange -Address $sheet.Cells["O10"] -Formula "=Sum(O3:O8)" -Bold
     
    Close-ExcelPackage $excel -Show
     
     

}





## Main Script
Measure-ElapsedTime ## Log the elapsed time
Install-RequiredModules -Modules @('Join-Object', 'ImportExcel') ## Ensure required modules are installed

Send-APICredentialsCookie ## Authenticate to API
Send-GetPartnerInfo $Script:cred0           ## Get Partner Info using the stored credentials

Measure-ElapsedTime ## Log the elapsed time
Send-EnumeratePartners $Script:PartnerId   ## Enumerate Partners


Measure-ElapsedTime ## Log the elapsed time

Send-EnumerateAccountStatistics ## Enumerate account statistics
Measure-ElapsedTime ## Log the elapsed time

Send-EnumerateUsers ## Enumerate users
Measure-ElapsedTime ## Log the elapsed time

Send-EnumerateAccountProfiles ## Enumerate account profiles
Measure-ElapsedTime ## Log the elapsed time

Send-EnumerateProducts ## Enumerate products
Measure-ElapsedTime ## Log the elapsed time

Add-Summary ## Add summary to the Excel file
Measure-ElapsedTime ## Log the elapsed time


ExitRoutine ## Exit the script






