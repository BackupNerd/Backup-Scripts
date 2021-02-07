<# ----- About: ----
    # Get SW Backup Add Notifier User 
    # Revision v03 - 2021-02-06
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
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate Notifier Users
    # List/ Add/ Remove User
    # Optionally export to XLS/CSV
    # Use the -AllPartners switch parameter to skip GUI partner selection (pending addition)
    # Use the -Export switch parameter to export statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/API-column-codes.htm 
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="ListUsers")]
    Param (
        [Parameter(ParameterSetName="AddUsers",Mandatory=$False)] [Switch]$Add,                     ## Add Notifier level user
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] [Switch]$List,                   ## List Notifier level users
        [Parameter(ParameterSetName="RemoveUsers",Mandatory=$False)] [switch]$Remove,               ## Remove Notifier level user
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] 
            [Parameter(ParameterSetName="RemoveUsers",Mandatory=$False)][Switch]$AllPartners,       ## Skip partner selection (Not Implemented)
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] [switch]$Export,                 ## Generate CSV / XLS Output Files
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] [switch]$Launch,                 ## Launch XLS or CSV file 
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] [string]$Delimiter = ',',        ## specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(ParameterSetName="ListUsers",Mandatory=$False)] $ExportPath = "$PSScriptRoot",   ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                     ## Remove Stored API Credentials
        
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Add Notifier User`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
    Write-output "  Current Parameters:"
    Write-output "  -Mode          = $($PsCmdlet.ParameterSetName)"
    Write-output "  -Add User      = $Add"
    Write-output "  -Remove User   = $Remove"
    Write-output "  -List Users    = $List"
    Write-output "  -All Partners  = $AllPartners"    
    Write-output "  -Export        = $Export"
    Write-output "  -Launch        = $Launch"
    Write-output "  -ExportPath    = $ExportPath"
    Write-output "  -Delimiter     = $Delimiter"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $urljson = "https://api.backup.management/jsonapi"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
 
            Write-Output "  Enter Exact, Case Sensitive Partner Name for SolarWinds Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($PartnerName.length -eq 0)
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

    Function GenerateStrongPassword ([Parameter(Mandatory=$false)][int]$PasswordLength = 22) {
        Add-Type -AssemblyName System.Web
        $IsComplex = $False
        do {
            $NewPWD=[System.Web.Security.Membership]::GeneratePassword($PasswordLength,1)
            If (($NewPWD -cmatch "[A-Z\p{Lu}\s]") -and ($NewPWD -cmatch "[a-z\p{Ll}\s]") -and ($NewPWD -match "[\d]") -and ($NewPWD -match "[^\w]")) {$IsComplex=$True}
        } While ($IsComplex -eq $False)
        return $NewPWD
    } ## Generate Strong Password  

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

        $RestrictedPartnerLevel = @("Root","Sub-root")

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
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername
            }

        if ($partner.error) {
            write-output "  $($partner.error.message)"
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
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
                }
                    Else {
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
                }
        }
        
    }  ## EnumeratePartners API Call

    Function Send-EnumerateUsers { 

        Write-Output "  Listing Notifier level users"    
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateUsers'
        $data.params = @{}
        $data.params.partnerIds = ([array]$script:PartnerId)

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:EnumerateUsers = $webrequest | convertfrom-json

            $Script:NotifierUsers = $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber | Where-object {$_.RoleId -eq 7} | Format-Table
    }  ## List all "Notifier level users"

    Function Send-AddUser { 

        $EmailRegex = '^[_a-z0-9-.+]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$'

        Write-Output "  Enter details to create a new 'Notifier level user' for Scheduled Email delivery, without Backup Console access.`n"

        Do { $Script:EmailAddress = Read-Host "  Enter User Email Address     " }until($Script:EmailAddress -match $EmailRegex) 
        Do { $Script:UserName     = Read-Host "  Enter UserName               " }until($Script:UserName.length -ge 6)
        Do { $Script:FirstName    = Read-Host "  Enter User First\ Given Name " }until($Script:FirstName.length -ge 2)
        Do { $Script:LastName     = Read-Host "  Enter User Last\ Sir Name    " }until($Script:LastName.length -ge 4)
        Do { $Script:PhoneNumber  = Read-Host "  Enter User Phone Number      " }until($Script:PhoneNumber.length -ge 8)
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'AddUser'
        $data.params = @{}
        $data.params.userInfo = @{}
        $data.params.userInfo.PartnerId    = [int]$script:PartnerId
        $data.params.userInfo.PhoneNumber  = $Script:PhoneNumber
        $data.params.userInfo.FirstName    = $Script:FirstName
        $data.params.userInfo.FullName     = $Script:LastName
        $data.params.userInfo.Name         = $Script:UserName
        $data.params.userInfo.Password     = GenerateStrongPassword
        $data.params.userInfo.Title        = "Created by $($Script:Authenticate.result.result.EmailAddress)"
        $data.params.userInfo.RoleId       = 7

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:NewUser = $webrequest | convertfrom-json

        Write-Output "  New User Id $($Script:NewUser.result.result) Added"    
    }  ## Add a "Notifier level user"

    Function Send-ModifyUser { 
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'ModifyUser'
        $data.params = @{}
        $data.params.userInfo = @{}
        $data.params.userInfo.Id = [int]$Script:NewUser.result.result
        $data.params.userInfo.EmailAddress = $Script:EmailAddress

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:ModifyUser = $webrequest | convertfrom-json

    }  ## Apply email address to "Notifier level user"

    Function Send-RemoveUser { 
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'RemoveUser'
        $data.params = @{}
        $data.params.userId = [int]$Script:SelectedUser.id

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:RemoveUser = $webrequest | convertfrom-json
            if ($Script:RemoveUser.error) {Write-Output "  $($Script:RemoveUser.error.message)"}
    }  ## Remove a "Notifier level user"

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $switch = $PSCmdlet.ParameterSetName
    Send-APICredentialsCookie
    #Write-Output $Script:strLineSeparator
    #Write-Output "" 
    Send-GetPartnerInfo $Script:cred0

    switch ($PsCmdlet.ParameterSetName)
    {
        "AddUsers" {
            Send-EnumeratePartners 
            Send-EnumerateUsers
            If ($Script:NotifierUsers -eq $null) {Write-Output "  Notifier level users not found" } else { $Script:NotifierUsers } 
            Send-AddUser
            Send-ModifyUser
            Send-EnumerateUsers
            If ($Script:NotifierUsers -eq $null) {Write-Output "  Notifier level users not found" } else { $Script:NotifierUsers } 
        }
        "ListUsers" {
            Send-EnumeratePartners
            Send-EnumerateUsers
            If ($Script:NotifierUsers -eq $null) {Write-Output "  Notifier level users not found" } else { $Script:NotifierUsers } 
        }
        "RemoveUsers" {
            Send-EnumeratePartners 
            Send-EnumerateUsers
            $Script:SelectedUser = $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber | Where-object {$_.RoleId -eq 7} | Out-GridView -Title "Current Partner | $Script:partnername | Select Notifier level user to delete" -OutputMode Single
            If (($Script:SelectedUser -eq $null) -or ($Script:SelectedUser -eq 0)) {Write-Output "  Notifier level user not selected or not found" } else {
                Send-RemoveUser
                Write-Output "  Removed Notifier level user"   
                $Script:SelectedUser | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber | Where-object {$_.RoleId -eq 7} | Format-Table
                Send-EnumerateUsers
                If ($Script:NotifierUsers -eq $null) {Write-Output "  Notifier level users not found" } else { $Script:NotifierUsers }
            }
        }
    }
    
    If ($Script:Export) {
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_Notifier_Users_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber | Where-object {$_.RoleId -eq 7} | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8
         
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
        }
    Start-Sleep -seconds 3
    Read-Host "  Press ENTER to exit..."
    Exit