<# ----- About: ----
    # N-able Cove Data Protection | Add NonInteractive API User
    # Revision v08 - 2025-02-20
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
    # Add a NonInteractive user that can be used with the API for N-able Cove Data Protection
    # NonInteractive users do not have SSO login access to https://backup.management
    # 
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate Users
    # List/ Add/ Remove Users
    # Optionally import from CSV
    # Use the -Add switch parameter to Add a New  user
    # Use the -AddCLI switch parameter to Add a New user via CLI
    # Use the -ImportUsers switch parameter to import users from CSV
    # Use the -ImportPath (?:\Folder) parameter to specify CSV file paths
    # Use the -List switch parameter to List current users
    # Use the -Remove switch parameter to Remove a current user
    # Use the -SubPartners switch parameter to allow GUI Subpartner selection
    # Use the -Force switch parameter to remove users without prompts
    # Use the -Userrole (default=6) parameter to set the user role ID (1-7)
    # Use the -Type (default='Reporting') parameter to set the user type (Reporting/API) in the Title field

    # Use the -Delimiter (default=',') parameter to set the delimiter for log output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/users.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm
    #
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="AddUsers")]
Param (
    [Parameter(ParameterSetName="AddUsers", Mandatory=$False)] [Switch]$Add,                     ## Add user via prompt
    [Parameter(ParameterSetName="AddCLI", Mandatory=$False)] [Switch]$AddCLI,                    ## Add user via CLI
    [Parameter(ParameterSetName="ListUsers", Mandatory=$False)] [Switch]$List,                   ## List users
    [Parameter(ParameterSetName="RemoveUsers", Mandatory=$False)] [Switch]$Remove,               ## Remove user
    [Parameter(ParameterSetName="ListUsers", Mandatory=$False)]
    [Parameter(ParameterSetName="AddUsers", Mandatory=$False)] 
    [Parameter(ParameterSetName="RemoveUsers", Mandatory=$False)] [Switch]$SubPartners,          ## GUI prompt for subpartner selection
    [Parameter(ParameterSetName="RemoveUsers", Mandatory=$False)] [Switch]$Force,                ## Remove interactive users without prompts
    [Parameter(ParameterSetName="ImportUsers", Mandatory=$False)] [Switch]$ImportUsers,          ## import users from CSV
    [Parameter(ParameterSetName="ImportUsers", Mandatory=$False)] [string]$ImportPath,           ## import users from CSV
    
    [Parameter(ParameterSetName="AddCLI", Mandatory=$False)] 
    [Parameter(ParameterSetName="AddUsers", Mandatory=$False, 
    HelpMessage="Select UserRole(1-7): 1:Superuser, 2:Administrator, 3:Manager, 4:Operator, 5:Supporter, 6:Reporter, 7:Notifier")]
    [ValidateRange(1, 7)] [Int]$Userrole = 1,                                                    ## User Role ID 1-7

    [Parameter(ParameterSetName="AddUsers", Mandatory=$False)]
    [ValidateSet('Reporting', 'API')] [String]$Type = "Reporting",                               ## Usage Type
    [Parameter(ParameterSetName="AddCLI", Mandatory=$True)] [Int]$UserpartnerID,                 ## Partner Name
    [Parameter(ParameterSetName="AddCLI", Mandatory=$True)] [String]$Username,                   ## Partner Name
    [Parameter(ParameterSetName="AddCLI", Mandatory=$False)]
    [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [String]$EmailAddress,                                                                       ## Email Address
    [Parameter(ParameterSetName="AddCLI", Mandatory=$False)] [String]$PhoneNumber,               ## Phone Number
    [Parameter(Mandatory=$False)] [String]$Delimiter = ',',                                      ## specify ',' or ';' Delimiter for XLS & CSV file   
    [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",                                 ## Export Path
    [Parameter(Mandatory=$False)] [Switch]$ClearCredentials                                      ## Remove Stored API Credentials
)

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    
    # Define a lookup table for role IDs and their corresponding roles
    $RoleLookup = @{
        1 = "Superuser"     ## Full access
        2 = "Administrator" ## 
        3 = "Manager"       ## 
        4 = "Operator"      ##
        5 = "Supporter"     ##
        6 = "Reporter"      ## effectively read-only
        7 = "Notifier"      ## Not visible in the Backup Management Console
        <# 
        Example usage: Get role name by role ID
        $RoleId = 3
        $RoleName = $RoleLookup[$RoleId]
        Write-Output "Role ID $RoleId corresponds to role: $RoleName"
        #>
    }


    Write-output "  Add User`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n$Syntax"
    Write-output "  Current Parameters:"
    Write-output "  -Mode          = $($PsCmdlet.ParameterSetName)"
    Write-output "  -Add           = $Add"
    Write-output "  -Remove        = $Remove"
    Write-output "  -List          = $List"
    Write-output "  -Role          = $($RoleLookup[$Userrole])"
    Write-output "  -SubPartners   = $SubPartners"    
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

    Function GenerateStrongPassword {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $false)]
            [int]$PasswordLength = $(Get-Random -Minimum 20 -Maximum 26)  # Set default password length between 18 and 22 characters
        )

        # Add the System.Web assembly to use the Membership class for password generation
        Add-Type -AssemblyName System.Web

        # Initialize a flag to check if the generated password is complex enough
        $IsComplex = $False

        # Loop until a complex password is generated
        do {
            # Generate a password with the specified length and at least one non-alphanumeric character
            $script:NewPWD = [System.Web.Security.Membership]::GeneratePassword($PasswordLength, 1)

            If (
                ($script:NewPWD -cmatch "[A-Z]{3,5}") -and  # Check for 3 to 5 uppercase letters
                ($script:NewPWD -cmatch "[a-z]{3,5}") -and  # Check for 3 to 5 lowercase letters
                ($script:NewPWD -match "\d{2,3}") -and      # Check for 2 to 3 digits
                ($script:NewPWD -match "[^\w]{2,2}") -and    # Check for at least 2 special characters
                ($script:NewPWD -notmatch "[/*=:;.>+`@\{\}\[\]]")  # Ensure no disallowed special characters
            ) {
                $IsComplex = $True  # Set flag to true if password meets all requirements
            }
        } While ($IsComplex -eq $False)  # Repeat until a complex password is generated
        return $script:NewPWD # Return the generated complex password
    } ## Generate Strong Password

    Function Get-RandomAlphanumericString {
	
        [CmdletBinding()]
        Param (
            [int] $length = (Get-Random -Minimum 8 -Maximum 15)
        )
        Begin{
        }
        Process{
            Write-Output ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) | Get-Random -Count $length  | % {[char]$_}) )
        }	
    }

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

Function Send-EnumerateUsers { 

    Write-Output "  Listing users"    
                
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

        $Script:EnumeratedUsers = $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber,@{Name="LastLoginTime";Expression={Convert-UnixTimeToDateTime $_.lastlogintime}},@{Name="FirstLoginTime";Expression={Convert-UnixTimeToDateTime $_.firstlogintime}},* -ErrorAction SilentlyContinue 

}  ## List all users

Function Send-AddUser { 

    $EmailRegex = '^[_a-z0-9-.+]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$'

    if ($PSCmdlet.ParameterSetName -eq "AddUsers") {
        Write-Output "  Enter details to create a new user for API Use without Backup.Management Console access.`n"
        Do { $Script:UserName     = Read-Host "  Enter Username               " }until($Script:UserName.length -ge 8)
        ## Comment out the followng lines if values are not desired
        Do { $Script:EmailAddress = Read-Host "  Enter User Email Address     " }until($Script:EmailAddress -match $EmailRegex) 
        Do { $Script:PhoneNumber  = Read-Host "  Enter User Phone Number      " }until($Script:PhoneNumber.length -ge 8)
    }

    if ($PSCmdlet.ParameterSetName -eq "AddCLI" ) {
        [int]$script:PartnerId = $userpartnerID
    }

    if ($PSCmdlet.ParameterSetName -eq "importusers" ) {
        [int]$script:PartnerId = $newuser.userpartnerID
    }

    $url = "https://api.backup.management/jsonapi"
    $Script:data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'AddUser'
    $data.params = @{}
    $data.params.userInfo = @{}
    $data.params.userInfo.PartnerId    = [int]$script:PartnerId
    $data.params.userInfo.PhoneNumber  = $Script:PhoneNumber
    $data.params.userInfo.Name         = $Script:UserName 
    $data.params.userInfo.Password     = GenerateStrongPassword
    $data.params.userInfo.Title        = "$($Type) - API USE - Added by $($Script:Authenticate.result.result.EmailAddress)"
    $data.params.userInfo.RoleId       = $userrole
    $data.params.userInfo.Flags        = @("AllowApiAuthentication","NonInteractiveUser")

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:NewUser = $webrequest | convertfrom-json

    if ($newuser.error) {
        Write-warning "  $($newuser.error.message -replace '&apos;', '''' -replace '&quot;', '""')"
        Write-warning $script:NewPWD
        # Log the error to a file
        $errorLogData = [PSCustomObject]@{
            timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action    = "AddUser"
            Error     = $newuser.error.message -replace '&apos;', '''' -replace '&quot;', '""'
        }

        $errorLogFilePath = "$ExportPath\ErrorLog.csv"
        $errorLogData | Export-Csv -Path $errorLogFilePath -Append -NoTypeInformation -Encoding UTF8
        return
    }
    else {
        Write-Output $Script:strLineSeparator
        Write-Output "  New User Id         :  $($Script:NewUser.result.result)"
        Write-Output $Script:strLineSeparator
        Write-Output "  -->  PartnerName    :  $($script:PartnerName)"
        Write-Output "  -->  UserName       :  $($data.params.userInfo.Name)"
        Write-Output "  -->  Password       :  $($data.params.userInfo.Password)"
        Write-Output $Script:strLineSeparator
        Write-Output ""
        Write-Output "  partnerId           :  $($script:PartnerId)"
        Write-Output "  UserId              :  $($Script:NewUser.result.result)"
        Write-Output "  Email               :  $($Script:EmailAddress)"
        Write-Output "  Phone               :  $($Script:PhoneNumber)"
        Write-Output "  Note                :  $($data.params.userInfo.Title)" 
        write-output "  Role                :  $($RoleLookup[$Userrole])`n`n`n`n"

        # Write the same data to a log file
        $logData = [PSCustomObject]@{
            timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action       = "AddUser"
            PartnerID    = $script:PartnerID
            UserID       = $Script:NewUser.result.result
            UserName     = $data.params.userInfo.Name
            Password     = $data.params.userInfo.Password
            EmailAddress = $Script:EmailAddress
            PhoneNumber  = $Script:PhoneNumber
            Note         = $data.params.userInfo.Title
            Role         = $RoleLookup[$Userrole]
        }

        $logFilePath = "$ExportPath\NewUserLog.csv"
        $logData | Export-Csv -Path $logFilePath -Append -NoTypeInformation -Encoding UTF8

        }

}  ## Add a user"

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

}  ## Apply email address to user"

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
        if ($Script:RemoveUser.error) {
            Write-Output "  $($Script:RemoveUser.error.message)"
        }
        else{
        # Write the same data to a log file
        $logData = [PSCustomObject]@{
            timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action       = "RemoveUser"
            PartnerID    = $SelectedUser.PartnerId
            UserID       = $SelectedUser.id
            UserName     = $SelectedUser.Name
            Password     = "Not Applicable"
            EmailAddress = $SelectedUser.EmailAddress
            PhoneNumber  = $SelectedUser.PhoneNumber
            Note         = $SelectedUser.Title
            Role         = $SelectedUser.RoleName
        }

        $logFilePath = "$ExportPath\NewUserLog.csv"
        $logData | Export-Csv -Path $logFilePath -Append -NoTypeInformation -Encoding UTF8
        }
}  ## Remove a user

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

$switch = $PSCmdlet.ParameterSetName
Send-APICredentialsCookie
#Write-Output $Script:strLineSeparator
#Write-Output "" 


switch ($PsCmdlet.ParameterSetName)
{
    "AddUsers" {
        Send-GetPartnerInfo $Script:cred0
        Send-EnumeratePartners 
        Send-AddUser
        Send-ModifyUser
    }
    "AddCLI" {
        Send-AddUser
        Send-ModifyUser
    }
    "ImportUsers" {
        $Script:NewUsers = Import-Csv -Path $ImportPath -Delimiter $Delimiter

        # Validate required columns in the CSV
        $requiredColumns = @('UserpartnerID', 'Username', 'EmailAddress', 'PhoneNumber', 'Userrole') 
        $missingColumns = @()

        foreach ($column in $requiredColumns) {
            if (-not ($Script:NewUsers | Get-Member -Name $column -MemberType NoteProperty)) {
                $missingColumns += $column
            }
        }

        if ($missingColumns.Count -gt 0) {
            Write-Warning "The following required columns are missing in the CSV file: $($missingColumns -join ', ')"
            Break
        }

        foreach ($Script:NewUser in $Script:NewUsers) {
            $Script:UserpartnerID = $Script:NewUser.UserpartnerID
            $Script:Username = $Script:NewUser.Username
            $Script:EmailAddress = $Script:NewUser.EmailAddress
            $Script:PhoneNumber = $Script:NewUser.PhoneNumber
            $Script:Userrole = $Script:NewUser.Userrole

            Send-AddUser
            Send-ModifyUser
        }
    }
    "ListUsers" {
        Send-GetPartnerInfo $Script:cred0
        Send-EnumeratePartners
        Send-EnumerateUsers
        If ($Script:enumeratedusers -eq $null) {
            Write-Output "  Users not found" 
        }
        else {
            $Script:enumeratedusers | Where-Object {$_.flags -contains "NonInteractive"} | select-object * -ExcludeProperty flags |  Format-Table
        } 
    }
    "RemoveUsers" {
        Send-GetPartnerInfo $Script:cred0
        Send-EnumeratePartners 
        Send-EnumerateUsers
        if ($Force) {
            $script:SelectedUsers = $Script:enumeratedusers | Where-Object {$_.flags -contains "NonInteractive"} | select-object *
        }
        else{
            $Script:SelectedUsers = $Script:enumeratedusers | Where-Object {$_.flags -contains "NonInteractive"} | select-object * -ExcludeProperty flags | Out-GridView -Title "Current Partner | $Script:partnername | Select Non-Interactive users to delete" -OutputMode Multiple
        }

        If (($Script:SelectedUsers -eq $null) -or ($Script:SelectedUsers -eq 0)) {
            Write-Output "  User not selected or not found" 
        } 
        else {
            foreach ($Script:SelectedUser in $Script:SelectedUsers) {
                Send-RemoveUser
                Write-Output "  Removed user"
            }   
            
            Send-EnumerateUsers
            $NoninteractiveUsers = $Script:enumeratedusers | Where-Object {$_.flags -contains "NonInteractive"}
            If ($NoninteractiveUsers -eq $null) {
                Write-Output "  Users not found" 
            }
            else {
                $NoninteractiveUsers | Format-Table
            }
        }
    }
}

Start-Sleep -seconds 3
Read-Host "  Press ENTER to exit..."
Exit

