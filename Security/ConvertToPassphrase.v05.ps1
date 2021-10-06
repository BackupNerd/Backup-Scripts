<# ----- About: ----
    # N-able Backup Convert to Passphrase Encryption
    # Revision v05 - 2021-10-05
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
    # Requires N-able Backup SuperUser credentials with Security officer role
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Authenticate to https://backup.management console
    # Lookup local N-able Backup device name to get partner CUID
    # Perform Clienttool Takeover to convert Private Key Encryption to Passphrase
    # *** Note: Once executed you can not convert back to private key encryption ***
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/convert-to-passphrase.htm
# -----------------------------------------------------------#>  ## Behavior

# ----- Variables and Paths ----
    Clear-Host
    Write-output "  Convert Private Key to Passphrase Encryption`n`n"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Global:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $Global:True_path = "C:\ProgramData\MXB\"


# ----- End Variables and Paths ----

# ----- Functions ----

    Function Authenticate {

        Write-Output $Script:strLineSeparator
        Write-Output "  Enter Your N-able Backup https:\\backup.management Login Credentials"
        Write-Output $Script:strLineSeparator

        $AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        $AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able Backup.Management API"
        $AdminPassword = Read-Host -AsSecureString "  Enter Password for N-able Backup.Management API"

    # (Convert Password to plain text)
        $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword))

    # (Show credentials for Debugging)
        Write-Output "  Logging on with the following Credentials"
        Write-Output ""
        Write-Output "  PartnerName:  $AdminLoginPartnerName"
        Write-Output "  UserName:     $AdminLoginUserName"
        Write-Output "  Password:     It's secure..."
        Write-Output $Script:strLineSeparator 

    # (Create authentication JSON object using ConvertTo-JSON)
        $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty method ‘Login’ |
        Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$AdminLoginUserName;password=$PlainTextAdminPassword}|
        Add-Member -PassThru NoteProperty id ‘1’) | ConvertTo-Json

    # (Call the JSON function with URL and authentication object)
        $Script:session = CallJSON $urlJSON $objAuth
        Start-Sleep -Milliseconds 100
        
    # (Variable to hold current visa and reused in following routines)
        $Script:visa = $session.visa
        $Script:PartnerId = [int]$session.result.result.PartnerId
        
        if ($session.result.result.Flags -notcontains "SecurityOfficer" ) {
            Write-Output "  Aborting Script: Invalid Credentials or the Specified User is Not an Authorized Security Officer !!!"
            Write-Output $Script:strLineSeparator
            start-sleep -seconds 15
            Break
            }  

    # (Get Result Status of Authentication)
        $AuthenticationErrorCode = $Session.error.code
        $AuthenticationErrorMsg = $Session.error.message

    # (Check if ErrorCode has a value)
        If ($AuthenticationErrorCode) {
            Write-Output "Authentication Error Code:  $AuthenticationErrorCode"
            Write-Output "Authentication Error Message:  $AuthenticationErrorMsg"

    # (Exit Script if there is a problem)
            Pause
            Break Script
            
        }
        Else {

    # (No error)
        
        }

    # (Print Visa to screen)
        #Write-Output $Script:strLineSeparator
        #Write-Output "Current Visa is: $Script:visa"
        #Write-Output $Script:strLineSeparator


    ## Authenticate Routine
    } 

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
        }

    Function SetAPICredentials {

        Write-Output $Global:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $APIcredpath) {
            Write-Output $Global:strLineSeparator 
            "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 

        $PartnerName = Read-Host -assecurestring "Enter EXACT Login Customer Name for N-able Backup.Management API  " | convertfrom-securestring | out-file $APIcredfile
        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
        #$CurrentUserPW = Read-Host -assecurestring "Enter User Password for $env:USERDOMAIN\$env:USERNAME  " | convertfrom-securestring | out-file -append $APIcredfile

        AuthenticateCookie 
    }  ## Set API credentials if not present

    Function GetAPICredentials {

        Write-Output $Global:strLineSeparator 
        Write-Output "  Getting Backup API Credentials" 
        if (Test-Path $APIcredfile) {
            Write-Output    $Global:strLineSeparator        
            "  Backup API Credential File Present"
            $APIcredentials = get-content $APIcredfile
            $global:cred0 = $APIcredentials[0] | Convertto-SecureString
            $global:cred0 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred0))
            $global:cred1 = [string]$APIcredentials[1] 
            $global:cred2 = $APIcredentials[2] | Convertto-SecureString
            $global:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred2))

            #$global:cred3 = $APIcredentials[3] | Convertto-SecureString
            #$global:cred3 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred3))

            Write-Output    $Global:strLineSeparator 
            Write-output "  Stored Backup API Partner  = $cred0"
            Write-output "  Stored Backup API User     = $global:cred1"
            Write-output "  Stored Backup API Password = Encrypted"

            #Write-output "  Stored User Account Password = Encrypted" 
            
            AuthenticateCookie

            }else{
            Write-Output    $Global:strLineSeparator 
            "  Backup API Credential File Not Present"
            SetAPICredentials
            
            } 
    }  ## Get API credentials if present

    Function AuthenticateCookie {

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.partner = $global:cred0
    $data.params.username = $global:cred1
    $data.params.password = $global:cred2
    
    $webrequest = Invoke-WebRequest -Method POST `
    -ContentType 'application/json' `
    -Body (ConvertTo-Json $data) `
    -Uri $url `
    -SessionVariable global:websession `
    -UseBasicParsing
    $global:cookies = $websession.Cookies.GetCookies($url)
    $global:websession = $websession
    
    #Write-output "$($global:cookies[0].name) = $($cookies[0].value)"

    $global:Authenticate = $webrequest | convertfrom-json
    $global:visa = $authenticate.visa
    }

    Function CheckInstallationType {

# ----- Check for RMM & Standalone Backup Installation Type ----
    $MOB_path = "$env:ALLUSERSPROFILE\Managed Online Backup\"
    $MOB_XMLpath = Join-Path -Path $MOB_path -ChildPath "\Backup Manager\StatusReport.xml"
    $MOB_clientpath = "$env:PROGRAMFILES\Managed Online Backup\"
    $SA_path = "$env:ALLUSERSPROFILE\MXB\"
    $SA_XMLpath = Join-Path -Path $SA_path -ChildPath "\Backup Manager\StatusReport.xml"
    $SA_clientpath = "$env:PROGRAMFILES\Backup Manager\"

# ----- Boolean vars to indicate if each exists

    $test_MOB = Test-Path $MOB_XMLpath
    $test_SA = Test-Path $SA_XMLpath

# ----- If both exist, get last modified time and set path of most recent as true_path

    If ($test_MOB -eq $True -And $test_SA -eq $True) {
	    $lm_MOB = [datetime](Get-ItemProperty -Path $MOB_XMLpath -Name LastWriteTime).lastwritetime
	    $lm_SA =  [datetime](Get-ItemProperty -Path $SA_XMLpath -Name LastWriteTime).lastwritetime
	    if ((Get-Date $lm_MOB) -gt (Get-Date $lm_SA)) {
		    $global:true_XMLpath = $MOB_XMLpath
            $global:true_path = $MOB_path
            $global:true_clientpath = $MOB_clientpath
            Write-Output $Global:strLineSeparator
            Write-Output "  Multiple Installations Found - RMM Managed Online Backup is Newest"
	    } else {
		    $global:true_XMLpath = $SA_XMLpath
            $global:true_path = $SA_path
            $global:true_clientpath = $SA_clientpath
            Write-Output $Global:strLineSeparator
            Write-Output "  Multiple Installations Found - Standalone/N-central Backup is Newest"
	    }

# ----- If one exists, set it as true_path

    } elseif ($test_SA -eq $True) {
    	$global:true_XMLpath = $SA_XMLpath
        $global:true_path = $SA_path
        $global:true_clientpath = $SA_clientpath
        Write-Output $Global:strLineSeparator
        Write-Output "  Standalone or N-central Backup Installation Found"
    } elseif ($test_MOB -eq $True) {
    	$global:true_XMLpath = $MOB_XMLpath
        $global:true_path = $MOB_path
        $global:true_clientpath = $MOB_clientpath
        Write-Output $Global:strLineSeparator
        Write-Output "  RMM Managed Online Backup Installation Found"
        Write-Output "  Conversion to Passphrase Not Currently Supported"
        $global:failed = 1

# ----- If none exist, report & fail check

    } else {
        Write-Output $Global:strLineSeparator
    	Write-Output "  Backup Manager Installation Type Not Found"
    	$global:failed = 1
    }
# ----- End Check for RMM & Standalone Backup Installation Type ----

    #If true_path is not null, get XML data
    if ($true_path -eq $SA_path) {
	[xml]$StatusReport = Get-Content $true_XMLpath
	#Get PartnerName
	$global:PartnerName = $StatusReport.Statistics.PartnerName
	$global:InstallationType = $StatusReport.Statistics.InstallationType
    }else{
	Write-Host "  StatusReport.xml Not Found"
	$global:failed = 1
    }
    
}  ## Check for RMM / Ncentral / Standalone Backup Installation Type

    Function GetPartnerInfo ($PartnerName) {      

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
        -SessionVariable global:websession `
        -UseBasicParsing
        $global:cookies = $websession.Cookies.GetCookies($url)
        $global:websession = $websession
    
        #Write-output "$($global:cookies[0].name) = $($cookies[0].value)"

        $Global:Partner = $webrequest | convertfrom-json
        #$global:visa = $authenticate.visa

        if ($Partner.result.result.Uid) {
            [String]$global:Uid = $Partner.result.result.Uid
            [String]$global:PartnerId = $Partner.result.result.Id
            [String]$global:PartnerName = $Partner.result.result.Name

            Write-Output $Global:strLineSeparator
            Write-output "  $PartnerName - $Uid"
            Write-Output $Global:strLineSeparator
            }else{
            Write-Output $Global:strLineSeparator
	        Write-Host "PartnerName or UID Not Found"
            Write-Output $Global:strLineSeparator
	        $global:failed = 1
            }
    }

    Function EnableAutoDeployment ($PartnerId) {      

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $visa
        $data.method = 'ModifyAutoDeploymentPartnerState'
        $data.params = @{}
        $data.params.partnerId = [Int]$PartnerId
        $data.params.autoDeploymentState = 'Enable'
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -depth 5) `
        -Uri $url `
        -SessionVariable global:websession `
        -UseBasicParsing
        $global:cookies = $websession.Cookies.GetCookies($url)
        $global:websession = $websession
    
        #Write-output "$($global:cookies[0].name) = $($cookies[0].value)"

        $Global:AutoDeployment = $webrequest | convertfrom-json
        #$global:visa = $authenticate.visa

    }






          
# ----- End Functions ----

    #GetAPICredentials            
    Authenticate

    CheckInstallationType 
    
    GetPartnerInfo $PartnerName

    EnableAutoDeployment $PartnerID
    

    If ($InstallationType -ne "AutoDeployed") {
    & "C:\Program Files\Backup Manager\ClientTool.exe" takeover -partner-uid $Uid -config-path "c:\Program Files\Backup Manager\config.ini"
    }else{
    Write-Host "  Passphrase Already Exists"
    }

    #If $global:failed is 1, cause scriptcheck to fail in dashboard
if ($global:failed -eq 1) {
	Exit 1001
} else {
	Exit 0
}