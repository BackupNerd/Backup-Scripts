<# ----- About: ----
    # N-able Backup Get Deleted Devices  
    # Revision v08 - 2021-09-21
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
    # For use with the Standalone edition of N-able Backup
    # Requires a minimum Backup User Role of "Supporter" (Role Id 5)
    # Tested with Powershell 5.1 & 7.0
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure API credentials 
    # Authenticate to https://backup.management console  (Minimum Backup API User Role of "Supporter" requried Role Id 5)
    # Check partner level/ (Optionally) Enumerate partners/ GUI select partner
    # Get Deleted Device list from Maximum Value Report for last 15 days
    # Export to XLS/CSV
    # Optionally launch XLS/CSV
    # Optionally Check/ Get/ Store secure SMP credentials 
    # Optionally Send email via SMTP
    #
    # Use the -Name parameter (Optional) to set exact, case sensitive partner name to report on i.e. 'Acme, Inc (bob@acme.net)'"
    # Use the -AllSubPartners switch parameter to skip GUI partner selection
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    ### Use the -Schedule (inprogress) Schedule Script to run Via Scheduled Task
    ### Use the -Store (inprogress) Copy version of Script to $TaskPath for Scheduled Tasks
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -SendMail switch parameter to send XLS & HTML output via Email
    # Use the -SMTPServer parameter to set SMTP Server Address
    # Use the -SMTPPort parameter to set SMTP Server Port
    # Use the -From parameter to set Email From: Address (normally same as SMTP User)
    # Use the -ReplyTo parameter to set Email ReplyTo: Address (Powershell 7.1 and later)
    # Use the -To parameter to set Email To: Addresses (comma seperated)
    # Use the -Cc parameter to set Email Cc: Addresses (comma seperated)
    # Use the -Bcc parameter to set Email Bcc: Addresses (comma seperated)
    #
    # Use the -ClearSMTPCredentials parameter to remove/reset stored SMTP credentials
    # Use the -ClearAPICredentials parameter to remove/reset stored API credentials
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm 
    # 
    # If you are having gmail "The SMTP server requires a secure connection or the client was not authenticated." issues see this link
    #https://support.google.com/mail/thread/89089797/the-smtp-server-requires-a-secure-connection-or-the-client-was-not-authenticated?hl=en

# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (

        [Parameter(Mandatory=$False)] [string]$Name,                                                        ## Exact, Case Sensitive Partner Name to Report on i.e. 'Acme, Inc (bob@acme.net)'
        [Parameter(Mandatory=$False)] [switch]$AllSubPartners,                                              ## Use to Skip interactive sub-partner selection
        [Parameter(Mandatory=$False)] [string]$Delimiter   = ',',                                           ## Specify ',' or ';' Delimiter for XLS & CSV file   
        [Parameter(Mandatory=$False)] [string]$ExportPath  = "$PSScriptRoot",                               ## Export Path
        [Parameter(Mandatory=$False)] [switch]$Schedule,                                                    ## Schedule Script to run Via Scheduled Task
        [Parameter(Mandatory=$False)] [switch]$Store,                                                       ## Store version of Script in $TaskPath for Scheduled Tasks
        [Parameter(Mandatory=$False)] [switch]$Launch,                                                      ## Launch XLS/CSV file 
        [Parameter(Mandatory=$False)] [switch]$SendMail,                                                    ## Send XLS/HTML output via Email
        [Parameter(Mandatory=$False)] [string]$SMTPServer = "smtp.gmail.com",                               ## SMTP Server Address
        [Parameter(Mandatory=$False)] [string]$SMTPPort   = "587",                                          ## SMTP Server Port
        [Parameter(Mandatory=$False)] [string]$From       = "N-able Backup <nable.backup.nerd@gmail.com>",  ## Email From: Address
        [Parameter(Mandatory=$False)] [string]$ReplyTo    = "Eric Harless <eric.harless@n-able.com>",       ## Email ReplyTo: Address (Powershell 7.1 and later)
        [Parameter(Mandatory=$False)] $To                 = @(                      
                                                            "eric.harless@n-able.com",
                                                            "nable.backup.nerd+TO@gmail.com"
                                                            ),                                              ## Email To: Addresses (comma seperated)
        [Parameter(Mandatory=$False)] $Cc                 = @(
                                                            "eric.harless@n-able.com",
                                                            "nable.backup.nerd+CC@gmail.com"
                                                            ),                                              ## Email Cc: Addresses (comma seperated)
        [Parameter(Mandatory=$False)] $Bcc                = @(
                                                            "eric.harless@n-able.com",
                                                            "nable.backup.nerd+BCC@gmail.com"
                                                            ),                                              ## Email Bcc: Addresses (comma seperated)
        [Parameter(Mandatory=$False)] [string]$TaskPath  = "C:\ProgramData\BackupNerdScripts",              ## Path to Store/Invoke Scheduled Backup Nerd Script Credentials and Tasks
        [Parameter(Mandatory=$False)] [switch]$ClearSMTPCredentials,                                        ## Remove/reset Stored SMTP Credentials
        [Parameter(Mandatory=$False)] [switch]$ClearAPICredentials                                          ## Remove/reset Stored API Credentials

    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Deleted Devices"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $ScriptParent = $PSScriptRoot
    $ScriptPath   = $PSCommandPath
    $ScriptLeaf   = Split-Path $ScriptPath -leaf
    $ScriptName   = $MyInvocation.MyCommand.Name
    $global:FullCommand  = $MyInvocation
    Push-Location $PSScriptRoot
    Write-output   "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $ScriptPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    
    Write-output "  Current Parameters:"
    Write-output "  -AllSubPartners  = $AllSubPartners"
    Write-output "  -Launch          = $Launch"
    Write-output "  -SendMail        = $SendMail"
    Write-output "  -SMTPServer      = $SMTPServer"
    Write-output "  -SMTPPort        = $SMTPPort"
    Write-output "  -Delimiter       = $Delimiter"
    Write-output "  -ExportPath      = $ExportPath"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $PSVersion = $PSVersionTable.PSVersion.Major
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

            Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
        DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
        WHILE ($PartnerName.length -eq 0)

        #$BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.PartnerName | Out-file $APIcredfile
        $BackupCred.UserName    | Out-file -append $APIcredfile
        $BackupCred.Password    | ConvertFrom-SecureString | Out-file -append $APIcredfile
        
        Start-Sleep -milliseconds 300

        Send-APICredentialsCookie  ## Attempt API Authentication

    }  ## Set API credentials if not present

    Function Get-APICredentials {

        $Script:APIcredfile = join-path -Path $TaskPath -ChildPath "$env:computername API_Credentials.Secure.enc"
        $Script:APIcredpath = Split-path -path $APIcredfile

        if (($ClearaPICredentials) -and (Test-Path $APIcredfile)) { 
            Remove-Item -Path $Script:APIcredfile
            $ClearAPICredentials = $Null
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
                    Write-output "  Stored Backup API Partner      = $Script:cred0"
                    Write-output "  Stored Backup API User         = $Script:cred1"
                    Write-output "  Stored Backup API Password     = Encrypted"
                    
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
        $Script:UserId = $authenticate.result.result.id
        $Script:RoleId = $authenticate.result.result.RoleId
        

        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate

    Function Set-SMTPCredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting SMTP Credentials" 
        if (Test-Path $SMTPcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  SMTP Credential Path Present" }else{ New-Item -ItemType Directory -Path $SMTPcredpath} 

        $Script:SMTPCred = Get-Credential -Message 'Enter SMTP User Email and Password to Send Mail'
        
        $SMTPCred.UserName | Out-file $SMTPcredfile
        $SMTPCred.Password | ConvertFrom-SecureString | Out-file -append $SMTPcredfile
        
        Start-Sleep -milliseconds 300

        Get-SMTPCredentials  ## Attempt API Authentication

    }  ## Set API credentials if not present

    Function Get-SMTPCredentials {

        $Script:SMTPcredfile = join-path -Path $TaskPath -ChildPath "$env:computername SMTP_Credentials.Secure.enc"
        $Script:SMTPcredpath = Split-path -path $SMTPcredfile

        if (($ClearSMTPCredentials) -and (Test-Path $SMTPcredfile)) { 
            Remove-Item -Path $Script:SMTPcredfile
            $ClearSMTPCredentials = $Null
            Write-Output $Script:strLineSeparator 
            Write-Output "  SMTP Credential File Cleared"
            Set-SMTPCredentials  ## Retry Authentication
            
            }else{ 
                Write-Output $Script:strLineSeparator 
                Write-Output "  Getting SMTP Credentials" 
            
                if (Test-Path $SMTPcredfile) {
                    Write-Output    $Script:strLineSeparator        
                    "  SMTP Credential File Present"
                    $Script:SMTPcred = get-content $SMTPcredfile
                    
                    $Script:SMTPCredentials=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Script:SMTPcred[0], ($Script:SMTPcred[1] | ConvertTo-SecureString)

                    Write-Output    $Script:strLineSeparator 
                    Write-output "  Stored SMTP User     = $($Script:SMTPcred[0])"
                    Write-output "  Stored SMTP Password = Encrypted"
                    
                }else{
                    Write-Output    $Script:strLineSeparator 
                    Write-Output "  SMTP Credential File Not Present"

                    Set-SMTPCredentials  ## Create SMTP Credential File if Not Found
                    }
                }

    }  ## Get SMTP credentials if present

#endregion ----- Authentication ----

#region ----- Backup.Management JSON Calls ----

    Function Set-ScriptLocation {
        # ----- Self Copy & Logging Logic ----

        $Script:ScriptBase = $Scriptleaf -replace '\..*' 
        $Script:ScriptFinal = Join-Path -Path $TaskPath -ChildPath $ScriptBase | Join-Path -ChildPath $ScriptFile
        $ScriptLog = Join-Path -Path $TaskPath -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.log"
        $ScriptLogParent = Split-path -path $ScriptLog
        mkdir -Force $ScriptLogParent

        Write-Host $Global:strLineSeparator 
        
        if ($PSversion -eq "5") {
            $SetCompressed = Invoke-WmiMethod -Path "Win32_Directory.Name='$ScriptLogParent'" -Name compress
            If (($SetCompressed.returnvalue) -eq 0) { "Items successfully compressed" } else { "Something went wrong!" }
        }

        #Test-Path -Path $ScriptFull,$ScriptFinal,$scriptLog

        If ($ScriptPath -eq $ScriptFinal) {
            Write-Host $Global:strLineSeparator 
            Write-host 'Script Already Running from Target Location'
            Write-Host $Global:strLineSeparator
            } Else {
            Write-Host $Global:strLineSeparator
            Write-host 'Copying Script to Target Location'
            Write-Host $Global:strLineSeparator
            Copy-item -Path $ScriptPath -Destination $ScriptLogParent -Force
            }

        }  ## ----- End Self Copy & Logging Logic ----

        Function Set-ScheduledTask {

        # ----- Windows Task Scheduler Logic ----
            $ScriptDesc = "N-able Backup\$scriptbase"
            $ScriptStart = "21:35"
            $ScriptSched = "MINUTE"
            $ScriptSchedMod = "5"
            
        <# ----- Usage: ----
            # $ScriptSched & $ScriptSchedMod Supported Parameters
            # "MINUTE"  1 - 1439  (Not Recommend with "FilesNotToBackup" Reg Key)
            # "HOURLY"  1 - 23 
            # "DAILY"   1 - 365
            # "WEEKLY"  1 - 52
            # "MONTHLY" 1 - 12
            # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/schtasks#
        # ------------------------------------------------------------#>

            $STAct = New-ScheduledTaskAction -Execute Powershell -WorkingDirectory 'C:\ProgramData\BackupNerdScripts\GetDeletedDevices' -argument "GetDeletedDevices.v07.ps1 -sendmail -allsubpartners -name 'HEH Computing IT'"
            $STPrin = New-ScheduledTaskPrincipal -UserID $env:USERDOMAIN\$env:USERNAME -LogonType S4U -RunLevel Highest
            #$STTrig = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days (365 * 20))
            $STTrig = New-ScheduledTaskTrigger -Daily -At 8am
            Register-ScheduledTask "N-able Backup\GetDeletedDevices4" -Action $STAct -Trigger $STTrig

            Write-Host "Creating Scheduled Task to Run every $ScriptSchedMod $ScriptSched"
            Write-Host $Global:strLineSeparator
            #SCHTASKS.EXE /Create /RU $env:USERDOMAIN\$env:USERNAME /SC $ScriptSched /MO $ScriptSchedMod /TN $ScriptDesc /ST $ScriptStart /RL HIGHEST /F /TR "Powershell $SCRIPT:ScriptFinal"
            Write-Host $Global:strLineSeparator

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

        $RestrictedPartnerLevel = @("Root","SubRoot")

        if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$Script:Level = $Partner.result.result.Level
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
                
                
                if ($AllSubPartners) {
                    $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  All Sub-Partners Selected"
                }else{
                    $Script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername" -OutputMode Single
            
                    if($null -eq $Selection) {
                        # Cancel was pressed
                        # Run cancel script
                        Write-Output    $Script:strLineSeparator
                        Write-Output    "  No Partners Selected"
                        Break
                    }
                    else {
                        # OK was pressed, $Selection contains what was chosen
                        # Run OK script
                        [int]$Script:PartnerId = $Script:Selection.Id
                        [String]$Script:PartnerName = $Script:Selection.Name
                    }
                }

        }
        
    }  ## EnumeratePartners API Call

    Function GetMVReport {
        Param ([Parameter(Mandatory=$False)][Int]$PartnerId) #end param
                
        $Script:end = (get-date).addDays(1).ToString('yyyy-MM-dd')
        $Script:start = (get-date $end).addDays(-15).ToString('yyyy-MM-dd')
        $Script:TempMVReport = "c:\windows\temp\TempMVReport.xlsx"
        
        Write-Output "  Requesting Deleted Devices from the Maximum Value Report for the period between $start and $end"

        $Script:url2 = "https://api.backup.management/statistics-reporting/high-watermark-usage-report?_=6e8d1e0fce68d&dateFrom=$($Start)T00%3A00%3A00Z&dateTo=$($end)T23%3A59%3A59Z&exportOutput=OneXlsxFile&partnerId=$($PartnerId)"
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Cookie", "__cfduid=d7cfe7579ba7716fe73703636ea50b1251593338423; visa=$Script:visa")
        $headers.Add("Authorization","Bearer $Script:visa")



        Write-output  "  $url2"

        Invoke-RestMethod -Uri $url2 `
        -Method GET `
        -Headers $headers `
        -ContentType 'application/json' `
        -WebSession $websession `
        -OutFile $Script:TempMVReport 

        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Host "  Module ImportExcel Already Installed"
        } 
        else {
            try {
                Install-Module -Name ImportExcel -Confirm:$False -Force      ## https://powershell.one/tricks/parsing/excel
            }
            catch [Exception] {
                $_.message 
                exit
            }
        }

        $Script:MVReportxls = Import-Excel -path "$Script:TempMVReport" -asDate "*Date" 
        $Script:MVReport = $Script:MVReportxls | select-object * 
        Remove-Item $Script:TempMVReport
    } 

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    $switch = $PSCmdlet.ParameterSetName

    Send-APICredentialsCookie
    Write-output "  Stored Backup API User Role Id = $RoleId"
    
    $Script:ScriptFinal
    $global:FullCommand

    
    if ($Schedule) {Set-ScheduledTask}
    if ($store) {Set-ScriptLocation}

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    if ($name) {Send-GetPartnerInfo $name}else{Send-GetPartnerInfo $Script:cred0}
    if ($AllSubPartners) {}else{Send-EnumeratePartners}
    
    GetMVReport $partnerId

    $csvoutputfile = "$ExportPath\$($CurrentDate)_RecentlyDeletedDevices_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
    $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
    $MVReportObj = "Parent1Name","Parent1Id","CustomerName","CustomerId","ComputerName","DeviceName","DeviceId","OsType","CustomerState","CreationDate","ProductionDate","DeviceDeletionDate","SelectedSizeGb","UsedStorageGb","O365Users"
    $Deleted = $Script:MVReport | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object $MVReportObj
    $Script:MVReport            | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object $MVReportObj | Export-CSV -path "$csvoutputfile" -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8
    $Script:MVReport            | Where-object {$_.DeviceDeletiondate -ge 1} | Select-object $MVReportObj | Export-Excel -path "$xlsoutputfile" -AutoFilter -AutoSize

    if ($sendmail) {
        Get-SMTPCredentials

        $style      = "<style>BODY{font: Arial 10px;}"
        $style     += "TABLE{border: 1px solid black; border-collapse: collapse;}"
        $style     += "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
        $style     += "TD{border: 1px solid black; padding: 5px;}"
        $style     += "TR:nth-child(even) {background-color: #eee;}"
        $style     += "TR:nth-child(odd) {background-color: #fff;}"
        $style     += "</style>"
        
        $Subject    = "Deleted Backup Device Count | $(@($deleted).count) | $($Partnername) | $($PartnerId) | $($CurrentDate)"
        $Body       = "<b>$(@($deleted).count) </b> recently deleted backup devices found between <b>$($start)</b> and <b>$($end)</b> for partner <b>$PartnerName</b> .<br /><br />" | Out-String
        $Body      += "You should contact backup <a href=https://success.n-able.com/new-case/technical-support>technical support</a> immediately for assisitance if any listed devices are believed to have been accidentially, unintentionally or maliciously deleted.<br /><br />" | Out-String
        $Body      += $Deleted | sort-object DeviceDeletionDate | Select-object @{N='Parent'; E={$_.Parent1Name}},@{N='Customer'; E={$_.CustomerName}},DeviceName,@{N='Created'; E={$_.CreationDate.tostring("yyyy, MMMM dd")}},@{N='Deleted'; E={$_.DeviceDeletiondate.tostring("yyyy, MMMM dd")}} | ConvertTo-Html -Head $style | Out-String
        $Attachment = "$xlsoutputfile"   
        
        Write-output "  Subject              = $Subject"
        Write-output "  From                 = $From"
        Write-output "  To                   = $To"
        Write-output "  Cc                   = $CC"
        Write-output "  Bcc                  = $BCC"

        if ($PSversion -eq "5") { 
        Send-MailMessage -From $From -To $To -Cc $Cc -Subject $Subject `
        -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
        -Credential $SMTPCredentials -Attachments $Attachment -BodyAsHtml -Priority High -Encoding UTF8
        }

        if ($PSversion -eq "7") { 
        Send-MailMessage -From $From -To $To -Cc $Cc -Replyto $Replyto -Subject $Subject `
        -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
        -Credential $SMTPCredentials -Attachments $Attachment -BodyAsHtml -Priority High -Encoding UTF8 -WA 0

        Write-output "  ReplyTo              = $ReplyTo"    
        } ## Powershell 7.1 Compatible -ReplyTo

    } 

    Write-output $Script:strLineSeparator

    ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
        
    if ($Launch) {
        If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
            Start-Process -filepath $xlsoutputfile
            Write-output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
            }else{
                Start-Process -filepath $csvoutputfile
            Write-output $Script:strLineSeparator
            Write-Output "  Opening CSV file"
            Write-output $Script:strLineSeparator            
            }
        }
        
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $csvoutputfile"
    Write-Output "  XLS Path = $xlsoutputfile"
    Write-Output ""

    Start-Sleep -seconds 10