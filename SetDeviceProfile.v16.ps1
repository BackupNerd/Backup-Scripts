<# ----- About: ----
    # Set Device Profile 
    # Revision v16 - 2020-10-25
    # Author: Eric Harless, Head Backup Nerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # (Support for Standalone edition of SolarWinds Backup Only)
    #
    # Use with -Clear parameter to Remove Stored API Credentials at start of script
    # Use with -Tail parameter to Display Logs via seperate Powershell Console
    #    
    # Authenticate
    # SelfCopy to ProgramData\MXB folder
    # Enumerate Partners
    # GUI Select Single Partner
    # Enumerate Profiles
    # GUI Select Single Profile
    # Enumerate Account Statistics
    # GUI Select Multiple Devices
    # Set Selected Devices to Selected Profile
    # Write to Log File
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/home.htm 
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/service-management/json-api/API-column-codes.htm
# -----------------------------------------------------------#>

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [Alias("Clear")] [switch]$ClearCredentials,  ## Remove Stored API Credentials at start of script
        [Parameter(Mandatory=$False)] [Alias("Tail")] [switch]$Monitorlog  ## Display Logs via seperate Powershell Console
    )   

    Clear-Host

# ----- Variables and Paths ----
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    Write-Output "  Set Device Profiles"
    Write-Output ""
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n" $Syntax
    $urlJSON = 'https://api.backup.management/jsonapi'


# ----- End Variables and Paths ---

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
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }

    }  ## Use Backup.Management credentials to Authenticate

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Format-FileSize() {
        Param ([int]$size)
        If ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
        ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
        ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
        ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} kB", $size / 1KB)}
        ElseIf ($size -gt 0) {[string]::Format("{0:0.00} B", $size)}
        Else {""}
}

    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $epoch = $epoch.AddSeconds($inputUnixTime)
        return $epoch
        }else{ return ""}
    }  ## Convert epoch time to date time

    Function Get-LogTimeStamp {
        #Get-Date -format s
        return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    }  ## Output proper timeStamp for log file 

    Function Set-Persistance {
        # ----- Self Copy & Logging Logic ----
        
            $Script:True_path = "C:\ProgramData\MXB\"
            $Script:ScriptFull = $Script:myInvocation.MyCommand.path
            $Script:ScriptPath = Split-Path -Parent $Script:MyInvocation.MyCommand.Path
            $Script:ScriptFile = Split-Path -Leaf $Script:MyInvocation.MyCommand.Path
            $Script:ScriptVer = [io.path]::GetFileNameWithoutExtension($Script:MyInvocation.MyCommand.Name)
            $Script:ScriptBase = $ScriptVer -replace '\..*'
            $Script:ScriptFinal = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath $ScriptFile
            $Script:ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.log"
            
            $Script:ScriptLogParent = Split-path -path $ScriptLog
            if (Test-Path $Script:ScriptLogParent) {
                Write-Output $Script:strLineSeparator 
                Write-Output "  ScriptLog Path Present" }else{ New-Item -ItemType Directory -Path $Script:ScriptLogParent}         
            #mkdir -Force $Script:ScriptLogParent
        
            Write-Output    $Script:strLineSeparator 
            $SetCompressed = Invoke-WmiMethod -Path "Win32_Directory.Name='$ScriptLogParent'" -Name compress
            If (($SetCompressed.returnvalue) -eq 0) { "  Log items successfully compressed" } else { "  Something went wrong!" }
        
            if (Test-Path $Script:ScriptLog) {
                Write-Output $Script:strLineSeparator 
                Write-Output "  Log File Present"
                $Script:Scriptlogsize = Format-FileSize((Get-Item $Scriptlog).length)
                }else{ New-Item -ItemType File -Path $Script:ScriptLog} 
    
            #Test-Path -Path $ScriptFull,$ScriptFinal,$scriptLog
            
            If ($ScriptFull -eq $ScriptFinal) {
                Write-Output    $Script:strLineSeparator 
                Write-Output    '  Script Already Running from Target Location'
                Write-Output    $Script:strLineSeparator
                } Else {
                Write-Output    $Script:strLineSeparator
                Write-Output    '  Copying Script to Target Location'
                Write-Output    $Script:strLineSeparator
                Copy-item -Path $ScriptFull -Destination $ScriptLogParent -Force
                }
                
        # ----- End Self Copy & Logging Logic ----
                }

#endregion ----- Data Conversion ----      

#region ----- Backup.Management JSON Calls ----
    Function CallJSON($url,$object) {

        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
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
            Write-output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
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

    Function Send-EnumeratePartners {
# ----- Get Partners via EnumeratePartners -----

# (Create the JSON object to call the EnumeratePartners function)
    $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $Script:visa |
        Add-Member -PassThru NoteProperty method ‘EnumeratePartners’ |
        Add-Member -PassThru NoteProperty params @{
                                                    parentPartnerId = $PartnerId 
                                                    fetchRecursively = "false"
                                                    fields = (0,10,1,5,17,18) 
                                                    } |
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the EnumeratePartners Object)
        [array]$EnumeratePartnersSession = CallJSON $urlJSON $objEnumeratePartners

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
            Write-Output    "  EnumerateAccountPartnersSession Error Code:  $EnumeratePartnersSessionErrorCode"
            Write-Output    "  EnumerateAccountPartnersSession Message:  $EnumeratePartnersSessionErrorMsg"
            Write-Output    $Script:strLineSeparator
            Write-Output    "  Exiting Script"
# (Exit Script if there is a problem)

            #Break Script
        }
            Else {
# (No error)

        $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,Externalcode
        $Script:EnumeratePartnersSessionResults += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} )

        $Script:SelectedPartner = $EnumeratePartnersSessionResults | Sort-object Level,Name | Select-Object Name,Id,Level,Externalcode |
            Where-object {$_.name -notlike "001???????????????- Recycle Bin"} |
            Out-Gridview -Title "Select a PARTNER from available PARTNERS to list available PROFILES" -OutputMode Single 

        if ($SelectedPartner.id -eq $null) {Write-output "  Script Exited, No Changes Made`n"; exit}

        Write-Output "$(Get-LogTimeStamp) Selected Partner = $($SelectedPartner.name) | $($SelectedPartner.id)" | Out-file $ScriptLog -append

        Write-Output    $Script:strLineSeparator
        Write-Output    "  Selected Partner = $($SelectedPartner.name) | $($SelectedPartner.id)"
        Write-Output    $Script:strLineSeparator
                
        }

    # ----- End Get Partners via EnumeratePartners -----
    }

    Function Send-EnumerateAccountProfiles {
# ----- Get Profiles via EnumerateAccountProfiles -----

# (Create the JSON object to call the EnumerateAccountProfiles function)
    $objEnumerateAccountProfiles = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $Script:visa |
        Add-Member -PassThru NoteProperty method ‘EnumerateAccountProfiles’ |
        Add-Member -PassThru NoteProperty params @{partnerId=[int]$SelectedPartner.id } |
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the EnumerateAccountProfiles Object)
    [array]$EnumerateAccountProfilesSession = CallJSON $urlJSON $objEnumerateAccountProfiles

    #$Script:visa = $EnumerateAccountProfilesSession.visa

    #Write-Output    $Script:strLineSeparator
    #Write-Output    "  Using Visa:" $Script:visa
    #Write-Output    $Script:strLineSeparator

# (Added Delay in case command takes a bit to respond)
    Start-Sleep -Milliseconds 100

# (Get Result Status of EnumerateAccountProfiles)
    $EnumerateAccountProfilesSessionErrorCode = $EnumerateAccountProfilesSession.error.code
    $EnumerateAccountProfilesSessionErrorMsg = $EnumerateAccountProfilesSession.error.message

# (Check for Errors with EnumerateAccountProfiles - Check if ErrorCode has a value)
    if ($EnumerateAccountProfilesSessionErrorCode) {
        Write-Output    $Script:strLineSeparator
        Write-Output    "  EnumerateAccountProfilesSession Error Code:  $EnumerateAccountProfilesSessionErrorCode"
        Write-Output    "  EnumerateAccountProfilesSession Message:  $EnumerateAccountProfilesSessionErrorMsg"
        Write-Output    $Script:strLineSeparator
        Write-Output    "  Exiting Script"
# (Exit Script if there is a problem)

        #Break Script
    }
    Else {
# (No error)

        $Script:EnumerateAccountProfilesSessionResults = $EnumerateAccountProfilesSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}}
        $Script:EnumerateAccountProfilesSessionResults += @( [pscustomobject]@{Name='None';Id="0"})
        
        $Script:SelectedProfile = $EnumerateAccountProfilesSessionResults | where-object {$_.id -ne "5038"} | select-object Name,@{l='Id';e={($_.Id).tostring()}} |
            Out-Gridview -Title "Select a PROFILE from available PROFILES to later apply to selected DEVICES" -OutputMode Single | Sort-Object Name,Id

        Write-Output "$(Get-LogTimeStamp) Selected Profile = $($SelectedProfile.name) | $($SelectedProfile.id)" | Out-file $ScriptLog -append

        if ($SelectedProfile.id -eq $null) {Write-output "  Script Exited, No Changes Made`n"; exit}

        Write-Output    $Script:strLineSeparator
        Write-Output    "  Selected Profile = $($SelectedProfile.name)"
        Write-Output    $Script:strLineSeparator
                    
        }

    }
    
    Function Send-EnumerateAccountStatistics {

# ----- Get Devices via EnumerateAccountStatistics -----

$Filter = "((!(PF =~ '[??????????]* - ????????-????-????-????-????????????') AND !(AR =~ '001???????????????- Recycle Bin')) AND (OP != 'Documents') AND (AT != 2))"
    $Columns = ('AU','AR','AN','OI','OP','OS','OT','PD','PF','PN','AT','MN','TS','AA843')
    $MaxDeviceCount=3000

# (Create the JSON object to call the EnumerateAccountStatistics function)
    $objEnumerateAccountStatistics = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
    Add-Member -PassThru NoteProperty visa $Script:visa |
    Add-Member -PassThru NoteProperty method ‘EnumerateAccountStatistics’ |
    Add-Member -PassThru NoteProperty params @{
                                                query = @{
                                                                  PartnerId=[int]$SelectedPartner.id
        	                                                      Filter=$filter
                                                                  Columns=$columns
        	                                                      StartRecordNumber = 0
        	                                                      RecordsCount=$MaxDeviceCount
                                                                  Totals=("COUNT(AT==1)","SUM(T3)","SUM(US)")
                                                               }
                                              }|
    
    Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 4

# (Call the JSON Web Request Function to get the EnumerateAccountStatistics Object)
    [array]$Script:EnumerateAccountStatisticsSession = CallJSON $urlJSON $objEnumerateAccountStatistics

    #$Script:visa = $EnumerateAccountStatisticsSession.visa

    #$SessionResults = $EnumerateAccountStatisticsSession.result.result | select-object 

    #$SessionResults | out-gridview -PassThru 

    #Write-Output    $Script:strLineSeparator
    #Write-Output    "  Using Visa:" $Script:visa
    #Write-Output    $Script:strLineSeparator

# (Added Delay in case command takes a bit to respond)
    Start-Sleep -Milliseconds 100

# (Get Result Status of EnumerateAccountStatistics)
    $SessionErrorCode = $EnumerateAccountStatisticsSession.error.code
    $SessionErrorMsg = $EnumerateAccountStatisticsSession.error.message

# (Check for Errors with EnumerateAccountStatistics - Check if ErrorCode has a value)
    if ($SessionErrorCode) {
        Write-Output    $Script:strLineSeparator
        Write-Output    "  EnumerateAccountStatistics Error Code:  $SessionErrorCode"
        Write-Output    "  EnumerateAccountStatistics Message:  $SessionErrorMsg"
        Write-Output    $Script:strLineSeparator
        Write-Output    "  Exiting Script"
# (Exit Script if there is a problem)

        #Break Script
    }
    Else {
# (No error)
    
    }
     
    $DeviceDetail = @()
    ForEach ( $Result in $EnumerateAccountStatisticsSession.result.result ) {
        $DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID   = [String]$Result.AccountId;
                                                                    PartnerID   = [string]$Result.PartnerId;
                                                                    Account     = $Result.Settings.AU -join '' ;
                                                                    TimeStamp   = Convert-UnixTimeToDateTime ($Result.Settings.TS -join '') ;                                                                  
                                                                    DeviceName  = $Result.Settings.AN -join '' ;
                                                                    MachineName = $Result.Settings.MN -join '' ; 
                                                                    PartnerName = $Result.Settings.AR -join '' ;
                                                                    Reference   = $Result.Settings.PF -join '' ;
                                                                    Product     = $Result.Settings.PN -join '' ;
                                                                    ProductID   = $Result.Settings.PD -join '' ;                                                                     
                                                                    Profile     = $Result.Settings.OP -join '' ;
                                                                    ProfileID   = $Result.Settings.OI -join '' ;                                                                
                                                                    OsType      = $Result.Settings.OT -join '' ;
                                                                    OS          = $Result.Settings.OS -join '' ;                                                                           
                                                                    AccountType = $Result.Settings.AT -join '' ;
                                                                    Notes       = $Result.Settings.AA843 -join '' }
        }
        
# (Summarize DeviceDetail)
    
    Write-Output "$(Get-LogTimeStamp) Total Device Count = $($DeviceDetail.Count)" | Out-file $ScriptLog -append

    Write-Output $Script:strLineSeparator
    Write-Output "  Total Device Count = $($DeviceDetail.Count)"
    Write-Output $Script:strLineSeparator
       
    #$DeviceDetail | out-gridview -PassThru | Export-Csv c:\data\Migration.csv -notype
    #$DeviceDetail | select PartnerId,Reference,PartnerName,DeviceName,MachineName,AccountID,Product,Profile,OsType,Notes | out-gridview -PassThru | Export-Csv -delimiter "`t" c:\data\backupdevices.csv -notype

    $Script:SelectedDevice = $DeviceDetail | Select-Object Account,PartnerId,Reference,PartnerName,DeviceName,MachineName,AccountID,Product,ProductID,AccountType,profileId,Profile,OS,OsType,Timestamp,Notes |
        Out-Gridview -Title "Assign PROFILE | $($selectedProfile.name) | to selected DEVICES | use SHIFT to select multiple DEVICES" -OutputMode Multiple

    $Script:SelectedDevice =     $Script:SelectedDevice | sort-object DeviceName
    Write-Output "$(Get-LogTimeStamp) Selected Device Count = $($Script:SelectedDevice.AccountId.count)" | Out-file $ScriptLog -append

    if ($SelectedDevice.AccountId -eq $null) {Write-output "  Script Exited, No Changes Made`n"; exit}

    Write-Output $Script:strLineSeparator
    Write-Output "  Selected Device Count = $($Script:SelectedDevice.AccountId.Count)"
    Write-Output $Script:strLineSeparator
        
    }   
          
    Function Send-ModifyAccount {
# ----- ModifyAccount -----

# (Create the JSON object to call the ModifyAccount function)
        Foreach ($Selection in $SelectedDevice) {
        
            $objModifyAccount = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
            Add-Member -PassThru NoteProperty visa $Script:visa |
            Add-Member -PassThru NoteProperty method ‘ModifyAccount’ |
            Add-Member -PassThru NoteProperty params @{
                                                        accountInfo = @{
                                                                        Id=[int]$($Selection.AccountID)
                                                                        ProfileId=[int]$($SelectedProfile.id)
                                                                        }
                                                        } |
            Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the ModifyAccount Object)
       
            $ModifyAccountSession = CallJSON $urlJSON $objModifyAccount
            
            #$Script:visa = $ModifyAccountSession.visa
            #Write-Output    $Script:strLineSeparator
            #Write-Output    "  Using Visa:" $Script:visa
            #Write-Output    $Script:strLineSeparator

# (Added Delay in case command takes a bit to respond)
            Start-Sleep -Milliseconds 100

# (Get Result Status of ModifyAccountSession)
            $ModifyAccountSessionErrorCode = $ModifyAccountSession.error.code
            $ModifyAccountSessionErrorMsg = $ModifyAccountSession.error.message

# (Check for Errors with ModifyAccountSession - Check if ErrorCode has a value)
            if ($ModifyAccountSessionErrorCode) {
                Write-Output    $Script:strLineSeparator
                Write-Output    "  ModifyAccountSession Error Code:  $ModifyAccountSessionErrorCode"
                Write-Output    "  ModifyAccountSession Message:  $ModifyAccountSessionErrorMsg"
                Write-Output    $Script:strLineSeparator
                Write-Output "$(Get-LogTimeStamp) | DEVICE | $($Selection.AccountID) | $($Selection.DeviceName) | ASSIGN PROFILE ERROR | $ModifyAccountSessionErrorMsg" | Out-file $ScriptLog -append

# (Exit Script if there is a problem)

                #Break Script
            }
            Else {
# (No error)
                Write-Output    "  Profile | $($SelectedProfile.Id) | $($SelectedProfile.Name) | Assigned to Device | $($Selection.AccountID) | $($Selection.DeviceName) | Old Profile |  $($Selection.profileId) | $($Selection.Profile)"    
                Write-Output "$(Get-LogTimeStamp) Profile | $($SelectedProfile.Id) | $($SelectedProfile.Name) | Assigned to Device | $($Selection.AccountID) | $($Selection.DeviceName) | Old Profile | $($Selection.profileId) | $($Selection.Profile)" | Out-file $ScriptLog -append
  
            }
        }
    }
   
  #endregion ----- Backup.Management JSON Calls ----

    Function Exit-Routine {
        Write-Output $Script:strLineSeparator
        Write-Output "  NOTE: Profile changes may not be updated in the console or via API call until after"
        Write-Output "        offline devices reconnect, running jobs complete, or servers check-in."       
        Write-Output "        Restarting the 'Backup Service Controller' service should force the update."          
        Write-Output $Script:strLineSeparator
        Write-Output "  Log file found here:"
        Write-Output $Script:strLineSeparator
        Write-Output "  & $scriptlog"
        Write-Output ""
        Write-Output "  $scriptlogsize"
        Write-Output $Script:strLineSeparator
        Start-Sleep -seconds 15
        }

#endregion ----- Functions ----

    Set-Persistance    

    Send-APICredentialsCookie

    Send-GetPartnerInfo $Script:cred0

    Send-EnumeratePartners

    If ($Monitorlog) { invoke-expression "cmd /c start powershell -NoExit -Command {(Get-Host).ui.RawUI.WindowTitle=`" SET PROFILE LOG - $scriptlog `"; Get-Content $scriptlog -tail 2 -Wait }" }

    Send-EnumerateAccountProfiles
    
    Send-EnumerateAccountStatistics
    
    Send-ModifyAccount
    
    Exit-Routine





    