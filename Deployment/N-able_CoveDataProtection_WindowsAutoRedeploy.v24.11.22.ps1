<# ----- About: ----
    # N-able Backup Windows Automatic Redeploy
    # Revision 2024-11-2 
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
    # For use with Cove Data Protection from N-able (Formerly the Standalone editon of N-able Backup)
    # Requires Superuser access credentials with Security Officer and API flag
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # !!IMPORTANT!! Script requests backup reinstallation credentials that should be regarded as !!CONFIDENTIAL!!
    # Generated passphrases are valid for 24 hours or single use
    #
    # Authenticate to https://backup.management console
    # Searches statusreport.xml for prior installed backup device name
    # Lookup device statistics
    # Request Passphrase
    # Download Backup Manager
    # Redeploy with identified credentials
    #
    # Pass the -AdminLoginPartnerName to specify the partner\customer name when authenticating to Backup.Management 
    # Pass email address and password via secure credentials
    # Use the -Force parameter to overwrite an existing installation
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/get-passphrase.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/re-install.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)] [string]$AdminLoginPartnerName,                                                ## Backup.Management Customer Name
        [Parameter(Mandatory=$true)] [ValidateNotNull()] [System.Management.Automation.PSCredential]$Credential,    ## Backup.Management User Credentials
        [Parameter(Mandatory=$false)] [switch]$Force                                                                ## Force Overwrite installation 
    )

    $script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $clienttool = "c:\program files\backup manager\clienttool.exe"
    $script:True_path = "C:\ProgramData\MXB\"

Function Authenticate {

    
    <# Show credentials for Debugging
    Write-Output "  Logging on with the following Credentials"
    Write-Output ""
    Write-Output "  PartnerName:  $AdminLoginPartnerName"
    Write-Output "  UserName:     $($Credential.username)"
    Write-Output "  Password:     It's secure..." 
    #>

# (Create authentication JSON object using ConvertTo-JSON)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $AdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
    
    $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
    Add-Member -PassThru NoteProperty method 'Login' |
    Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$Credential.username;password=$AdminPassword}|
    Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json

# (Call the JSON function with URL and authentication object)
    $script:session = CallJSON $urlJSON $objAuth
    Start-Sleep -Milliseconds 100
    
# (Variable to hold current visa and reused in following routines)
    $script:visa = $session.visa
    $script:PartnerId = [int]$session.result.result.PartnerId
    
    if ($session.result.result.Flags -notcontains "SecurityOfficer" ) {
        Write-Warning "  Aborting Script: Invalid Credentials or the Specified User is Not an Authorized Security Officer !!!"
        Write-Output $script:strLineSeparator
        start-sleep -seconds 5
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


    } ## Authenticate Function

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

Function Get-DeviceName {
    [xml]$status = get-content -path "C:\ProgramData\MXB\Backup Manager\StatusReport.xml"

    if ($status) {$script:AccountName = $status.Statistics.account}else{write-warning "  No status.xml file found, auto redeploy not possible ";Break}
    }

Function Convert-UnixTimeToDateTime($inputUnixTime){
    if ($inputUnixTime -gt 0 ) {
    $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
    $epoch = $epoch.ToUniversalTime()
    $epoch = $epoch.AddSeconds($inputUnixTime)
    return $epoch
    }else{ return ""}
}  ## Convert epoch time to date time 

Function Get-DeviceList {

    # ----- Enumerate Account Statistics -----
    
        $script:BackupFilter = "an == '$script:accountName'"        ## Filter to match device name
        
    # (Create the JSON object to call the EnumerateAccountStatistics function)
        $objEnumerateAccountStatistics = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
        Add-Member -PassThru NoteProperty visa $script:visa |
        Add-Member -PassThru NoteProperty method 'EnumerateAccountStatistics' |
        Add-Member -PassThru NoteProperty params @{
                                                    query = @{
                                                                      PartnerId=[int]$partnerId
                                                                      Filter="$BackupFilter"
                                                                      Columns=("AU","AR","AN","MN","OP","OT","OS","PF","PN","QW","CD","TS","TL")
                                                                      StartRecordNumber = 0
                                                                      RecordsCount = 1
                                                                   }
                                                  }|
        
        Add-Member -PassThru NoteProperty id '1')| ConvertTo-Json -Depth 4
    
    # (Call the JSON Web Request Function to get the EnumerateAccountStatistics Object)
        $EnumerateAccountStatisticsSession = CallJSON $urlJSON $objEnumerateAccountStatistics
    
        #Write-Output $script:strLineSeparator
        #Write-Output "  Using Visa: $($script:visa)"
        #Write-Output $script:strLineSeparator
    
    # (Added Delay in case command takes a bit to respond)
        Start-Sleep -Milliseconds 200
    
    # (Get Result Status of GetRecycleBin)
        $EnumerateAccountStatisticsErrorCode = $EnumerateAccountStatisticsSession.error.code
        $EnumerateAccountStatisticsErrorMsg = $EnumerateAccountStatisticsSession.error.message
    
    # (Check for Errors with EnumerateAccountStatistics - Check if ErrorCode has a value)
        if ($EnumerateAccountStatisticsErrorCode) {
            Write-Output $script:strLineSeparator
            Write-Output "  Get EnumerateAccountStatistics Info Error Code:  $EnumerateAccountStatisticsErrorCode"
            Write-Output "  Get EnumerateAccountStatistics Info Error Message:  $EnumerateAccountStatisticsErrorMsg"
            Write-Output $script:strLineSeparator
            Write-Output "  Exiting Script"
    # (Exit Script if there is a problem)
    
            #Break Script
        }
        Else {
    # (No error)
        
        }
        $script:DeviceDetail = @()
        ForEach ( $Result in $EnumerateAccountStatisticsSession.result.result ) {
            $script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID    = [String]$Result.AccountId;
                                                                        PartnerID     = [string]$Result.PartnerId;
                                                                        DeviceName    = $Result.Settings.AN -join '' ;
                                                                        InstallationKey = $Result.Settings.QW -join '' ;
                                                                        ComputerName  = $Result.Settings.MN -join '' ;
                                                                        PartnerName   = $Result.Settings.AR -join '' ;
                                                                        Reference     = $Result.Settings.PF -join '' ;
                                                                        Account       = $Result.Settings.AU -join '' ;
                                                                        Creation      = Convert-UnixTimeToDateTime ($Result.Settings.CD -join '') ;
                                                                        TimeStamp     = Convert-UnixTimeToDateTime ($Result.Settings.TS -join '') ;  
                                                                        LastSuccess   = Convert-UnixTimeToDateTime ($Result.Settings.TL -join '') ;   
                                                                        OsType        = $Result.Settings.OT -join '' ;
                                                                        OSVersion     = $Result.Settings.OS -join '' ;
                                                                        Product       = $Result.Settings.PN -join '' ;
                                                                        Profile       = $Result.Settings.OP -join '' }
            }
    
        $script:Detail = $DeviceDetail | sort-object PartnerName
        $script:Detail | format-list
        Write-output "  $($detail.accountid.count) Devices to process"
}

Function Download-BackupManager {

    $HttpCDN = "http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
    $HttpsCDN = "https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
    $BackupManager = "C:\Users\Public\Downloads\mxb-windows-x86_x64.exe"

    Remove-Item -Path $BackupManager -Force -ErrorAction SilentlyContinue

    write-output "  Downloading Backup Manager"; (New-Object System.Net.WebClient).DownloadFile($HttpCDN,$BackupManager)
    ## Attempt download over HTTP

    if ((Test-Path $BackupManager -PathType leaf) -eq $false) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Ssl3 -bor [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

        write-output "  Attempting Backup Manager Download via SSL3, TLS 1.0 -> 1.3"; (New-Object System.Net.WebClient).DownloadFile($HttpsCDN,$BackupManager)
    }
}

Function Stop-BackupProcess {
    stop-process -name "BackupFP" -Force -ErrorAction SilentlyContinue
}

Function Redeploy-BackupManager {

    $BMConfig = "C:\Program Files\Backup Manager\config.ini"
    $BackupManager = "C:\Users\Public\Downloads\mxb-windows-x86_x64.exe"

    if (((Test-Path $BMConfig -PathType leaf) -eq $false) -or ($Force)) {            

        Download-BackupManager
        Stop-BackupProcess
        Write-Output "  Redeploying Backup Manager with provided credentials"        
        Write-Output ""

        $process = start-process -FilePath $BackupManager -ArgumentList "-silent -user $DeviceName -password $installationkey -passphrase $Passphrase" -PassThru

        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "N-able Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Installer" -Completed
                Break
            }
        }           
        Remove-Item -Path $BackupManager -Force
        Get-BackupService
        Get-InitError


        } else {
        Write-Warning "  Redeploy aborted, existing Backup Manager CONFIG.INI found`n"
        Break
        }
}

Function Get-BackupService {
    $BackupService = get-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
    
    if ($backupservice.status -eq "Stopped") {
        Write-Output "  Backup Service : $($BackupService.status)`n"
        
    }
    elseif ($backupservice.status -eq "Running") {
        Write-Output "  Backup Service : $($BackupService.status)`n"

    }
    else{
        Write-Warning "  Backup Service : Not Present`n"
    }
}

Function Get-RecoveryConsoleService {
    $Script:RecoveryConsoleService = get-service -name "Recovery Console Service" -ErrorAction SilentlyContinue
    
    if ($RecoveryConsoleService.status -eq "Stopped") {
        Write-Output "  Recovery Console : $($RecoveryConsoleService.status)"
        Write-Warning "  Aborting Backup Manager Redeployment"
        Start-Sleep -Seconds 10
        Break
    }
    elseif ($RecoveryConsoleService.status -eq "Running") {
        Write-Output "  Recovery Console : $($RecoveryConsoleService.status)"
        Write-Warning "  Aborting Backup Manager Redeployment"
        Start-Sleep -Seconds 10
        Break
    }
    else{
        Write-Output "  Recovery Console : Not Present"
    }
}

Function Get-RecoveryServiceController {
    $script:RecoveryServiceController = get-service -name "Recovery Service Controller" -ErrorAction SilentlyContinue
    
    if ($RecoveryServiceController.status -eq "Stopped") {
        Write-Output "  Recovery Service : $($RecoveryServiceController.status)"
        Write-Warning "  Aborting Backup Manager Redeployment"
        Start-Sleep -Seconds 10
        Break
    }
    elseif ($RecoveryServiceController.status -eq "Running") {
        Write-Output "  Recovery Service : $($RecoveryServiceController.status)"
        Write-Warning "  Aborting Backup Manager Redeployment"
        Start-Sleep -Seconds 10
        Break
    }
    else{
        Write-Output "  Recovery Service : Not Present"
    }
}

Function Get-InitError {
    start-sleep -seconds 10
    $initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
    if ($($initmsg.message)) { Write-Warning "  InitMsg Error  : $($initmsg.message)" }
}

#endregion ----- Functions ----

    Authenticate
    Get-DeviceName
    Get-DeviceList

# ---- Get Passphrase Section ----

    Write-Output $script:strLineSeparator
    Write-Output "  Requesting Passphrase"
    Write-Output $script:strLineSeparator

# ----- GenerateReinstallationPassphrase -----

# (Create the JSON object to call the GenerateReinstallationPassphrase function)

    Foreach ($d in $Detail) {
        [int]$id = $d.AccountID
        $DeviceName = $d.DeviceName
        $installationkey = $d.installationkey

        Write-Output "  Getting Passphrase/Credentials for Device $id"
        Start-Sleep -Milliseconds 300
        
        $objGetPassphrase = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
        Add-Member -PassThru NoteProperty visa $script:visa |
        Add-Member -PassThru NoteProperty method 'GenerateReinstallationPassphrase' |
        Add-Member -PassThru NoteProperty params @{accountId=$id } |
        Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the GetPassphrase Object)
       
        $script:GetPassphraseSession = CallJSON $urlJSON $objGetPassphrase
        $script:Passphrase = [string]$GetPassphraseSession.result.result

        Get-RecoveryServiceController
        Get-RecoveryConsoleService
        Get-BackupService
        Redeploy-BackupManager

    }
    






