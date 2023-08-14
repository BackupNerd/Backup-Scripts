<# ----- About: ----
    # N-able Backup Redeploy Command
    # Revision v02 - 2022-02-21 
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
    # Requires Superuser access credentials with Security Officer and API flag
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # !!WARNING!! Output contains backup credentials and should be regarded as !!CONFIDENTIAL!!
    # Generated passphrases are valid for 24 hours or single use
    #
    # Authenticate to https://backup.management console
    # Searches status.xml for prior installed device name
    # Look up device statistics
    # Request Passphrase
    # Download Backup Manager
    # Redeploy with identified credentials
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/get-passphrase.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/re-install.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [String]$AdminLoginPartnerName,                          ## C
        [Parameter(Mandatory=$False)] [String]$AdminLoginUserName,                            ## S
        [Parameter(Mandatory=$False)] [String]$AdminPassword                                  ## E
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  N-able Backup Redeploy Command`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    $clienttool = "c:\program files\backup manager\clienttool.exe"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $script:True_path = "C:\ProgramData\MXB\"
   
#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function Authenticate {

    Write-Output $script:strLineSeparator
    Write-Output "  Enter Your N-able Backup https:\\backup.management Login Credentials"
    Write-Output $script:strLineSeparator

    if ($AdminLoginPartnerName -eq $null) {$AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"}
    if ($AdminLoginUserName -eq $null) {$AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able Backup.Management API"}
    $AdminPassword = Read-Host -AsSecureString "  Enter Password for N-able Backup.Management API"

# (Convert Password to plain text)
    $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword))

# (Show credentials for Debugging)
    Write-Output "  Logging on with the following Credentials"
    Write-Output ""
    Write-Output "  PartnerName:  $AdminLoginPartnerName"
    Write-Output "  UserName:     $AdminLoginUserName"
    Write-Output "  Password:     It's secure..."
    Write-Output $script:strLineSeparator 

# (Create authentication JSON object using ConvertTo-JSON)
    $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
    Add-Member -PassThru NoteProperty method 'Login' |
    Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$AdminLoginUserName;password=$PlainTextAdminPassword}|
    Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json

# (Call the JSON function with URL and authentication object)
    $script:session = CallJSON $urlJSON $objAuth
    Start-Sleep -Milliseconds 100
    
# (Variable to hold current visa and reused in following routines)
    $script:visa = $session.visa
    $script:PartnerId = [int]$session.result.result.PartnerId
    
    if ($session.result.result.Flags -notcontains "SecurityOfficer" ) {
        Write-Output "  Aborting Script: Invalid Credentials or the Specified User is Not an Authorized Security Officer !!!"
        Write-Output $script:strLineSeparator
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
    #Write-Output $script:strLineSeparator
    #Write-Output "Current Visa is: $script:visa"
    #Write-Output $script:strLineSeparator

## Authenticate Routine
} 
#endregion ----- Authentication ----

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

Function get-devicename {
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
    
        $script:BackupFilter = "an == '$script:accountName'"                         ## Filter to exclude never deployed devices and M365 
        
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
                                                                        PartnerID    = [string]$Result.PartnerId;
                                                                        DeviceName   = $Result.Settings.AN -join '' ;
                                                                        Password     = $Result.Settings.QW -join '' ;
                                                                        ComputerName = $Result.Settings.MN -join '' ;
                                                                        PartnerName  = $Result.Settings.AR -join '' ;
                                                                        Reference    = $Result.Settings.PF -join '' ;
                                                                        Account      = $Result.Settings.AU -join '' ;
                                                                        Creation     = Convert-UnixTimeToDateTime ($Result.Settings.CD -join '') ;
                                                                        TimeStamp    = Convert-UnixTimeToDateTime ($Result.Settings.TS -join '') ;  
                                                                        LastSuccess  = Convert-UnixTimeToDateTime ($Result.Settings.TL -join '') ;   
                                                                        OsType       = $Result.Settings.OT -join '' ;
                                                                        OSVersion    = $Result.Settings.OS -join '' ;
                                                                        Product      = $Result.Settings.PN -join '' ;
                                                                        Profile      = $Result.Settings.OP -join '' }
            }
    
            $script:Detail = $DeviceDetail | sort-object PartnerName
            $script:Detail | format-list
        Write-output "  $($detail.accountid.count) Devices to process"
}

Function Download-BackupManager {
    "  Downloading Backup Manager"
    (New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")
}

Function Stop-BackupProcess {
    stop-process -name "BackupFP" -Force -ErrorAction SilentlyContinue
}

Function Redeploy-BackupManager {

    $BMConfig = "C:\Program Files\Backup Manager\config.ini"

    if ((Test-Path $BMConfig -PathType leaf) -eq $false) {            

    
        Download-BackupManager
        Stop-BackupProcess
        Write-Output "  Redeploying Backup Manager with provided credentials"        
        Write-Output ""

        $process = start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-silent -user $DeviceName -password $Password -passphrase $Passphrase" -PassThru

        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "N-able Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Installer" -Completed
                Break
            }
        }           
       
        Get-BackupService
        Get-InitError

        } else {
        
        Write-Output ""
        Write-Output "  Redeploy aborted, existing Backup Manager CONFIG.INI found"
        Write-Output ""
        
        Break
        }
}

Function Get-BackupService {
    $BackupService = get-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
    
    if ($backupservice.status -eq "Stopped") {
    Write-Output "  Backup Service : $($BackupService.status)"
    }
    elseif ($backupservice.status -eq "Running") {
        Write-Output "  Backup Service : $($BackupService.status)"
        #start-sleep -seconds 10
        #$initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
        #if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
    }
    else{
    Write-Output "  Backup Service : Not Present"
    }
}

Function Get-InitError {
    start-sleep -seconds 10
    $initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
    if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
}

#endregion ----- Functions ----

    Authenticate
    get-devicename
    Get-DeviceList

# ---- Get Passphrase Section ----

    Write-Output $script:strLineSeparator
    Write-Output "  Requesting Passphrase"
    Write-Output $script:strLineSeparator

# ----- GenerateReinstallationPassphrase -----

# (Create the JSON object to call the GenerateReinstallationPassphrase function)

#Write-output "partnername`tostype`tosversion`tcreation`tlastsuccess`ttimestamp`tid`tcomputername`tdevicename`tpassword`tpassphrase`tdownload`tcommand"

    Foreach ($d in $Detail) {
        [int]$id = $d.AccountID
        $ComputerName = $d.ComputerName
        $DeviceName = $d.DeviceName
        $password = $d.password
        $partnername = $d.PartnerName
        $ostype = $d.OsType
        $osversion = $d.OsVersion 
        $creation = $d.creation 
        $LastSuccess = $d.LastSuccess 
        $timestamp = $d.timestamp 
        Write-Output "  Getting Passphrase/Credentials for Device $id"
        start-sleep -Milliseconds 300
        
        $objGetPassphrase = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
        Add-Member -PassThru NoteProperty visa $script:visa |
        Add-Member -PassThru NoteProperty method 'GenerateReinstallationPassphrase' |
        Add-Member -PassThru NoteProperty params @{accountId=$id } |
        Add-Member -PassThru NoteProperty id '1')| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the GetPassphrase Object)
       
        $script:GetPassphraseSession = CallJSON $urlJSON $objGetPassphrase

        $script:Passphrase = [string]$GetPassphraseSession.result.result

        Redeploy-BackupManager

    }
    
