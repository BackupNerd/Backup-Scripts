Clear-Host
<## ----- About: ----
    # Bulk Generate Backup Manager Redeploy Commands
    # Revision v12 - 2020-08-14
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
    # 
    # For user with the Standalone editionof SolarWinds Backup
    # Requires Superuser access credetion with Security Officer flag
    #
    # Bulk generate Backup Manager reinstallation credentials and commands
    # for your backup devices that use Passphase based encryption key management
    # Generates 2 CSV files, one with credentials and one with poweshell downloader and installer commands 
    # Script automatically opens CSV files after generation
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/get-passphrase.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/re-install.htm

# -----------------------------------------------------------#>

# ----- Define Variables ----
    $Global:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
# ----- End Define Variables ----

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

    Function Authenticate {
# ----- Authenticate ----
    #AdminLoginPartnerName = "ExactCompanyName"
    #$AdminLoginUserName = "user@email.address"
    #$AdminPassword = "Ent3rMyP@ss0rd!"

    Clear-Host
    Write-Host $Global:strLineSeparator
    Write-Host "  Enter Your SolarWinds Backup https:\\backup.management Login Credentials"
    Write-Host $Global:strLineSeparator

    $AdminLoginPartnerName = Read-Host -Prompt "  Enter EXACT Login Partner/Customer Name"
    $AdminLoginUserName = Read-Host -Prompt "  Enter Your User Name/Email Address"
    $AdminPassword = Read-Host -AsSecureString "  Enter Your Login Password"

# (Convert Password to plain text)
    $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword))

# (Show credentials for Debugging)
    Write-Host "  Logging on with the following Credentials"
    Write-Host ""
    Write-Host "  PartnerName: " $AdminLoginPartnerName
    Write-Host "  UserName:    " $AdminLoginUserName
    Write-Host "  Password:     It's secure..."
    Write-Host $Global:strLineSeparator 

# (Create authentication JSON object using ConvertTo-JSON)
    $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
    Add-Member -PassThru NoteProperty method ‘Login’ |
    Add-Member -PassThru NoteProperty params @{partner=$AdminLoginPartnerName;username=$AdminLoginUserName;password=$PlainTextAdminPassword}|
    Add-Member -PassThru NoteProperty id ‘1’) | ConvertTo-Json

# (Call the JSON function with URL and authentication object)
    $session = CallJSON $urlJSON $objAuth
    Start-Sleep -Milliseconds 100
    
# (Variable to hold current visa and reused in following routines)
    $Global:visa = $session.visa
    $Global:PartnerId = $session.result.result.PartnerId

# (Get Result Status of Authentication)
    $AuthenticationErrorCode = $Session.error.code
    $AuthenticationErrorMsg = $Session.error.message

# (Check if ErrorCode has a value)
    If ($AuthenticationErrorCode) {
        Write-Host "Authentication Error Code: " $AuthenticationErrorCode
        Write-Host "Authentication Error Message: " $AuthenticationErrorMsg

# (Exit Script if there is a problem)
        Pause
        Break Script
        
    }
    Else {

# (No error)
    
    }

# (Print Visa to screen)
    #Write-Host $Global:strLineSeparator
    #Write-Host "Current Visa is:" $global:visa
    #Write-Host $Global:strLineSeparator

# ----- End Authenticate ----
}

    AUTHENTICATE

    $CurrentDate = Get-Date
    $CurrentDate = $CurrentDate.ToString('yyy-MM-dd_hh-mm-ss')

    $csvfile = "$PSScriptRoot\$($CurrentDate)_BMcreds.csv"
    $csvcommands = "$PSScriptRoot\$($CurrentDate)_BMcommands.csv"
    $url = "http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
    $downloadOld = '(New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")'
    $installer = '& c:\windows\temp\mxb-windows-x86_x64.exe'

    $BackupFilter = "OT >= 1"                             ## Filter to exclude never deployed devices and M365  

    $CountToLookup = 600                                  ## Maximum device count to return

    #$Global:array | select-object "os type","partner name","device id","device name","computer name",password | sort-object "OS type" | format-table
    Write-output "partnername`tostype`tid`tcomputername`tdevicename`tpassword`tpassphrase" | Out-file $csvfile
    Write-output "partnername`tostype`tid`tcomputername`tdevicename`tdownload`tcommand" | Out-file $csvcommands


# ----- Enumerate Account Statistics -----

# (Create the JSON object to call the EnumerateAccountStatistics function)
    $objEnumerateAccountStatistics = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
    Add-Member -PassThru NoteProperty visa $global:visa |
    Add-Member -PassThru NoteProperty method ‘EnumerateAccountStatistics’ |
    Add-Member -PassThru NoteProperty params @{
                                                query = @{
                                                                  PartnerId=$partnerId
        	                                                      Filter="$BackupFilter"
                                                                  Columns=("AU","AR","AN","MN","OP","OT","PF","PN","QW")
        	                                                      StartRecordNumber = 0
        	                                                      RecordsCount=$CountToLookup
                                                               }
                                              }|
    
    Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 4

# (Call the JSON Web Request Function to get the EnumerateAccountStatistics Object)
    $EnumerateAccountStatisticsSession = CallJSON $urlJSON $objEnumerateAccountStatistics

    $EnumerateAccountStatisticsResults = $EnumerateAccountStatisticsSession.result.result | select-object 

    #Write-Host $Global:strLineSeparator
    #Write-Host "  Using Visa:" $global:visa
    #Write-Host $Global:strLineSeparator

# (Added Delay in case command takes a bit to respond)
    Start-Sleep -Milliseconds 200

# (Get Result Status of GetRecycleBin)
    $EnumerateAccountStatisticsErrorCode = $EnumerateAccountStatisticsSession.error.code
    $EnumerateAccountStatisticsErrorMsg = $EnumerateAccountStatisticsSession.error.message

# (Check for Errors with EnumerateAccountStatistics - Check if ErrorCode has a value)
    if ($EnumerateAccountStatisticsErrorCode) {
        Write-Host $Global:strLineSeparator
        Write-Host "  Get EnumerateAccountStatistics Info Error Code: " $EnumerateAccountStatisticsErrorCode
        Write-Host "  Get EnumerateAccountStatistics Info Error Message: " $EnumerateAccountStatisticsErrorMsg
        Write-Host $Global:strLineSeparator
        Write-Host "  Exiting Script"
# (Exit Script if there is a problem)

        #Break Script
    }
    Else {
# (No error)
    
    }
    
    $DeviceDetail = @()
    ForEach ( $Result in $EnumerateAccountStatisticsSession.result.result ) {
        $DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID    = [String]$Result.AccountId;
                                                                    PartnerID    = [string]$Result.PartnerId;
                                                                    DeviceName   = $Result.Settings.AN -join '' ;
                                                                    Password     = $Result.Settings.QW -join '' ;
                                                                    ComputerName = $Result.Settings.MN -join '' ;
                                                                    PartnerName  = $Result.Settings.AR -join '' ;
                                                                    Reference    = $Result.Settings.PF -join '' ;
                                                                    Account      = $Result.Settings.AU -join '' ;
                                                                    OsType       = $Result.Settings.OT -join '' ;
                                                                    Product      = $Result.Settings.PN -join '' ;
                                                                    Profile      = $Result.Settings.OP -join '' }
        }



    #$DeviceDetail | out-gridview 

    #$Detail = $DeviceDetail | where-object{$_.computername -ne $env:computername} 

# ---- Get Passphrase Section ----

    Write-Host $Global:strLineSeparator
    Write-Host "  Requesting Passphrase"
    Write-Host $Global:strLineSeparator

# ----- GenerateReinstallationPassphrase -----

# (Create the JSON object to call the GenerateReinstallationPassphrase function)
    Foreach ($d in $Detail) {
        [int]$id = $d.AccountID
        $ComputerName = $d.ComputerName
        $DeviceName = $d.DeviceName
        $password = $d.password
        $partnername = $d.PartnerName
        $ostype = $d.OsType 
        Write-Host "  Getting Passphrase for $id"
        start-sleep -Milliseconds 500
        
        $objGetPassphrase = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $global:visa |
        Add-Member -PassThru NoteProperty method ‘GenerateReinstallationPassphrase’ |
        Add-Member -PassThru NoteProperty params @{accountId=$id } |
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the GetPassphrase Object)
       
        $GetPassphraseSession = CallJSON $urlJSON $objGetPassphrase
        $Passphrase = $GetPassphraseSession.result.result

        Write-output "$partnername`t$ostype`t$id`t$computername`t$deviceName`t$password`t$passphrase" | Out-file $csvfile -append
        Write-output ("$partnername`t$ostype`t$id`t$computername`t$devicename`t$downloadold`t" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
    }

    Write-output $Global:strLineSeparator
    Write-output "  Redeployment Credentials Files found here:"
    Write-output $Global:strLineSeparator
    Write-output " & `"$csvfile`""

    Write-output ""
    Write-output " & `"$csvcommands`""
    Write-output $Global:strLineSeparator
    & "$csvfile"
    & "$csvcommands"
    start-sleep -seconds 15