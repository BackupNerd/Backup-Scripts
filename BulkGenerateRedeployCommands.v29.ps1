<# ----- About: ----
    # Bulk Generate Redeploy Commands
    # Revision v29 - 2020-12-24 
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
    # Requires Superuser access credentials with Security Officer flag
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # !!WARNING!! Output Files may contain device credentials and should be regarded as !!CONFIDENTIAL!!
    #
    # Authenticate to Backup.Management
    # GUI Select Devices
    # Auto Launch XLS or CSV output
    # 
    # Script Bulk generates Backup Manager reinstallation credentials and commands
    # for your backup devices that use Passphase based encryption key management
    # Generates CSV and XLSX files with individual device credentials, download and installer scripts
    # Passphrases generated are valid for 24 hours or a Single Use
    #
    # Use the -DeviceCount ## (Default=1000) Parameter to define how many devices to process
    # Use the -Silent Switch Parameter to Skip GUI Device Selection and AutoLaunch of XLS or CSV file
    # Use the -ExportPath (?:\Folder) Parameter to specify alternate XLS and CSV file path
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/get-passphrase.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/re-install.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 1000,                     ## Change default Number of devices to lookup
        [Parameter(Mandatory=$False)] [switch]$Silent,                              ## Skip GUI Device Selection and Autolaunch of XLS or CSV file
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot"                 ## Export Path
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Bulk Generate Redeploy Commands`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $Script:True_path = "C:\ProgramData\MXB\"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $csvcommands = "$ExportPath\$($CurrentDate)_CONFIDENTIAL_BackupManagerRedeployCommands.csv"
    $WinDownload = '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")'
    $MacDownload = 'curl -o mxb-macosx-x86_64.pkg http://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg'
    $Macinstaller = 'osascript -e "do shell script \"sudo installer -pkg mxb-macosx-x86_64.pkg -target /\"with administrator privileges"'
   
#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function Authenticate {

    Write-Output $Script:strLineSeparator
    Write-Output "  Enter Your SolarWinds Backup https:\\backup.management Login Credentials"
    Write-Output $Script:strLineSeparator

    $AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for SolarWinds Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    $AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for SolarWinds Backup.Management API"
    $AdminPassword = Read-Host -AsSecureString "  Enter Password for SolarWinds Backup.Management API"

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
#endregion ----- Authentication ----
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
    
            function Get-Release-Ref ($ref) {
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
        
            $null = $ws, $wb, $xl | ForEach-Object {Get-Release-Ref $_}
    
            # del $CSVFile
        }
    } ## Save as output XLS Routine

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
    
        $BackupFilter = "at == 1"                         ## Filter to exclude never deployed devices and M365  
        
    # (Create the JSON object to call the EnumerateAccountStatistics function)
        $objEnumerateAccountStatistics = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $Script:visa |
        Add-Member -PassThru NoteProperty method ‘EnumerateAccountStatistics’ |
        Add-Member -PassThru NoteProperty params @{
                                                    query = @{
                                                                      PartnerId=[int]$partnerId
                                                                      Filter="$BackupFilter"
                                                                      Columns=("AU","AR","AN","MN","OP","OT","OS","PF","PN","QW","CD","TS","TL")
                                                                      StartRecordNumber = 0
                                                                      RecordsCount=$DeviceCount
                                                                   }
                                                  }|
        
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 4
    
    # (Call the JSON Web Request Function to get the EnumerateAccountStatistics Object)
        $EnumerateAccountStatisticsSession = CallJSON $urlJSON $objEnumerateAccountStatistics
    
        #Write-Output $Script:strLineSeparator
        #Write-Output "  Using Visa: $($Script:visa)"
        #Write-Output $Script:strLineSeparator
    
    # (Added Delay in case command takes a bit to respond)
        Start-Sleep -Milliseconds 200
    
    # (Get Result Status of GetRecycleBin)
        $EnumerateAccountStatisticsErrorCode = $EnumerateAccountStatisticsSession.error.code
        $EnumerateAccountStatisticsErrorMsg = $EnumerateAccountStatisticsSession.error.message
    
    # (Check for Errors with EnumerateAccountStatistics - Check if ErrorCode has a value)
        if ($EnumerateAccountStatisticsErrorCode) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Get EnumerateAccountStatistics Info Error Code:  $EnumerateAccountStatisticsErrorCode"
            Write-Output "  Get EnumerateAccountStatistics Info Error Message:  $EnumerateAccountStatisticsErrorMsg"
            Write-Output $Script:strLineSeparator
            Write-Output "  Exiting Script"
    # (Exit Script if there is a problem)
    
            #Break Script
        }
        Else {
    # (No error)
        
        }
        
        $Script:DeviceDetail = @()
        ForEach ( $Result in $EnumerateAccountStatisticsSession.result.result ) {
            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID    = [String]$Result.AccountId;
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
    
        if ($silent) {
            $Script:Detail = $DeviceDetail | sort-object PartnerName
        }else{
            $Script:Detail = $DeviceDetail | sort-object PartnerName | out-gridview -OutputMode Multiple -Title "Select all Required Devices and Click OK to Request PassPhrases"
        }
        Write-output "  $($detail.accountid.count) Devices to process"
}
    

#endregion ----- Functions ----

    Authenticate
    
    Get-DeviceList

# ---- Get Passphrase Section ----

    Write-Output $Script:strLineSeparator
    Write-Output "  Requesting Passphrase"
    Write-Output $Script:strLineSeparator

# ----- GenerateReinstallationPassphrase -----

# (Create the JSON object to call the GenerateReinstallationPassphrase function)

Write-output "partnername`tostype`tosversion`tcreation`tlastsuccess`ttimestamp`tid`tcomputername`tdevicename`tpassword`tpassphrase`tdownload`tcommand" | Out-file $csvcommands

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
        
        $objGetPassphrase = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $Script:visa |
        Add-Member -PassThru NoteProperty method ‘GenerateReinstallationPassphrase’ |
        Add-Member -PassThru NoteProperty params @{accountId=$id } |
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the GetPassphrase Object)
       
        $Script:GetPassphraseSession = CallJSON $urlJSON $objGetPassphrase

        $Script:Passphrase = [string]$GetPassphraseSession.result.result

        If (($d.OSVersion -match "macos") -or ($d.OSVersion -match "OS X"))  {
            Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`t$passphrase`t$MACdownload`t$macinstaller") | Out-file $csvcommands -append
        }elseif (($d.OSVersion -notmatch "macos") -and ($d.OSVersion -notmatch "OS X") -and ($d.OSVersion -notmatch "Windows")) {
            Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`t$passphrase`tManual`tManual") | Out-file $csvcommands -append
        }elseif (($d.OSVersion -match "Windows") -and ($GetPassphraseSession.error.message)) {
            Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`tUsing Private Encryption Key`t$Windownload`t" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -encryption-key `"INSERT_PRIVATE_KEY`"") | Out-file $csvcommands -append
        }elseif (($d.OSVersion -match "Windows") -and ($passphrase -ne $null)) {
            #Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`t$passphrase`t$Windownload`t" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
            Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`t$passphrase`t$Windownload;" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
            #Write-output ("$partnername`t$ostype`t$osversion`t$creation`t$lastsuccess`t$timestamp`t$id`t$computername`t$devicename`t$password`t$passphrase`t$tmpDownload;" + '& $env:tmp\mxb-windows-x86_x64.exe' + " -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
        }
    }
    
    $xlscommands = $csvcommands.Replace("csv","xlsx")
    Save-CSVasExcel $csvcommands
    
    If ($silent) { Write-Output "  Silent:  $Silent"
        }else{
            If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
                Start-Process "$xlscommands"
                }else{
                Start-Process "$csvcommands"
                }
        }
    Write-output $Script:strLineSeparator
    Write-output "  NOTE: Files contain credentials and should be regarded as CONFIDENTIAL"
    Write-output "  Redeployment Command Files found here:"
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $csvcommands"
    Write-Output "  XLS Path = $xlscommands"
    Write-Output ""

    start-sleep -seconds 10