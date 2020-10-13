Clear-Host
<## ----- About: ----
    # Bulk Generate Backup Manager Redeploy Commands
    # Revision v25 - 2020-13-06
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

<# ----- Behavior: ----
    # 
    # For use with the Standalone edition of SolarWinds Backup
    # Requires Superuser access credentials with Security Officer flag
    #
    # Bulk generates Backup Manager reinstallation credentials and commands
    # for your backup devices that use Passphase based encryption key management
    # Generates CSV and XLSX files with individual device credentials, download and installer scripts
    # Script automatically opens CSV or XLSX file after generation
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/get-passphrase.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/re-install.htm
# -----------------------------------------------------------#>  ## Behavior

# ----- Define Variables ----
    $Global:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $csvcommands = "$PSScriptRoot\$($CurrentDate)_BackupManagerRedeployCommandCredentials.csv"
    $url = "http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
    $Download = '(New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")'
    $tmpDownload = '(New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$ENV:tmp\mxb-windows-x86_x64.exe")'
    $macDownload = 'curl -o mxb-macosx-x86_64.pkg https://cdn.cloudbackup.management/maxdownloads/mxb-macosx-x86_64.pkg'
    $installer = '& c:\windows\temp\mxb-windows-x86_x64.exe'
    $macinstaller = 'osascript -e "do shell script \"sudo installer -pkg mxb-macosx-x86_64.pkg -target /\"with administrator privileges"'
    
    #$Global:array | select-object "os type","partner name","device id","device name","computer name",password | sort-object "OS type" | format-table
    


# ----- End Define Variables ----

# ----- Functions ----

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
}

    Function Authenticate {

    Clear-Host
    Write-Host $Global:strLineSeparator
    Write-Host "  Bulk Request Passphrase Reinstallation Commands & Credentials"
    Write-Host $Global:strLineSeparator
    Write-Host "  Enter Your SolarWinds Backup https:\\backup.management Login Credentials"
    Write-Host $Global:strLineSeparator

    $AdminLoginPartnerName = Read-Host -Prompt "  Exact\Case Sensitive (https:\\backup.management) Customer Name i.e.  AcmeTech Inc (tom@atech.net)"
    $AdminLoginUserName = Read-Host -Prompt "  (https:\\backup.management) User Name i.e.  dan@atech.net"
    $AdminPassword = Read-Host -AsSecureString "  (https:\\backup.management) Password"

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
    $global:session = CallJSON $urlJSON $objAuth
    Start-Sleep -Milliseconds 100
    
# (Variable to hold current visa and reused in following routines)
    $Global:visa = $session.visa
    $Global:PartnerId = [int]$session.result.result.PartnerId
    
    if ($session.result.result.Flags -notcontains "SecurityOfficer" ) {
        Write-Host "  Aborting Script: The Specified User is Not an Authorized Security Officer !!!"
        Write-Host $Global:strLineSeparator
        start-sleep -seconds 15
        Break
        }  

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
} ## Authenticate Routine

    Authenticate
    
    Write-output "partnername`tostype`tosversion`tid`tcomputername`tdevicename`tpassword`tpassphrase`tdownload`tcommand" | Out-file $csvcommands
    

    Function Get-DeviceList {

# ----- Enumerate Account Statistics -----

    $BackupFilter = "at == 1"                         ## Filter to exclude never deployed devices and M365  
    $CountToLookup = 5000                             ## Maximum device count to return


# (Create the JSON object to call the EnumerateAccountStatistics function)
    $objEnumerateAccountStatistics = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
    Add-Member -PassThru NoteProperty visa $global:visa |
    Add-Member -PassThru NoteProperty method ‘EnumerateAccountStatistics’ |
    Add-Member -PassThru NoteProperty params @{
                                                query = @{
                                                                  PartnerId=[int]$partnerId
        	                                                      Filter="$BackupFilter"
                                                                  Columns=("AU","AR","AN","MN","OP","OT","OS","PF","PN","QW")
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
    

    $global:DeviceDetail = @()
    ForEach ( $Result in $EnumerateAccountStatisticsSession.result.result ) {
        $Global:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID    = [String]$Result.AccountId;
                                                                    PartnerID    = [string]$Result.PartnerId;
                                                                    DeviceName   = $Result.Settings.AN -join '' ;
                                                                    Password     = $Result.Settings.QW -join '' ;
                                                                    ComputerName = $Result.Settings.MN -join '' ;
                                                                    PartnerName  = $Result.Settings.AR -join '' ;
                                                                    Reference    = $Result.Settings.PF -join '' ;
                                                                    Account      = $Result.Settings.AU -join '' ;
                                                                    OsType       = $Result.Settings.OT -join '' ;
                                                                    OSVersion    = $Result.Settings.OS -join '' ;
                                                                    Product      = $Result.Settings.PN -join '' ;
                                                                    Profile      = $Result.Settings.OP -join '' }
        }




    $Global:Detail = $DeviceDetail | sort-object PartnerName | out-gridview -OutputMode Multiple -Title "Select all Required Devices and Click OK to Request PassPhrases"

    #$Detail = $DeviceDetail | where-object{$_.computername -ne $env:computername} 

    }

    Get-DeviceList

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
        $osversion = $d.OsVersion 
        Write-Host "  Getting Passphrase/Credentials for Device $id"
        start-sleep -Milliseconds 100
        
        $objGetPassphrase = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
        Add-Member -PassThru NoteProperty visa $global:visa |
        Add-Member -PassThru NoteProperty method ‘GenerateReinstallationPassphrase’ |
        Add-Member -PassThru NoteProperty params @{accountId=$id } |
        Add-Member -PassThru NoteProperty id ‘1’)| ConvertTo-Json -Depth 3

# (Call the JSON Web Request Function to get the GetPassphrase Object)
       
        $GetPassphraseSession = CallJSON $urlJSON $objGetPassphrase
        $Passphrase = $GetPassphraseSession.result.result

        If (($d.OSVersion -match "macos") -or ($d.OSVersion -match "OS X"))  {
            Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`t$passphrase`t$MACdownload`t$macinstaller") | Out-file $csvcommands -append
        }
        elseif (($d.OSVersion -notmatch "macos") -and ($d.OSVersion -notmatch "OS X") -and ($d.OSVersion -notmatch "Windows")) {
            Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`t$passphrase`tManual`tManual") | Out-file $csvcommands -append
        }
        elseif ($passphrase -eq $null) {
            Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`tUsing Private Encryption Key`t$download`t" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -encryption-key `"INSERT_PRIVATE_KEY`"") | Out-file $csvcommands -append
        }else{
            #Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`t$passphrase`t$download`t" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
            Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`t$passphrase`t$download;" + '& c:\windows\temp\' + "mxb-windows-x86_x64.exe -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append
            #Write-output ("$partnername`t$ostype`t$osversion`t$id`t$computername`t$devicename`t$password`t$passphrase`t$tmpDownload;" + '& $env:tmp\mxb-windows-x86_x64.exe' + " -silent -user `"$DeviceName`" -password `"$password`" -passphrase `"$passphrase`"") | Out-file $csvcommands -append

        }

    }

    Write-output $Global:strLineSeparator
    Write-output "  Redeployment Credentials Files found here:"
    Write-output $Global:strLineSeparator
    Write-output " & `"$csvcommands`""
    
    $xlscommands = $csvcommands.Replace("csv","xlsx")
    Save-CSVasExcel $csvcommands
    
    Write-output " & `"$xlscommands`""
    Write-output $Global:strLineSeparator
    
    If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
        Start-Process "$xlscommands"
        }else{
        Start-Process "$csvcommands"
        }
    start-sleep -seconds 15