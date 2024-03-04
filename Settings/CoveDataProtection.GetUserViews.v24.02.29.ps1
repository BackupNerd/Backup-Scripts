<# ----- About: ----
    # Cove Data Protection | Get User Views
    # Revision v24.02.29
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com
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
    # For use with the standalone edition of N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Get | Set Secure Credentials & Authenticate
    # Enumerate and Export Select User Views
    # Optionally Import Select User Views
    #
    # Use the -Export switch parameter to export select views to .JSON output files 
    # Use the -ExportPath (?:\Folder) parameter to specify .JSON output file path
    # Use the -Import switch parameter to import select views from .JSON output files to the current authenticated user
    # Use the -TaskPath (?:\Folder) parameter to specify where to store script credentials, task & logs
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # API & script authentication info and related document can be found at the URLS below:
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/login.htm
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
       
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Export")]
    Param (
        [Parameter(ParameterSetName="Export",Mandatory=$False)] [switch]$Export,                      ## Save User Views as .JSON output files
        [Parameter(ParameterSetName="Import",Mandatory=$False)] [switch]$Import,                      ## Browse and Import User View from .JSON files
        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",                          ## Export Path (Script location is the default)
        [Parameter(Mandatory=$False)] [string]$TaskPath = "$env:userprofile\BackupNerdScripts",       ## Path to Store/Invoke Scheduled Backup Nerd Script Credentials and Tasks
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                       ## Remove stored API credentials at script run
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    #Requires -Version 5.1
    $ConsoleTitle   = "Cove Data Protection | Get User Views"                       ## Update with full script name
    $ShortTitle     = "GetUserViews"                                                ## Update with short script name
    $host.UI.RawUI.WindowTitle = $ConsoleTitle

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

    Write-Output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "`n  Script Parameter Syntax:`n`n  $Syntax"
       
    $CurrentDate    = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
    $ShortDate      = Get-Date -format "yyyy-MM-dd"
   
    if ($ExportPath) {
        $ExportPath = Join-path -path $ExportPath -childpath "$($ShortTitle)_$($ShortDate)" 
    }else{
        $ExportPath = Join-path -path $dir -childpath "$($ShortTitle)_$($ShortDate)"
    }

    If ($ExportPath) {mkdir -force -path $ExportPath | Out-Null}

    Write-Output "  Current Parameters:"
    Write-Output "  -Import         = $Import"
    Write-Output "  -Export         = $Export"
    Write-Output "  -ExportPath     = $ExportPath"
    Write-Output "  -TaskPath       = $TaskPath"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'

    #$Script:True_path = "C:\ProgramData\MXB\"
    #$Script:MXB_Path = "C:\ProgramData\MXB\"
    #$script:BackupScripts = "C:\ProgramData\BackupNerdScripts"

    $Script:APIcredfile = join-path -Path $Taskpath -ChildPath "$env:computername $env:username CDP_API_Credentials.Secure.enc"
    $Script:APIcredpath = Split-path -path $Script:APIcredfile

    #$filterDate = (Get-Date).AddDays(-$Days)
    #$counter = 1

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Set-APICredentials {

        Write-Output $Script:strLineSeparator 
        Write-Output "  Setting Backup API Credentials" 
        if (Test-Path $Script:APIcredpath) {
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $Script:APIcredpath} 
        Write-Output $Script:strLineSeparator     
        Write-Output "  Enter Exact, Case Sensitive Customer Name for N-able Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"

        DO{ $PartnerName = Read-Host "  Enter Login Customer Name" }
        WHILE ($partnerName.length -eq 0)
        
        $PartnerName | out-file $Script:APIcredfile

        $BackupCred = Get-Credential -UserName "" -Message 'Enter Login Email and Password for N-able Backup.Management API'
        $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

        $BackupCred.UserName | Out-file -append $Script:APIcredfile
        $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $Script:APIcredfile
        
        Authenticate  ## Attempt API Authentication

    } ## Set API credentials if not present

    Function Get-APICredentials {

        $Script:APIcredfile = join-path -Path $Taskpath -ChildPath "$env:computername $env:username CDP_API_Credentials.Secure.enc"

        $Script:APIcredpath = Split-path -path $script:APIcredfile

        if (($ClearCredentials) -and (Test-Path $Script:APIcredfile)) { 
            Remove-Item -Path $Script:APIcredfile
            $ClearCredentials = $Null
            Write-Output $Script:strLineSeparator 
            Write-Output "  Backup API Credential File Cleared"
            Authenticate  ## Retry Authentication
            
        }else{ 
            Write-Output $Script:strLineSeparator 
            Write-Output "  Getting Backup API Credentials" 
        
            if (Test-Path $Script:APIcredfile) {
                Write-Output    $Script:strLineSeparator        
                "  Backup API Credential File Present"
                $APIcredentials = get-content $Script:APIcredfile
                
                $Script:cred0 = [string]$APIcredentials[0] 
                $Script:cred1 = [string]$APIcredentials[1]
                $Script:cred2 = $APIcredentials[2] | Convertto-SecureString 
                $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))

                Write-Output    $Script:strLineSeparator 
                Write-Output "  Stored Backup API Customer = $Script:cred0"
                Write-Output "  Stored Backup API User     = $Script:cred1"
                Write-Output "  Stored Backup API Password = Encrypted"
                
            }else{
                Write-Output    $Script:strLineSeparator 
                Write-Output "  Backup API Credential File Not Present"

                Set-APICredentials  ## Create API Credential File if Not Found
            }
        }
    } ## Get API credentials if present

    Function Authenticate {
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

        #Debug Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

        if ($authenticate.visa) { 
            $Script:visa = $authenticate.visa
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-Output "  Authentication Failed: Please confirm your Backup.Management Customer Name and Credentials"
            Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            Write-Warning $authenticate.error.message

            Set-APICredentials  ## Create API Credential File if Authentication Fails
        }
    } ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Authenticate
            }
        }
    } ## Check Visa Time and Authenticate again if needed

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $epoch = $epoch.AddSeconds($inputUnixTime)
        return $epoch
        }else{ return ""}
    } ## Convert epoch time to date time 

    Function Get-LogTimeStamp {
        #Get-Date -format s
        return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    } ## Output proper timeStamp for log file

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

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function CallJSON($url,$object) {
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
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
            Write-Output "  $PartnerName - $partnerId - $Uid"
            Write-Output $Script:strLineSeparator
        }else{
            Write-Output $Script:strLineSeparator
            Write-Host "  Lookup for Root, Sub-root and Distributor Partner Level Not Allowed"
            Write-Output $Script:strLineSeparator
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername
        }

        if ($partner.error) {
            Write-Output "  $($partner.error.message)"
            $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
            Send-GetPartnerInfo $Script:partnername

        }

    } ## get PartnerID and Partner Level

    Function Send-EnumerateUsers ($PartnerID) { 

        Write-Output "  Enumerating users"    
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateUsers'
        $data.params = @{}
        $data.params.partnerIds = [array]$PartnerID
    
        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:EnumerateUsers = $webrequest | convertfrom-json
    
            $Script:SelectedUsers = $Script:EnumerateUsers.result.result | Select-Object Partnerid,Name,emailAddress,FullName,id,Roleid,Title,Flags,PhoneNumber,* -ea SilentlyContinue | Out-GridView -Title "Current Partner | $Script:partnername | Select Users" -OutputMode Multiple 
    
    }  ## List all users    

    Function Send-EnumerateUserSettings ([int]$userId) { 

        Write-Output "  Enumerating user views"   
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateUserSettings'
        $data.params = @{}
        $data.params.userId = $userId
    
        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:EnumerateUserSettings = $webrequest | convertfrom-json   
            $Script:EnumerateUserSettings.error.message
            if ( $Script:EnumerateUserSettings -ne $null ) {
                Write-Output "  $($Script:EnumerateUserSettings.result.result.Count) total views found"
            }

            $Script:SelectedViews = $EnumerateUserSettings.result.result | Where-Object {$_.Type -eq "Custom"} | Select-Object UserId,Id,Name,Type,View | Out-GridView -Title "Current Partner | $Script:partnername | $($selecteduser.emailaddress) | Select 'Custom' User Views to Export to (.JSON) Files" -OutputMode Multiple

            if ( $Script:SelectedViews -eq $null ) {
                Write-Warning "0 Custom Views Exist or 0 Views Selected to Export"
            }

            foreach ($selectedView in $Script:Selectedviews) {

                $selectedView | Select-Object * | ConvertTo-Json -depth 10 | out-file "$($ExportPath)\User_$($selectedView.UserId)_$(($SelectedUser.emailaddress -split("@"))[0])_View_$($selectedView.Id)_$($selectedView.name -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9 -]`",`"`")).json"

            }
    
    }  ## Get User Settings

    Function Send-AddUserSettings ($SelectedView) { 

        Write-Output "  Attempting to Import User View | $($SelectedView.Name) | to User | $($Authenticate.result.result.emailAddress)"    
                    
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'AddUserSettings'
        $data.params = @{}
        $data.params.userSettings = @{}
        $data.params.userSettings.UserId = $Authenticate.result.result.id
        $data.params.userSettings.Name = $SelectedView.Name 
        $data.params.userSettings.Type = $SelectedView.Type
        $data.params.userSettings.View = $SelectedView.View

         $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            #$Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:AddUserSettings = $webrequest | convertfrom-json  
           if ($AddUserSettings.error) {
            $adderror = [System.Web.HttpUtility]::HtmlDecode($AddUserSettings.error.message)
            write-warning $adderror
           }else{$AddUserSettings.result.result}
    }  ## Add User Settings

    Function Open-FileName($initialDirectory) {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Exported Cove User Views (*.json)|*.json"
        $OpenFileDialog.title = "Select an Exported Cove User View to Import (*.json)"
        $result = $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        #$OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    } ## GUI Prompt for Filename to open



#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

Switch ($PSCmdlet.ParameterSetName) {
    'Import' {

        Authenticate
        $OpenFileName = Open-FileName "$ExportPath" ; if ($null -eq $OpenFileName) {Break} 
        $SelectedView = Get-Content -Path "$OpenFileName" | Out-String | ConvertFrom-Json
        Send-AddUserSettings $selectedview
    } 
    'Export' {

        $Script:DeviceDetail = @()

        Authenticate
    
        Write-Output $Script:strLineSeparator
        Write-Output "" 
        
        Send-GetPartnerInfo $Script:cred0
        
        if ($ExportPath) {
            $ExportPath = Join-path -path $ExportPath -childpath "$($partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9-]`",`"`"))" 
            mkdir -force -path $ExportPath | Out-Null
        }
    
        if (($Authenticate.result.result.partnerId -eq 1) -or ($DefaultPartner)) {

            Send-EnumerateUsers $Script:partnerid
    
            foreach ($selecteduser in $selectedusers) {

                Send-EnumerateUserSettings $selecteduser.id
           }      
        
        }
        else{
            $selecteduser = $Authenticate.result.result
            Send-EnumerateUserSettings $selecteduser.id
        }
    
        Write-Output $Script:strLineSeparator
        Write-Output "  View Export Path = $exportpath"
        Write-Output ""
        Start-Sleep -seconds 1

    }
}
