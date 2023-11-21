<# ----- About: ----
    # Bulk Delete Orphaned Documents Devices
    # Revision v3 - 2023-09-03
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
    # For use with N-able | Cove Data Protection 
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Authenticate to https://backup.management console
    # Enumerate Documents devices
    # Select orphaned Documents devices with a Timestamp > X and Creation Date > X for deletion
    #    
    # Use the -DeviceTimeStamp ## parameter to query devices with a Timestamp older than XX days ago
    # Use the -DeviceCreation ## parameter to query devices with a Creation Date older than XX days ago
    # Use the -DeviceCount ## (default=1000) parameter to limit the maximum device count to return
    # Use the -CoveLoginPartnerName to specify the case sensitive Cove Customer Name for API authentication
    # Use the -CoveLoginUserName to specify the Cove User Name for API authentication
    # Use the -CovePlainTextPassword to specify the Cove Password for API authentication
    # Use the -ForceDelete switch parameter to force deletion without additional confirmation
 
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)][String]$CoveLoginPartnerName,                                ##  Remove / replace with prompt for AMP Usage
        [Parameter(Mandatory=$False)][String]$CoveLoginUserName,                                   ##  Remove / replace with prompt for AMP Usage   
        [Parameter(Mandatory=$False)][String]$CovePlainTextPassword,                               ##  Remove / replace with prompt for AMP Usage
        [Parameter(Mandatory=$False)][string]$TaskPath = "c:\ProgramData\BackupNerdScripts",        ##  Base Path to Store Script Credentials and Logs
        [Parameter(Mandatory=$False)][ValidateRange(30,[int]::MaxValue)][Int]$DeviceTimeStamp=60,   ##  Timestamp older than XX days ago
        [Parameter(Mandatory=$False)][ValidateRange(60,[int]::MaxValue)][Int]$DeviceCreation=90,    ##  Creation older than XX days ago
        [Parameter(Mandatory=$False)][ValidateRange(1,[int]::MaxValue)][Int]$DeviceCount=1000,      ##  Maximum device count to return
        [Parameter(Mandatory=$False)][Switch]$ForceDelete
        
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    [int]$DeviceTimeStamp = 60
    [int]$DeviceCreation = 90
    [int]$DeviceCount = 3000
    [string]$TaskPath = "c:\ProgramData\BackupNerdScripts"

    ## - Remove Credential to be prompted or passed from Script or AMP
    [string]$CoveLoginPartnerName = $null
    [string]$CoveLoginUserName = $null
    [string]$CovePlainTextPassword = $null
    ## - Remove Credential to be prompted or passed from Script or AMP
    
    ## - Remove or comment out for use in AMP
    #Requires -Version 5.1 -RunAsAdministrator
    $ConsoleTitle = "Bulk Delete Orphaned Documents Devices"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    ## - Remove or comment out for use in AMP
    

    #$ScriptParent = $PSScriptRoot
    $ScriptPath = $PSCommandPath ; # Write-host "Scriptpath = $ScriptPath"
    $ScriptLeaf = Split-Path $ScriptPath -leaf ; # Write-host "Scriptleaf = $ScriptLeaf"
    $ScriptBase = $Scriptleaf -replace '\..*' ; # Write-host "ScriptBase = $ScriptBase"
    Split-Path $ScriptPath | Push-Location
    $ScriptLog = Join-Path -Path $Taskpath -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.log.csv"
    $ScriptLogParent = Split-path -path $ScriptLog
    if (Test-Path $ScriptLogParent){Write-output "ScriptLogPath $ScriptLogParent Exist" }Else{mkdir -Force $ScriptLogParent}
       
    Write-Output "$ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "Script Parameter Syntax:`n`n  $Syntax"

    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"

    Write-output "  Current Parameters:"
    Write-output "  -DeviceTimeStamp    = $DeviceTimeStamp"
    Write-output "  -DeviceCreation     = $DeviceCreation"
    Write-output "  -DeviceCount        = $DeviceCount"
    Write-output "  -TaskPath           = $TaskPath"
    Write-output "  -LogPath            = $ScriptLog"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

    Function Authenticate {

        Write-Output $script:strLineSeparator
        Write-Output "  Enter Your N-able | Cove https:\\backup.management Login Credentials"
        Write-Output $script:strLineSeparator

        if ($CoveLoginPartnerName -eq "") {$CoveLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"}
        if ($CoveLoginUserName -eq "") {$CoveLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able | Cove Backup.Management API"}
        if ($CovePlainTextPassword -eq "") {
            $CovePassword = Read-Host -AsSecureString "  Enter Password for N-able | Cove Backup.Management API"
            # (Convert SecureString Password to plain text)
            $CovePlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($CovePassword))
            }

        # (Show credentials for Debugging)
        Write-Output "  Logging on with the following Credentials`n"
        Write-Output "  PartnerName:  $CoveLoginPartnerName"
        Write-Output "  UserName:     $CoveLoginUserName"
        Write-Output "  Password:     It's secure..."

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.method = 'Login'
        $data.params = @{}
        $data.params.name = [String]$CoveLoginPartnerName
        $data.params.username = [String]$CoveLoginUserName
        $data.params.password = [String]$CovePlainTextPassword

        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

        $Script:Session = Invoke-RestMethod @params 
      
    # (Variable to hold current visa and reused in following routines)
        $script:visa = $session.visa
        $script:PartnerId = [int]$session.result.result.PartnerId
            
    # (Get Result Status of Authentication)
        $AuthenticationErrorCode = $Session.error.code
        $AuthenticationErrorMsg = $Session.error.message

    # (Check if ErrorCode has a value)
        If ($AuthenticationErrorCode) {
            Write-Output "Authentication Error Code:  $AuthenticationErrorCode"
            Write-Output "Authentication Error Message:  $AuthenticationErrorMsg"
            Pause
            Break Script
        }# (Exit Script if there is a problem)
        Else {

        } # (No error)
        Write-Output $Script:strLineSeparator
        Write-Output "" 


    # (Print Visa to screen)
        #Write-Output $script:strLineSeparator
        #Write-Output "Current Visa is: $script:visa"
        #Write-Output $script:strLineSeparator

    ## Authenticate Routine
    }  ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Authenticate
            }
        }
    }
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

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

    Function Send-GetPartnerInfo ($PartnerName) { 

        $RestrictedPartnerLevel = @("Root","SubRoot")
            
        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'GetPartnerInfo'
        $data.params = @{}
        $data.params.name = [String]$PartnerName

        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            WebSession  = $websession
            ContentType = 'application/json; charset=utf-8'
        }   

            $Script:Partner = Invoke-RestMethod @params 

        if ($RestrictedPartnerLevel -notcontains $Partner.result.result.Level) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-Output "  $Level - $PartnerName - $partnerId - $Uid"
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

    Function Send-EnumerateDevices {

        #$DeviceFilter = "US == 0 AND T3 == 0"
        $Script:DeviceFilter = "OI == 5038 AND OP == 'Documents' AND PD == 21266 AND PN == 'Documents' AND TS < $DeviceTimeStamp.days().ago() AND CD < $DeviceCreation.days().ago()"
        # Note: < > signs in device filters may appear reversed due to counting backwards Epoch time backwards
        

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.Filter = $DeviceFilter
        $data.params.query.Columns = @("AU","TS","CD","TL","AR","AN","AL","LN","OP","MN","OI","OS","PD","AP","PF","PN","US","T3","AA843","AA77","T7")
        $data.params.query.OrderBy = "TS ASC"
        $data.params.query.SelectionMode = "Merged"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $DeviceCount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            ContentType = 'application/json; charset=utf-8'
        }  

        $Script:DocumentsDevices = Invoke-RestMethod @params 
   
        if ($Partner.result.result.Uid) {
            [String]$Script:PartnerId = $Partner.result.result.Id
            [String]$Script:PartnerName = $Partner.result.result.Name
    
            Write-Output $Script:strLineSeparator
            Write-output "  Searching  $PartnerName "
            Write-Output $Script:strLineSeparator
            }else{
            Write-Output $Script:strLineSeparator
            Write-Host "PartnerName Not Found"
            Write-Output $Script:strLineSeparator
            }
             
        $Script:DeviceDetail = @()

        if ($DocumentsDevices.result.result.count -eq 0) {
            Write-Output "No Orphaned Documents Devices found greater that $DocumentsAge Days`n  Exiting Script"
            Start-Sleep -seconds 10
            Break
        }
        
        Write-Output "  Requesting details for $($DocumentsDevices.result.result.count) Orphaned Documents devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        ForEach ( $DeviceResult in $DocumentsDevices.result.result ) {
            Get-VisaTime
            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{
                LogTime        = $CurrentDate ;
                Action         = "Deletion" ;
                AccountID      = [String]$DeviceResult.AccountId;
                PartnerID      = [string]$DeviceResult.PartnerId;
                ComputerName   = $DeviceResult.Settings.MN -join '' ;
                DeviceName     = $DeviceResult.Settings.AN -join '' ;
                DeviceAlias    = $DeviceResult.Settings.AL -join '' ;
                PartnerName    = $DeviceResult.Settings.AR -join '' ;
                Reference      = $DeviceResult.Settings.PF -join '' ;
                DataSources    = $DeviceResult.Settings.AP -join '' ;
                Account        = $DeviceResult.Settings.AU -join '' ;
                Location       = $DeviceResult.Settings.LN -join '' ;
                Notes          = $DeviceResult.Settings.AA843 -join '' ;
                TempInfo       = $DeviceResult.Settings.AA77 -join '' ;
                OS             = $DeviceResult.Settings.OS -join '' ;
                Creation       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                TimeStamp      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;  
                LastSuccess    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '') ;
                SelectedGB     = [math]::Round([Decimal](($DeviceResult.Settings.T3 -join '') /1GB),2) ;  
                UsedGB         = [math]::Round([Decimal](($DeviceResult.Settings.US -join '') /1GB),2) ;
                Errors         = $DeviceResult.Settings.T7 -join '' ;
                Errors_FS      = $DeviceResult.Settings.F7 -join '' ;
                Product        = $DeviceResult.Settings.PN -join '' ;
                ProductID      = $DeviceResult.Settings.PD -join '' ;
                Profile        = $DeviceResult.Settings.OP -join '' ;
                ProfileID      = $DeviceResult.Settings.OI -join '' ;
                Filter         = $DeviceFilter
            }
        }   
    }

    Function Send-RemoveAccount ([int]$accountidtodelete) {
        Get-VisaTime
        Write-Output "Removing DeviceId $($Deviceidtodelete.AccountID) | DeviceName $($Deviceidtodelete.DeviceName) | ComputerName $($Deviceidtodelete.ComputerName) | Partner $($Deviceidtodelete.PartnerName) | Product $($Deviceidtodelete.Product) | Last Timestamp $($Deviceidtodelete.TimeStamp)"

        $url = "https://api.backup.management/jsonapi"
        $method = 'POST'
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $script:visa
        $data.method = 'RemoveAccount'
        $data.params = @{accountId=$accountidtodelete}
        $jsondata = (ConvertTo-Json $data -depth 6)

        $params = @{
            Uri         = $url
            Method      = $method
            Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
            ContentType = 'application/json; charset=utf-8'
        }  

        $Script:RemoveAccount = Invoke-RestMethod @params

        if ($Script:RemoveAccount.error) {
            write-warning "$($Script:RemoveAccount.error.message)"
        }else{
            $Deviceidtodelete | Select-Object LogTime,Action,AccountID,Creation,LastSuccess,Timestamp,PartnerName,Reference,DeviceName,ComputerName,DeviceAlias,Location,DataSources,TempInfo,SelectedGB,UsedGB,Profile,ProfileID,Product,ProductID,Errors,Filter | Export-Csv -Path $ScriptLog -Append -NoTypeInformation
        }
    }

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    Authenticate
    Send-GetPartnerInfo $CoveLoginPartnerName
    Send-EnumerateDevices

    if ($FORCEDELETE) {
        $SelectedDevices = $Script:DeviceDetail

        if ($null -eq $SelectedDevices) {
            Write-Warning "No selections made"
            Break
        }

    }ELSE{
        $SelectedDevices = $Script:DeviceDetail | Select-Object LogTime,Action,AccountID,Creation,LastSuccess,Timestamp,PartnerName,Reference,DeviceName,ComputerName,DeviceAlias,Location,DataSources,TempInfo,SelectedGB,UsedGB,Profile,ProfileID,Product,ProductID,Errors,Filter| Out-GridView -Title "$($DocumentsDevices.result.result.count)  | Orphaned Document Devices (Greater Than $DeviceCreation Days Since Creation and $DeviceTimeStamp Days Since Last Backup Agent Check-in) | Select Devices to Purge" -OutputMode Multiple
        
        if ($null -eq $SelectedDevices) {
            Write-Warning "No selections made"
            Break
        }

    }

    if ($SelectedDevices) {

        $SelectedDevice | Select-Object * | Export-Csv -Path $ScriptLog

        Foreach ($Deviceidtodelete in $selectedDevices) {
            Send-RemoveAccount $Deviceidtodelete.accountId
        }
        Write-Warning "In case of accidental device deletion it may be possible for Cove technical support to undelete devices within 28 days."
    }

