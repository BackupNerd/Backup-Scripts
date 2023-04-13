<# ----- About: ----
    # Get N-able Backup Selections
    # Revision v03 - 2023-04-04
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
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Enumerate & Post Fixed Disk Volume Letters and Used Size to custom column AA2646 in the https://backup.management console  
    # Enumerate & Post Inclusions to custom column AA2647 in the https://backup.management console  
    # Enumerate & Post Exclusions to custom column AA2648 in the https://backup.management console  
    # Enumerate & Post Filters to custom column AA2649 in the https://backup.management console 
    # Enumerate & Post USB Volumes to custom column AA2650 in the https://backup.management console 
    # 
    #   Note: Partner must add custom column AA2646-AA2650 to view data in the https://backup.management console
    # 
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/console/custom-columns.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
       
        [Parameter(Mandatory=$False)] [string]$TaskPath = "$env:userprofile\BackupNerdScripts", ## Path to Store/Invoke Scheduled Backup Nerd Script Credentials and Tasks
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                 ## Remove Stored API Credentials at start of script
    )   

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get Cove Backup Selections"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $dir = Split-Path $scriptpath
    Push-Location $dir
    $CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyy-MM-dd"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'


#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

  #Requires -Version 5.1 
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $script:strLineSeparator = "  ---------"
    $urlJSON = 'https://api.backup.management/jsonapi'
    $clienttool = "c:\program files\backup manager\clienttool.exe"
    $script:True_path = "C:\ProgramData\MXB\"

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

    Function Send-APICredentialsCookie {
            # Speak "Connecting via Secure A.P.I."
              ## Read API Credential File before Authentication
        
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
                $script:roleId = $authenticate.result.result.roleid
                }
        
        }  ## Use Backup.Management credentials to Authenticate
    Function Authenticate {

        Write-Output $script:strLineSeparator
        Write-Output "  Enter Your N-able | Cove https:\\backup.management Login Credentials"
        Write-Output $script:strLineSeparator

        $AdminLoginPartnerName = ""
        $AdminLoginUserName = ""

        if ($AdminLoginPartnerName -eq $null) {$AdminLoginPartnerName = Read-Host -Prompt "  Enter Exact, Case Sensitive Partner Name for N-able | Cove Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"}
        if ($AdminLoginUserName -eq $null) {$AdminLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able | Cove Backup.Management API"}

        if ($PlainTextAdminPassword -eq $null) {
            $AdminPassword = Read-Host -AsSecureString "  Enter Password for N-able | Cove Backup.Management API"
            # (Convert SecureString Password to plain text)
            $PlainTextAdminPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdminPassword))
            }
    
        # (Show credentials for Debugging)
        Write-Output "  Logging on with the following Credentials`n"
        Write-Output "  PartnerName:  $AdminLoginPartnerName"
        Write-Output "  UserName:     $AdminLoginUserName"
        Write-Output "  Password:     It's secure..."

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

    # (Print Visa to screen)
        #Write-Output $script:strLineSeparator
        #Write-Output "Current Visa is: $script:visa"
        #Write-Output $script:strLineSeparator

    ## Authenticate Routine
        }  ## Use Backup.Management credentials to Authenticate
 
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

        $RestrictedPartnerLevel = @("Root","SubRoot"
                                    )
            
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

        if ($RestrictedPartnerLevel -notcontains $Partner.result.result.Level) {
            [String]$Script:Uid = $Partner.result.result.Uid
            [int]$Script:PartnerId = [int]$Partner.result.result.Id
            [String]$script:Level = $Partner.result.result.Level
            [String]$Script:PartnerName = $Partner.result.result.Name

            Write-Output $Script:strLineSeparator
            Write-output "  $Level - $PartnerName - $partnerId - $Uid"
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

    Function Send-UpdateCustomColumn($DeviceId,$ColumnId,$Message) {

        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Content-Type", "application/json")
        $headers.Add("Cookie", "__cfduid=d110201d75658c43f9730368d03320d0f1601993342")
        $headers.Add("Authorization", "Bearer $script:visa")
        
        $body = "{
        `n      `"jsonrpc`":`"2.0`",
        `n      `"id`":`"jsonrpc`",
        `n      `"method`":`"UpdateAccountCustomColumnValues`",
        `n      `"params`":{
        `n      `"accountId`": $DeviceId,
        `n      `"values`": [[$ColumnId,`"$Message`"]]
        `n      }
        `n  }
        `n"
        
        $script:updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body
        $script:updateCC.error.message
        }
        
#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    Function RND {
        Param(
        [Parameter(ValueFromPipeline,Position=3)]$Value,
        [Parameter(Position=0)][string]$unit = "MB",
        [Parameter(Position=1)][int]$decimal = 2

        )
        "$([math]::Round(($Value/"1$Unit"),$decimal)) $Unit"

    <# Usage Examples

    1.23123123123123123 | RND '' 6
    1.231231 

    RND KB 2 234234234.234234234234
    228744.37 KB

    234234234.234234234234 | RND KB 2
    228744.37 KB

    234234234.234234234234 | RND MB 4
    223.3832 MB

    234234234.234234234234 | RND GB 1
    0.2 GB

    234234234.234234234234 | RND 
    0.22 GB

    234234234.234234234234 | RND MB 0
    223 MB

    1223234234234.234234234234 | RND TB 2
    1.11 TB

    write-output "12312312312123.123123123" | RND KB
    12023742492.31 KB

    write-output "TEST $(12312312312123.12312312 | RND GB 2)" 
    TEST 12023742492.31 KB
    #>

    } ## Rounding function for B,KB,MB,GB,TB


    Function Get-FixedVolumes {
        $Volumes = Get-Volume

        $FixedVolumes = $Volumes | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.DriveLetter)}

        $FixedVolumes1 = $FixedVolumes | Select-Object @{l='Letter';e={$_.DriveLetter + ":"}},@{l='Type';e={$_.DriveType}},@{l='TotalGB';e={([Math]::Round(($_.Size/1GB),2))}},@{l='FreeGB';e={([Math]::Round(($_.SizeRemaining/1GB),2))}},@{l='UsedGB';e={([Math]::Round((($_.Size - $_.SizeRemaining)/1GB),2))}}
        
        $FixedVolumes2 = $FixedVolumes1 | Select-Object Letter,Type,UsedGB,@{l='String';e={$_.Letter + ' ' + $_.UsedGB +' GB'}}

        $FixedVolumeString = $FixedVolumes2.String -join " | "

        Send-UpdateCustomColumn $Script:Instance.AccountId 2646 "$FixedVolumeString"
        Write-output "  Fixed Volumes = $FixedVolumeString"
    }
    Function Get-Inclusions {

        $Script:Include = $Selections | Where-Object {(($_.type -eq "Inclusive") -and ($_.DSRC -eq "FileSystem")) }

        if ($Script:Include) {
            if ($Script:include[0].path -eq "") {$Script:IncludeBase = "FileSystem"}else{$Script:IncludeBase = $null}
            $Script:IncludeString = $Script:Include.path.replace("\","\\") -join " | "
        }else{
            $Script:IncludeBase = $null
            $Script:IncludeString = "-"
        }

        Send-UpdateCustomColumn $Script:Instance.AccountId 2647 "$Script:IncludeBase $Script:IncludeString"
        Write-output "  Inclusions = $Script:IncludeBase $($Script:IncludeString.replace("\\","\"))"
    }
    Function Get-Exclusions {    

        $Script:Exclude = $Selections | Where-Object {(($_.type -eq "Exclusive") -and ($_.DSRC -eq "FileSystem")) }

        if ($Script:Exclude) {
            if ($Script:include[0].path -ne "") {
                $Script:ExcludeBase = "FileSystem"
            }else{
                $Script:ExcludeBase = $null
            }
            $Script:ExcludeString = $Script:Exclude.path.replace("\","\\") -join " | "
        }else{
            if (($Script:Include) -and ($Script:include[0].path -ne "")){
                $Script:ExcludeBase = "FileSystem"
                $Script:ExcludeString = $null
            }else{
                $Script:ExcludeBase = $null
                $Script:ExcludeString = "-"
            }
        }

        Send-UpdateCustomColumn $Script:Instance.AccountId 2648 "$Script:ExcludeString"
        Write-output "  Exclusions = $Script:ExcludeBase$Script:ExcludeString"
    }
    Function Get-Filters {    
    
        & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.filter.list | out-file C:\programdata\mxb\filters.csv

        $Filters = import-csv -path C:\programdata\mxb\filters.csv -Header value

        if ($Filters) {
                 
            $FilterString = $Filters.value.replace("\","\\") -join " | "
            
            Send-UpdateCustomColumn $Script:Instance.AccountId 2649 "$FilterString"
            Write-output "  Filters = $($FilterString.replace("\\","\"))"
            
        }

    }
    Function Get-USBVolumes {
    # ----- Get Disk Partion Letter for USB / Non USB Bustype ----
        Get-Disk | Select-Object Number | Update-Disk
        $Disk = Get-Disk | Where-Object -FilterScript {$_.Bustype -eq "USB"} | Select-Object Number
    # (Exclude Null Partition Drive Letters)
    if ($disk -ne $null) {
         
            $USBvol = Get-Partition -DiskNumber $Disk.Number | Where-Object {$_.DriveLetter -ne "`0"} | Select-Object @{name="DriveLetter"; expression={$_.DriveLetter+":\"}} | Sort-Object DriveLetter
        # ----- End Get Disk Partion Letter for USB / Non USB Bustypes ----
            if ($USBvol) {
                $USBvolString = $USBvol.driveletter.replace("\","\\") -join " | " 
                Send-UpdateCustomColumn $Script:Instance.AccountId 2650 "$USBvolString"
                Write-output "  USB Volumes = $($USBvolString.replace("\\","\")) "
            }
        }
    }
  
    Function Get-DiskTypes {
        # ----- Get Disk Info ----

            [array]$DiskTypes = get-disk | select-object @{name="Id"; expression={$_.Disknumber}},@{name="Size"; expression={$_.size | RND GB 0}},@{name="Type"; expression={$_.partitionstyle}},@{name="Bus"; expression={$_.BusType}},@{name="Boot"; expression={$_.isboot -replace("True","Boot")-replace("False","")}},@{name="System"; expression={$_.issystem -replace("True","Sys")-replace("False","")}},@{name="Letter"; expression={}} | Sort-Object Id
            $disktypes | Format-Table
            $USBvol = Get-Partition -DiskNumber $Disk.Number | Where-Object {$_.DriveLetter -ne "`0"} | Select-Object @{name="DriveLetter"; expression={$_.DriveLetter+":\"}} | Sort-Object DriveLetter
        # ----- End Get Disk Partion Letter for USB / Non USB Bustypes ----
            if ($USBvol) {
                $USBvolString = $USBvol.driveletter.replace("\","\\") -join " | " 
                Send-UpdateCustomColumn $Script:Instance.AccountId 2650 "$USBvolString"
                Write-output "  USB Volumes = $($USBvolString.replace("\\","\")) "
            }
        }

    Authenticate

    Do {

        $BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
        $BackupStatus = & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.status.get
        $status = @("Idle","Scanning")
        if (($BackupService.status -ne "Running") -or ($status -notcontains $BackupStatus)) { 
            Write-Warning "Backup Manager Not Ready"
            Write-Output "  Service $($BackupService.status)"
            Write-Output "  Job $backupstatus"
            Write-Output "  Retrying"
            Start-Sleep -seconds 120 
        }else{   
            if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) {
                Write-Warning "Backup Manager Not Running" 
            }else{ 
                try { 
                    $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.selection.list -delimiter "," | out-file C:\programdata\mxb\selections.csv
                    
                    $Script:Selections = import-csv -path C:\programdata\mxb\selections.csv

                    $Instances = get-childitem -Recurse -path c:\programdata\mxb *.info -file | Where-Object {$_.lastwritetime -gt (Get-Date).AddDays(-30)} | Select-Object Name,Directory,lastwritetime | Sort-Object lastwritetime -Descending  
        
                    $InstanceInfo = join-path -Path $Instances[0].Directory -ChildPath $Instances[0].Name
        
                    $Script:Instance = Get-Content -Path $InstanceInfo | Out-String | Convertfrom-json
        
                    write-output "  ID $($Script:Instance.AccountId)"

                    $profileInfo = $InstanceInfo.replace("info","profile") 

                    $Profile = Get-Content -Path $ProfileInfo | Out-String | Convertfrom-json

                    if ($Profile) {
                        $profileDetail = $profile.profileData.BackupDataSourceSettings | where {$_.datasource -like "*fileSystem"}
                        
                        $profileDetail[0].DataSource
                        $profileDetail[0].Policy
                        $profileDetail[0].SelectionCollection.Selection
                        $profileDetail[0].SelectionModification
                        $profileDetail[0].ExclusionFilter
                    }

                    Get-FixedVolumes
                    Get-Inclusions
                    Get-Exclusions
                    Remove-Item C:\programdata\mxb\selections.csv
                    Get-Filters
                    Remove-Item C:\programdata\mxb\filters.csv
                    Get-USBVolumes
                    
                }catch{ 
                    Write-Warning "Oops: $_" 
                }
            }
        }
    }until (($BackupService.status -eq "Running") -and (($status -contains $BackupStatus)))

### Additional checks

    Write-output "Items that can be protected ??"
    get-Volume | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.Driveletter)} | Format-Table

    ## Fixed Volumes + Operationsl + Drive letter

    Write-output "Items not able to be protected ??"
    get-Volume | Where-Object {($_.DriveType -ne "Fixed") -or ($_.OperationalStatus -ne "OK") -or ($_.DriveLetter -eq $null)} | Format-Table

    ## Unprotected Removable, Non Operational, or No Drive Letter