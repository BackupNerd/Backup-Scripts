param(
    [Alias("apiuser")] [Parameter(Mandatory=$true)] [string]$username = (Read-Host "Enter username"),
    [Alias("token")][Parameter(Mandatory=$true)] [Security.SecureString]$password = (Read-Host "Enter password" -AsSecureString)

    ## Strip out secure credential prompt from parm list to allow for script to be passed credentials when run 
    ## via RMM / MDM / Automation (i.e. N-able Automation Manager, ConnectWise Automate, etc.
)

clear-host
Push-Location (Split-Path $MyInvocation.MyCommand.Path)

#region ----- Authentication ----

Function CDP-Authenticate {

    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
    $plaintextpassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.username = $username
    $data.params.password = $plaintextpassword # $password

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url `
        -SessionVariable Script:websession `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Authenticate = $webrequest | convertfrom-json

    if ($authenticate.visa) { 
        $Script:visa = $authenticate.visa
        Write-Output "Authenticated as $($Authenticate.result.result.emailaddress)"
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Customer Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
        Write-Output    $Script:strLineSeparator 
        Write-Output $authenticate.error.message
        break
    }
} ## Use Backup.Management credentials to Authenticate
    
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

Function RND {
    Param(
    [Parameter(ValueFromPipeline,Position=3)]$Value,
    [Parameter(Position=0)][string]$unit = "MB",
    [Parameter(Position=1)][int]$decimal = 2
    )
    "$([math]::Round(($Value/"1$Unit"),$decimal)) $Unit"

<# Usage Examples
    1.23123123123123123 | RND '' 6                      ## output = 1.231231 
    RND KB 2 234234234.234234234234                     ## output = 228744.37 KB
    234234234.234234234234 | RND KB 2                   ## output = 228744.37 KB
    234234234.234234234234 | RND MB 4                   ## output = 223.3832 MB
    234234234.234234234234 | RND GB 1                   ## output = 0.2 GB
    234234234.234234234234 | RND                        ## output = 0.22 GB
    234234234.234234234234 | RND MB 0                   ## output = 223 MB
    1223234234234.234234234234 | RND TB 2               ## output = 1.11 TB
    write-output "12312312313.123123123" | RND KB       ## output = 12023742.49 KB
    write-output "TEST $(1231231231.2312 | RND GB 2)"   ## output = TEST 1.15 GB
#>
} ## Rounding function for B,KB,MB,GB,TB

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
    
Function Send-GetPartnerInfo1 ($PartnerName) { 

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

Function Send-UpdateCustomColumn.old ($DeviceId,$ColumnId,$Message) {

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
    if ($script:updateCC.error.message) {}else{Write-output " Updating $DeviceId | Column AA$ColumnId | Value: $Message"}
}

Function Send-UpdateCustomColumn ($DeviceId, $ColumnId, $Message) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Cookie", "__cfduid=d110201d75658c43f9730368d03320d0f1601993342")
    $headers.Add("Authorization", "Bearer $script:visa")
    
    $body = @"
    {
        "jsonrpc": "2.0",
        "id": "jsonrpc",
        "method": "UpdateAccountCustomColumnValues",
        "params": {
            "accountId": $DeviceId,
            "values": [
                [$ColumnId, "$Message"]
            ]
        }
    }
"@
    
    $script:updateCC = Invoke-RestMethod 'https://cloudbackup.management/jsonapi' -Method 'POST' -Headers $headers -Body $body
    if ($script:updateCC.error.message) {
        # Handle error if needed
    } else {
        Write-Output "`nUpdating $DeviceId | Column AA$ColumnId | Value: $Message"
    }
}

#endregion ----- Backup.Management JSON Calls ----
    
#endregion ----- Functions ----
    
Function Get-FixedVolumes {

    $Volumes = Get-Volume

    $FixedVolumes = $Volumes | Where-Object {($_.DriveType -eq "Fixed") -and ($_.OperationalStatus -eq "OK") -and ($_.DriveLetter)}

    $FixedVolumes1 = $FixedVolumes | Select-Object @{l='Letter';e={$_.DriveLetter + ":"}},@{l='Type';e={$_.DriveType}},@{l='TotalGB';e={([Math]::Round(($_.Size/1GB),0))}},@{l='FreeGB';e={([Math]::Round(($_.SizeRemaining/1GB),0))}},@{l='UsedGB';e={([Math]::Round((($_.Size - $_.SizeRemaining)/1GB),0))}}
    
    $FixedVolumes2 = $FixedVolumes1 | Select-Object Letter,Type,UsedGB,@{l='String';e={$_.Letter + ' ' + $_.UsedGB +'GB'}}

    $FixedVolumeString = $FixedVolumes2.String -join " | "

    Send-UpdateCustomColumn $Script:Instance.AccountId 2646 "$((Get-Date).ToString("[yy-MM-dd HH:mm]")) $FixedVolumeString"
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
            $Script:ExcludeBase = "FileSystem "
        }else{
            $Script:ExcludeBase = $null
        }
        $Script:ExcludeString = $Script:Exclude.path.replace("\","\\") -join " | "
    }else{
        if (($Script:Include) -and ($Script:include[0].path -ne "")){
            $Script:ExcludeBase = "FileSystem "
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
       
    }else {
        Send-UpdateCustomColumn $Script:Instance.AccountId 2649 ""
        Write-output "  Filters = Not present"
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
    }else{
        Send-UpdateCustomColumn $Script:Instance.AccountId 2650 ""
        Write-output "  USB Volumes = Not present"
    }
}
      
Function Get-Schedules {    

    & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.schedule.list | out-file C:\programdata\mxb\schedules.csv
                
    [array]$schedules = import-csv -path C:\programdata\mxb\schedules.csv -Delimiter `t -Header id,actv,name,time,days,dsrc,presid,postsid
    
    if ($schedules -and $profileDetail) {
        Send-UpdateCustomColumn $Script:Instance.AccountId 3091 "Set by Profile"
        Write-output $Schedules
    
    }elseif ($schedules){

        $ActiveSchedules = $schedules | Where-Object {$_.actv -eq "yes"}
        $ActiveScheduleCount = $ActiveSchedules.Count
        $ActiveScheduleString = ($ActiveSchedules | Sort-Object time | ForEach-Object {
            "$($_.dsrc) $($_.time)"
         
        }) -join " | "

        $ActiveScheduleString = $ActiveScheduleString -replace("FileSystem","FS") -replace("SystemState","SS") -replace("Exchange","EX") -replace("SharePoint","SP") -replace("VssMsSql","MSSQL") -replace("Oracle","OR") -replace("NetworkShares","NS") -replace( "VMware","VM") -replace("Hyper-V","HV") -replace("MySQL","MySQL") 

        #($schedules | Where-Object {$_.actv -eq "yes"}) | Select-Object ACTV,DSRC,TIME | Sort-Object DSRC | Format-Table
      

        Send-UpdateCustomColumn $Script:Instance.AccountId 3091 "$ActiveScheduleCount | $ActiveSchedulestring"
        Write-output "  Schedules = $ActiveScheduleString"
    }else {
        Send-UpdateCustomColumn $Script:Instance.AccountId 3091 ""
        Write-output "  Active Schedules = Not present"
    }
}

Function Detect-Datasources {
    [array]$script:datasources = @()
        Write-output "`n[Scanning For Installed DataSources]"
    $CheckDataSources = @(
        @{Name = "Microsoft SQL Server"; Vendor = "*Microsoft*"; Product = "Microsoft SQL Server*"; Datasource = "MSSQL"},
        @{Name = "Microsoft Exchange"; Vendor = "*Microsoft*"; Product = "*Exchange*"; Datasource = "Exchange"},
        @{Name = "Microsoft Sharepoint"; Vendor = "*Microsoft*"; Product = "*Sharepoint*"; Datasource = "SharePoint"},
        @{Name = "MySQL"; Vendor = "*Oracle*"; Product = "MySQL*"; Datasource = "MySQL"},
        @{Name = "Oracle"; Vendor = "*Oracle*"; Product = "*Oracle*"; DataSource = "Oracle"}   
            ## Write-output "`n[Scanning For Microsoft Hyper-V DataSource] - TBD"
    )

    foreach ($dataSource in $CheckDataSources) {
        $detected = Get-WmiObject -Class Win32_Product | Where-Object {
            ($_.vendor -like $dataSource.Vendor) -and ($_.name -like $dataSource.Product)
        } | Select-Object Name, Vendor, Version -Unique | Sort-Object version

        if ($detected) {
            $detected[-1]
            $datasources += $dataSource.datasource
        }
    }

    if ($datasources) {    
        $script:datasourcesstring = ($datasources  | sort-object) -join " | "

        Send-UpdateCustomColumn $Script:Instance.AccountId 2906 "$datasourcesstring"
        Write-Output "`nDetected Data Sources       | $Script:datasourcesstring" 
    }else{
        Send-UpdateCustomColumn $Script:Instance.AccountId 2906 ""
        Write-Output "`nDetected Data Sources       | No Additional Data Sources Detected" 
    }
}

Function Get-Datasources {
    $StatusReportxml = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml"
    
    $Script:ActiveDataSources = (([Xml] (Get-Content $StatusReportxml)).SelectSingleNode("//ActivePlugins")."#text" -replace '(.)', '$1-' -split '(?<=-)' | ForEach-Object { 
        $_ -replace "F-", "FileSystem" `
           -replace "S-", "SystemState" `
           -replace "X-", "Exchange" `
           -replace "H-", "Hyper-V" `
           -replace "P-", "SharePoint" `
           -replace "Z-", "MSSQL" `
           -replace "L-", "MySQL" `
           -replace "N-", "NetworkShares" `
           -replace "W-", "VMware" `
           -replace "Y-", "Oracle" 
    }) | Where-Object { $_ -ne "" } | Sort-Object

    Write-Output "`nActive Data Sources         | $($Script:ActiveDataSources -join ' | ') "

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"
    $ConfiguredDatasources = & $clienttool -machine-readable control.selection.list | ConvertFrom-String | Select-Object -skip 1 -property p1,p2 -unique | ForEach-Object {If ($_.P2 -eq "Inclusive") {Write-Output $_.P1}}
    $ConfiguredDatasources = ($ConfiguredDatasources -replace("vss","") | Sort-Object | ForEach-Object { $_ }) -join " | "

    Write-Output "`nConfigured Data Sources     | $ConfiguredDatasources" 
} ## Get Active Data Sources from Status.xml and ClientTool.exe

### Main Script    

CDP-Authenticate

$wait = 30
$COUNTER=0
Do {
    $BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
    $BackupStatus = & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.status.get
    $status = @("Idle","Scanning","Backup")

    if (($BackupService.status -ne "Running") -or ($status -notcontains $BackupStatus)) { 
        Write-Warning "Backup Manager Not Ready"
        Write-Output "  Service $($BackupService.status)"
        Write-Output "  Job $backupstatus"
        $COUNTER ++
        Write-output "Attempt $Counter"
        Write-Output "  Retrying in $wait seconds"
        Start-Sleep -seconds $wait


        $BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
        $BackupStatus = & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.status.get

    }else{   
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) {
            Write-Warning "Backup Manager Not Running" 
        }else{ 
            try { 
                #$ErrorActionPreference = 'Stop'; 
                & "C:\Program Files\Backup Manager\ClientTool.exe" -machine-readable control.selection.list -delimiter "," | out-file C:\programdata\mxb\selections.csv
                
                $Script:Selections = import-csv -path C:\programdata\mxb\selections.csv

                [array]$Instances = get-childitem -Recurse -path c:\programdata\mxb *.info -file | Where-Object {$_.lastwritetime -gt (Get-Date).AddDays(-30)} | Select-Object Name,Directory,lastwritetime | Sort-Object lastwritetime -Descending  
    
                $InstanceInfo = join-path -Path $Instances[0].Directory -ChildPath $Instances[0].Name
    
                $Script:Instance = Get-Content -Path $InstanceInfo | Out-String | Convertfrom-json
    
                write-output "  ID $($Script:Instance.AccountId)"

                $profileInfo = $InstanceInfo.replace("info","profile") 

                if ((Test-Path $ProfileInfo) -and (Get-Content -Path $ProfileInfo | Out-String).Trim()) {
                    $currentProfile = Get-Content -Path $ProfileInfo | Out-String | Convertfrom-json
                    $profileDetail = $currentProfile.profileData.BackupDataSourceSettings | where {$_.datasource -like "*fileSystem"}
                    
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
                    Get-Schedules
                    Remove-Item C:\programdata\mxb\schedules.csv 
                    Detect-Datasources
                    Get-Datasources
                    [switch]$completed = $true
             
                }catch{ 
                        Write-Warning "Oops: $_" 
                    }
                }
            }
        }until (($completed) -or ($counter -ge 5))
    




#Phase 1 (file system available / selected, excluded. Filtered)
        #Done# Convert to Automation policy 
        #Done# Automation Policy Authentication
        #Done# Update Timestamp
        #Done# Clear selection values if null   
        #Done# Error checking
        #Done# local schedules
        #Done# schedule clear values 

        ## mixed environments with  no, profile, profile, keep local selections filters and exclusions
       
#Phase 2 (schedules)

        ## profile schedules
        #Partial# schedule Error checking


#Phase 3 (Other data sources)
        #Partial# Other Data sources 
        #Partial# Other Data sources / selections / filters
        #Partial# Other Data sources / selections / filters / Schedules

#Phase 4 (Misc metrics)
        ## Send data to a webhook (Partial code complete)
        ## Error Messages (code already written)
        ## Throttles (code already written)
        ## gui password (code already written)
        ## Health Check
        ## Backup Register size.





