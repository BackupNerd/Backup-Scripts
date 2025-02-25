<# ----- About: ----
    # Bulk Delete Orphaned, Inactive & Unused Devices
    # Revision v.04 - 2025-02-23
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
    # The script requires PowerShell 5.1 or higher and should be run with administrative privileges.
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # This script is designed to bulk delete Orphaned, Inactive & Unused devices from the N-able | Cove Data Protection platform.
    # The script is intended for use by N-able | Cove Data Protection partners and customers who need to manage their devices effectively.
    # It uses the N-able | Cove Data Protection API to authenticate and enumerate devices.
    # The script allows for filtering of devices based on various criteria such as inactivity, empty devices, and storage size.
    # The script can be run with or without user input, and it can be configured to force deletion without confirmation.
    # The script logs the devices deleted to a CSV file for reference.

    # The script is designed to be run in a Windows environment and requires the N-able | Cove Data Protection API credentials for authentication.
    # Care should be taken when running this script, as it will permanently delete devices from the N-able | Cove Data Protection platform.
    #
    # Warning: confirm script settings before running in a production environment.
    #
    # Accidental deletion may be able to be undone by N-able | Cove Data Protection technical support if reported within 28 days.
    #
    # Authenticate to https://backup.management console
    # Enumerate devices that match the filter
    # Select Orphaned, Inactive & Unused devices that match the filter for deletion
    # Use the -CoveLoginUserName to specify the Cove User Name for API authentication
    # Use the -CovePlainTextPassword to specify the Cove Password for API authentication
    # Use the -TaskPath to specify the path to store script credentials and logs
    # Use the -ExcludeArchive switch parameter to exclude devices with archives
    # Use the -ExcludeUndefined switch parameter to exclude undefined devices
    # Use the -ExcludeWorkstations switch parameter to exclude workstations
    # Use the -ExcludeServers switch parameter to exclude servers
    # Use the -DeviceCount ## parameter to limit the maximum device count to return
    # Use the -InactiveTimeStamp to specify the time stamp for inactive devices
    # Use the -UnusedTimeStamp to specify the time stamp for unused devices
    # Use the -SelectedSizeLessThanBytes to specify the size of selected devices
    # Use the -UsedStorageLessThanBytes to specify the used storage size for devices
    # Use the -RestrictedPartnerLevel to specify the partner level to restrict access to
    # Use the -ForceDelete switch parameter to force deletion without additional confirmation
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/API-column-codes.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/enumerate-device-statistics.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/advanced-filter-expressions.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/console-new/advanced-filter-syntax.htm
 
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)][String]$CoveLoginUserName,                                ##  Remove / replace with prompt for AMP Usage   
        [Parameter(Mandatory=$False)][SecureString]$CovePlainTextPassword,                      ##  Remove / replace with prompt for AMP Usage
        [Parameter(Mandatory=$False)][string]$TaskPath = "c:\ProgramData\BackupNerdScripts",    ##  Base Path to Store Script Credentials and Logs
        [Parameter(Mandatory=$False)][int]$DeviceCount = 10,                                    ##  Maximum Device Count to Return
        [Parameter(Mandatory=$False)][switch]$ExcludeArchive =$true,                            ##  set switch to $true to exclude devices with archives from deletion
        [Parameter(Mandatory=$False)][switch]$ExcludeUndefined = $false,                        ##  set switch to $true to exclude undefined devices from deletion
        [Parameter(Mandatory=$False)][switch]$ExcludeDocuments = $true,                         ##  set switch to $true to exclude documents devices from deletion      
        [Parameter(Mandatory=$False)][switch]$ExcludeWorkstations = $true,                      ##  set switch to $true to exclude workstations from deletion
        [Parameter(Mandatory=$False)][switch]$ExcludeServers = $true,                           ##  set switch to $true to exclude servers from deletion
        [Parameter(Mandatory=$False)][int]$InactiveTimeStamp = 90,                              ##  Set TimeStamp age for Inactive Devices in Days to Remove regarless of Storage size
        [Parameter(Mandatory=$False)][int]$UnusedTimeStamp = 30,                                ##  Set TimeStamp age for Unused Devices in Days to Remove if also under X size
        [Parameter(Mandatory=$False)][int]$SelectedSizeLessThanBytes = 0,                       ##  Empty Device with Less Than X Selected Size in Bytes
        [Parameter(Mandatory=$False)][int]$UsedStorageLessThanBytes = 0,                        ##  Empty Device with Less Than X Used Storage in Bytes
        [Parameter(Mandatory=$False)][Switch]$ForceDelete = $false
        
    )

if ( $ExcludeUndefined -and $ExcludeServers -and $ExcludeWorkstations -and $ExcludeDocuments ) {
    write-warning "All Backup Manager device types (Server, Workstation, Documents, Undedefined) have been excluded.  Please check your parameters and do not exclude all device types."
    Break
}

if ($excludeArchive) { $ARex= "(AS == 0 ) AND " }
if ($excludeUndefined) { $UNex= " AND (OT != 0)" }
if ($excludeWorkstations) { $WSex= " AND (OT != 1)" }

if ($excludeWorkstations -and $ExcludeDocuments ) { 
    $WSex= " AND (OT != 1)" 
    $DOex= " AND (OP != 'Documents')"
}
elseif ($excludeDocuments -and ($false -eq $ExcludeWorkstations)) {
    #$WSex= " AND (OT == 1)"
    $DOex= " AND (OP != 'Documents')" 
}
elseif (($false -eq $excludeDocuments) -and $ExcludeWorkstations) {
    $WSex= $null
    $DOex= " AND (OP == 'Documents')"
}

if ($excludeServers) { $SVex= " AND (OT != 2)" }

$DeviceFilter = "$($arex)( (AT == 1)$($unex)$($wsex)$($svex)$($DOex) ) AND ( ( TS <= $InactiveTimeStamp.days().ago() ) OR ( TS <= $UnusedTimeStamp.days().ago() AND US <= $UsedStorageLessThanBytes AND T3 <= $SelectedSizeLessThanBytes ) )"

#region ----- Environment, Variables, Names and Paths ----
Clear-Host
$ConsoleTitle = "Bulk Delete Orphaned, Inactive & Unused Devices"

## - Remove Credential to be prompted or passed from Script or AMP
#[string]$CoveLoginUserName = $null
#[string]$CovePlainTextPassword = $null
## - Remove Credential to be prompted or passed from Script or AMP

#[int]$DeviceTimeStamp = 60
#[int]$DeviceCreation = 90
#[int]$DeviceCount = 50
#[string]$TaskPath = "c:\ProgramData\BackupNerdScripts"
#$DeviceFilter = "US == 0 AND T3 == 0"
#$DeviceFilter = "OI == 5038 AND OP == 'Documents' AND PD == 21266 AND PN == 'Documents' AND TS < $DeviceTimeStamp.days().ago() AND CD < $DeviceCreation.days().ago()"
#$DeviceFilter = "( OT != 2 ) AND ( ( OT == 0 ) OR ( TS < 90.days().ago() ) OR ( TS < 30.days().ago() AND US == 0 AND T3 == 0 ) )"
#$RestrictedPartnerLevel = @("Root","SubRoot","Distributor","Subdistributor")
        
# Note: < > signs in device filters may appear reversed due to counting backwards Epoch time backwards


## - Remove or comment out for use in AMP
#Requires -Version 5.1 -RunAsAdministrator
$host.UI.RawUI.WindowTitle = $ConsoleTitle
## - Remove or comment out for use in AMP

$ScriptPath = $PSCommandPath
$ScriptBase = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
$ScriptLog = Join-Path -Path $TaskPath -ChildPath "$ScriptBase.log.csv"
$ScriptLogParent = Split-Path -Path $ScriptLog
if (-not (Test-Path $ScriptLogParent)) { mkdir -Force $ScriptLogParent }

Write-Output "$ConsoleTitle`n`n$ScriptPath"
Write-Output "Script Parameter Syntax:`n`n$(Get-Command $PSCommandPath -Syntax)"

$CurrentDate = Get-Date -format "yyy-MM-dd_HH-mm-ss"
$ShortDate = Get-Date -format "yyy-MM-dd"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Script:strLineSeparator = "  ---------"

Write-output "  Current Parameters:"
Write-output "  -TaskPath               = $TaskPath"
Write-output "  -LogPath                = $ScriptLog"
Write-output "  -DeviceCount            = $DeviceCount"
Write-output "  -DeviceFilter           = $DeviceFilter"
Write-Warning "Paste this Advanced Filter in Your N-able | Cove Data Protection Console to see what devices would be selected for deletion:"
$Script:strLineSeparator 
Write-output "  -ExcludeUndefined       = $ExcludeUndefined"
Write-output "  -ExcludeWorkstations    = $ExcludeWorkstations"
Write-output "  -ExcludeDocuments       = $ExcludeDocuments"
Write-output "  -ExcludeServers         = $ExcludeServers"
Write-output "  -ExcludeArchives        = $ExcludeArchive"
$Script:strLineSeparator 
Write-output "  -InactiveTimeStamp      = $InactiveTimeStamp Days"
$Script:strLineSeparator 
Write-output "  -UnusedTimeStamp        = $UnusedTimeStamp Days"
Write-output "  -SelectedSize(T3)      <= $SelectedSizeLessThanBytes bytes"
Write-output "  -UsedStorage(US)       <= $UsedStorageLessThanBytes bytes"
$Script:strLineSeparator 
write-output "  -ForceDelete            = $ForceDelete"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function Authenticate {

    Write-Output $script:strLineSeparator
    Write-Output "  Enter Your N-able | Cove https:\\backup.management API User Credentials"
    Write-Output $script:strLineSeparator

    if ([string]::IsNullOrEmpty($CoveLoginUserName)) {
        $CoveLoginUserName = Read-Host -Prompt "  Enter Login UserName or Email for N-able | Cove Backup.Management API"
    }

    if ([string]::IsNullOrEmpty($CovePlainTextPassword)) {
        $CovePassword = Read-Host -AsSecureString "  Enter Password for N-able | Cove Backup.Management API"
        # (Convert SecureString Password to plain text)
        $CovePlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($CovePassword))
    }

    # (Show credentials for Debugging)
    Write-Output "  Logging on with the following Credentials`n"
    Write-Output "  UserName:     $CoveLoginUserName"
    Write-Output "  Password:     It's secure..."

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
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
        If ((Get-Date).ToUniversalTime() -gt $VisaTime.AddMinutes(10)) {
            Authenticate
        } else {
            if ($Debug) {
                Write-Output "Visa UTC      :$($VisaTime)"
                Write-Output "Current UTC   :$((Get-Date).ToUniversalTime())"
                Write-Output "Visa Refesh @ :$($VisaTime.AddMinutes(10))"
                $TimeDifference = $VisaTime.AddMinutes(10) - (Get-Date).ToUniversalTime()
                Write-Output "Visa Refresh  :$TimeDifference"
                Write-Output "Visa is still valid, no need to re-authenticate"
            }
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

Function Send-GetPartnerInfobyid ($partnerid) { 

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:visa
    $data.method = 'GetPartnerInfoById'
    $data.params = @{}
    $data.params.partnerId = $PartnerId

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

    $Script:RestrictedPartnerLevel = @("Root","SubRoot","Distributor","Subdistributor")

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

Function Send-GetPartnerInfo ($PartnerName) { 
        
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
        [int]$Script:PartnerId = $Partner.result.result.Id
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
    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $script:visa
    $data.method = 'EnumerateAccountStatistics'
    $data.params = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId = $script:PartnerId
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

    $Script:InactiveDevices = Invoke-RestMethod @params 

    $Script:DeviceDetail = @()

    if ($InactiveDevices.result.result.count -eq 0) {
        Write-Output "No orphaned, inactive or obsolete devices found that meet Filter criteria.`n  Exiting Script"
        Start-Sleep -seconds 10
        Break
    }
    
    Write-Output "  Requesting details for $($InactiveDevices.result.result.count) Orphaned, Inactive & Unused devices."
    Write-Output "  Please be patient, this could take some time."
    Write-Output $Script:strLineSeparator
    ForEach ( $DeviceResult in $InactiveDevices.result.result ) {
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
Send-GetPartnerInfobyid $script:PartnerId
Send-EnumerateDevices

if ($FORCEDELETE) {
    $SelectedDevices = $Script:DeviceDetail

    if ($null -eq $SelectedDevices) {
        Write-Warning "No selections made"
        Break
    }

}ELSE{
    $SelectedDevices = $Script:DeviceDetail | Select-Object LogTime,Action,AccountID,Creation,LastSuccess,Timestamp,PartnerName,Reference,DeviceName,ComputerName,DeviceAlias,Location,DataSources,TempInfo,SelectedGB,UsedGB,Profile,ProfileID,Product,ProductID,Errors,Filter| Out-GridView -Title "$Script:PartnerName | Displaying $($InactiveDevices.result.result.count) | Orphaned, Inactive & Unused Device matching Device Filter | Select Devices to Purge" -OutputMode Multiple
    
    if ($null -eq $SelectedDevices) {
        Write-Warning "No selections made"
        Break
    }

}

if ($SelectedDevices) {
    get-visaTime

    $SelectedDevice | Select-Object * | Export-Csv -Path $ScriptLog -Append

    Foreach ($Deviceidtodelete in $selectedDevices) {
        Send-RemoveAccount $Deviceidtodelete.accountId
    }
    Write-Warning "In case of accidental device deletion it may be possible for Cove technical support to undelete devices within 28 days."
}

