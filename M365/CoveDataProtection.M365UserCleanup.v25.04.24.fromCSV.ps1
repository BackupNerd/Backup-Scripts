<# ----- About: ----
    # Cove Data Protection | M365 User Cleanup Offboarding From CSV
    # Revision v25.04.24
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com    
    # Reddit https://www.reddit.com/r/Nable/
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
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
    # For use with the Standalone edition of N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
  # -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Prompt for CSV file to load
    # Load CSV file to remove backup history from Selected Mailboxes
    # Deselect Mail and OneDrive selections from Mailbox
    # Remove Backup History for selected Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)
    # 
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
# -----------------------------------------------------------#>  ## Behavior

<# ----- Troubleshooting: ----
    # If you encounter issues running the script, ensure that the script is unblocked and the execution policy is set correctly.
    # You may need to login as an adminsitrator to perform these tasks.
    #
    # To unblock the script:
    # 1. Right-click the script file in File Explorer and select 'Properties'.
    # 2. In the 'General' tab, check the 'Unblock' checkbox if it is present.
    # 3. Click 'Apply' and then 'OK'.
    #
    # Alternatively, you can unblock the script using PowerShell:
    # 1. Open PowerShell as an administrator.
    # 2. Run the following command to unblock the script:
    #    Unblock-File -Path "C:\Path\To\Your\Script.ps1"
    #
    # To set the execution policy to allow scripts to run:
    # 1. Open PowerShell as an administrator.
    # 2. Run the following command to set the execution policy:
    #    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
    # 3. If prompted, confirm the change by typing 'Y' and pressing Enter.
    #
    # Note: Setting the execution policy to 'Unrestricted' allows all scripts to run, which is less secure but can be useful for troubleshooting.
    #       Alternatively, you can use 'Bypass' to completely bypass the execution policy for the current session:
    #       Run the following command to set the execution policy to 'Bypass':
    #       Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
    #       This setting is temporary and only applies to the current PowerShell session.
# -----------------------------------------------------------#>  ## Troubleshooting

[CmdletBinding()]
    Param (

        [Parameter(Mandatory=$False)] [switch]$DeleteFromCSV = $true,                           ## Load CSV file to remove backup history from Selected Mailboxes
        [Parameter(Mandatory=$False)] [string]$delimiter = ",",                                 ## Delimiter
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                 ## Remove Stored API Credentials at start of script
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    #Requires -Version 5.1
    $scriptStartTime = Get-Date  # Set the start time at the beginning of the script
    $ConsoleTitle = "Cove Data Protection | M365 User Cleanup Offboarding From CSV"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle

    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax 
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    $CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
    $ShortDate = Get-Date -format "yyyy-MM-dd"

    if ($ExportPath) {$ExportPath = Join-path -path $ExportPath -childpath "M365_$shortdate"}else{$ExportPath = Join-path -path $dir -childpath "M365_$shortdate"}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    If ($exportcombined) {mkdir -force -path $ExportPath | Out-Null}
    $urlJSON = 'https://api.backup.management/jsonapi'

    $culture = get-culture; $delimiter = $culture.TextInfo.ListSeparator

    Write-output "  Current Parameters:"
    Write-output "  -Delimiter                      = $delimiter"
    
#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Set-APICredentials {

    Write-Output $Script:strLineSeparator 
    Write-Output "  Setting Backup API Credentials" 
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential Path Present" }else{ New-Item -ItemType Directory -Path $APIcredpath} 
        Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove | Backup.Management API i.e. 'Acme, Inc (bob@acme.net)'"
    DO{ $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.length -eq 0)
    $PartnerName | out-file $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove | Backup.Management API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName" 

    $BackupCred.UserName | Out-file -append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-file -append $APIcredfile
    
    Start-Sleep -milliseconds 300

    Send-APICredentialsCookie  ## Attempt API Authentication

}  ## Set API credentials if not present

Function Get-APICredentials {

    $Script:True_path = "C:\ProgramData\MXB\"
    if (-not (Test-Path -Path $Script:True_path)) {
        New-Item -ItemType Directory -Path $Script:True_path -Force
    }
    $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-Path -Path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) { 
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator 
        Write-Output "  Backup API Credential File Cleared"
        Send-APICredentialsCookie  ## Retry Authentication
        
        }else{ 
            Write-Output $Script:strLineSeparator 
            Write-Output "  Getting Backup API Credentials" 
        
            if (Test-Path $APIcredfile) {
                Write-Output    $Script:strLineSeparator        
                "  Backup API Credential File Present"
                $APIcredentials = get-content $APIcredfile
                
                $Script:cred0 = [string]$APIcredentials[0] 
                $Script:cred1 = [string]$APIcredentials[1]
                $Script:cred2 = $APIcredentials[2] | Convertto-SecureString 
                $Script:cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))

                Write-Output    $Script:strLineSeparator 
                Write-output "  Stored Backup API Partner  = $Script:cred0"
                Write-output "  Stored Backup API User     = $Script:cred1"
                Write-output "  Stored Backup API Password = Encrypted`n"
                
            }else{
                Write-Output    $Script:strLineSeparator 
                Write-Output "  Backup API Credential File Not Present"

                Set-APICredentials  ## Create API Credential File if Not Found
                }
            }

}  ## Get API credentials if present

Function Send-APICredentialsCookie {

    Get-APICredentials  ## Read API Credential File before Authentication

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    #$data.params.partner = $Script:cred0
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
        $script:UserId = $authenticate.result.result.id

        if (($authenticate.result.result.flags -notcontains "securityofficer") -and 
            ($CleanupSharedBillable -or $CleanupUnlicensed -or $CleanupDeletedBillable -or $CleanupDeletedShared -or $CleanupDeletedBillableShared)
            ) { Write-warning "Security Office Role not found, Deletion of Historic Backup data not allowed!"}
            
    }else{
        Write-Output    $Script:strLineSeparator 
        Write-Warning "`n  $($authenticate.error.message)"
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"

        Write-Output    $Script:strLineSeparator 
        
        Set-APICredentials  ## Create API Credential File if Authentication Fails
    }

}  ## Use Backup.Management credentials to Authenticate

Function Visa-Check {
     if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){   
            Send-APICredentialsCookie
        }
    }
}  ## Recheck remaining Visa time and reauthenticate

#endregion ----- Authentication ----

#region ----- Data Conversion ----

function Show-ElapsedTime {
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $scriptStartTime
    Write-Output "Elapsed time: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s"
}  ## Function to display the elapsed time

Function Convert-UnixTimeToDateTime($UnixToConvert) {
    if ($UnixToConvert -gt 0 ) { $Epoch2Date = ((Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()).AddSeconds($UnixToConvert)
    return $Epoch2Date }else{ return ""}
}  ## Convert epochtime to datetime #Rev.03

Function Convert-DateTimeToUnixTime($DateToConvert) {
    $Date2Epoch = (New-TimeSpan -Start (Get-Date -Date "1970-01-01 00:00:00Z")-End (Get-Date -Date $DateToConvert)).TotalSeconds
    Return $Date2Epoch
}  ## Convert datetime to epochtime #Rev.03

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----

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
        #$Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
        $Script:Partner = $webrequest | convertfrom-json

    $RestrictedPartnerLevel = @("Root","SubRoot","Distributor")
    
    if ($Partner.result.result.Level -notin $RestrictedPartnerLevel) {
        [String]$Script:Uid = $Partner.result.result.Uid
        [int]$Script:PartnerId = [int]$Partner.result.result.Id
        [String]$script:Level = $Partner.result.result.Level
        [String]$Script:PartnerName = $Partner.result.result.Name

        Write-Output $Script:strLineSeparator
        Write-output "  $PartnerName - $partnerId - $Uid"
        Write-Output $Script:strLineSeparator
        }else{
        Write-Output $Script:strLineSeparator
        Write-Host "  Lookup for $($Partner.result.result.Level) Partner Level Not Allowed"
        Write-Output $Script:strLineSeparator
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername
        }

    if ($partner.error) {
        write-warning "  $($partner.error.message)"
        $Script:PartnerName = Read-Host "  Enter EXACT Case Sensitive Customer/ Partner displayed name to lookup i.e. 'Acme, Inc (bob@acme.net)'"
        Send-GetPartnerInfo $Script:partnername

    }else{$script:visa = $Partner.visa}

} ## get PartnerID and Partner Level    

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

Function Open-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Comma Seperated Value (*.csv)|*.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
} ## GUI Prompt for Filename to open

Function UpdateDatasources ($AccountToken,$EntityId,$MailBoxSelection,$ExchangeAutoInclude,$OneDriveSelection,$OneDriveAutoInclude) {

    $url = "https://api.backup.management/management_api"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'UpdateDataSources'
    $data.params = @{}
    $data.params.accountToken = $accountToken
    $data.params.dataSources = @{}
    $data.params.dataSources.DataSourceEntityAutoInclusions = @(
        [pscustomobject]@{DataSourceType="Exchange";EntityAutoInclusionType="$ExchangeAutoInclude"},
        [pscustomobject]@{DataSourceType="OneDrive";EntityAutoInclusionType="$OneDriveAutoInclude"}
        )

    $data.params.dataSources.Selections = @(
        [pscustomobject]@{EntityId="$entityId";DataSourceSelections = @(
            [pscustomobject]@{DataSourceType="Exchange";SelectionType="$MailBoxSelection"},
            [pscustomobject]@{DataSourceType="OneDrive";SelectionType="$OneDriveSelection"}
            )}
        )

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        WebSession  = $websession
        ContentType = 'application/json; charset=utf-8'
        TimeoutSec  = 300
    }   

    $Script:UpdateDataSource = Invoke-RestMethod @params 
    if ($Script:UpdateDataSource.error) {Write-output $Script:UpdateDataSource.error.message $entityId }else{Write-output "Updating UserGuid $entityId"} 
}

Function Remove-BackupHistory {
    param (
        [string]$DataSourceType,
        [string]$AccountToken,
        [string]$UserGuid
    )

    $url = "https://api.backup.management/management_api"
    $method = 'POST'
    $data = @{
        jsonrpc = '2.0'
        id      = '2'
        method  = 'RemoveBackup'
        params  = @{
            accountToken   = $AccountToken
            dataSourceType = $DataSourceType
            entityId       = $UserGuid
        }
    }

    $jsondata = $data | ConvertTo-Json -Depth 6

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = $jsondata
        ContentType = 'application/json; charset=utf-8'
    }

    $Script:RemovedBackupHistory = Invoke-RestMethod @params

    if ($Script:RemovedBackupHistory.error) {
        Write-Warning "$($Script:RemovedBackupHistory.error.message)"
        if ($authenticate.result.result.Flags -notcontains "SecurityOfficer") {
            Write-Warning "API User does not have Security Officer rights"
        }
    }
}

Send-APICredentialsCookie


if ($DeleteFromCSV) {

    $openfilename = Open-Filename
    $Deleted = Import-Csv -path $OpenFileName -Delimiter $delimiter

    $Script:ToDelete = $Deleted | Out-GridView -Title "Select Accounts to Remove from Backup and Delete Backup History" -OutputMode Multiple


    foreach ($item in $ToDelete) {
        visa-check
        Write-output "Attempting to deselect Mail and OneDrive selections from Mailbox"

        UpdateDataSources $item.AccountToken $item.UserGuid "Excluded" $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Exclude Mailbox selection & Onedrive for user

        Write-output "`nAttempting to remove Backup History for selected Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
        Remove-BackupHistory "Exchange" $item.accounttoken $item.UserGuid
        Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
    }

}


#endregion ----- Export Deleted ----
