<# ----- About: ----
    # Cove Data Protection | M365 User Cleanup Offboarding 
    # Revision v24.11.01 - 2026-01-23
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
    # Check partner level/ Enumerate partners/ GUI select partner
    # Enumerate devices/ GUI select M365 devices
    # Export selected Domains and Users to XLS/CSV
    # Optionally identify Shared Billable mailboxes and export to CSV
    # Optionally select Shared Billable mailboxes for deselection and removal
    # Optionally identify Deleted Shared / Deleted Billable mailboxes and export to CSV
    # Optionally select Deleted Shared / Deleted Billable mailboxes for deselection and removal
    #
    # Use the -AllPartners switch parameter to skip GUI partner selection
    # Use the -AllDevices switch parameter to skip GUI device selection
    # Use the -DeviceCount ## (default=10000) parameter to define the maximum number of domains/devices returned
    #
    # Use the -OneDriveAge ## days (Default=30) parameter to define the age in days of shared billable users to identify
    # Set the -ExportSharedBillable switch parameter to $true to identify shared billable users due to OneDrive history
    # Set the -CleanupSharedBillable switch parameter to $true to prompt for removal of shared billable users due to OneDrive history
    #
    # Use the UnlicensedAge ## days (Default=30) parameter to define the age in days of unlicensed users to identify
    # Set the -ExportUnlicensed Switch parameter to $true to identify unlicensed users
    # Set the -CleanupUnlicenses switch parameters to $true to prompt for removal of unlicensed users
    #
    # Use the -DeletedAge ## days (Default=30) parameter to define the age in days of deleted users to identify
    # Set the -ExportDeleted switch parameter to $true to identify deleted users
    # Set the -CleanupDeleted<type> switch parameters to $true to prompt for removal of deleted users
    #
    # Set the -ExportCombined switch parameter to $true (REQUIRED) to export combined M365 device and user statistics to XLS/CSV files
    # Use the -ExportPath (?:\Folder) parameter to specify XLS/CSV file path
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm

# -----------------------------------------------------------#>  ## Behavior

<# ----- Example Execution: ----
    # This section provides examples of how to execute the script from a command prompt or PowerShell prompt with various parameters.
    #
    # Example 1: Run the script with default parameters
    # Command Prompt (cmd.exe):
    #    powershell -File .\CoveDataProtection.M365UserCleanup.v24.11.01a.ps1
    # PowerShell: 
    #    .\CoveDataProtection.M365UserCleanup.v24.11.01a.ps1
    #
    # Example 2: Run the script with both an export path and skip GUI partner / device selection
    # PowerShell:
    #    .\CoveDataProtection.M365UserCleanup.v24.11.01a.ps1 -ExportPath "C:\Exports" -AllPartners -AllDevices
    #
    # Example 3: Run the script and cleanup unlicensed users older than 60 days
    # PowerShell:
    #    .\CoveDataProtection.M365UserCleanup.v24.11.01a.ps1 -CleanupUnlicensed -UnlicensedAge 60
    #
    # Example 4: Run the script and cleanup unlicensed users older than 60 days and shared billable users older than 90 days    
    # PowerShell:
    #    .\CoveDataProtection.M365UserCleanup.v24.11.01a.ps1 -CleanupUnlicensed -UnlicensedAge 60 -CleanupSharedBillable -OneDriveAge 90
    #
    # Note: Ensure that the script is unblocked and the execution policy is set correctly before running the script.
# -----------------------------------------------------------#>  ## Example Execution

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
        [Parameter(Mandatory=$False)] [string]$PartnerName = "",                                  ## Exact, Case Sensitive Partner Name to skip prompt
        [Parameter(Mandatory=$False)] [bool]$AllPartners=$true,                              ## $AllPartners = $true, to Skip GUI partner selection
        [Parameter(Mandatory=$False)] [bool]$AllDevices=$true,                               ## $AllDevices = $true, to Skip GUI device selection
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 10000,                                ## Set maximum number of device / domain results to return

        [Parameter(Mandatory=$False)] [bool]$ExportCombined = $true,                          ## Generate combined XLS/CSV output files for M365 devices and users
        [Parameter(Mandatory=$False)] [bool]$Launch = $false,                                          ## Launch combined XLS/CSV outputfile if generated

        [Parameter(Mandatory=$False)] [bool]$ExportSharedBillable = $true,                    ## Export Billable Shared Mailboxes with Onedrive History list
        [Parameter(Mandatory=$False)] [bool]$CleanupSharedBillable = $false,                   ## Prompt to Cleanup Billable Shared Mailboxes with Onedrive History
        [Parameter(Mandatory=$False)] [int]$OneDriveAge = 45,                                   ## Ignore Billable Shared Mailboxes with OneDrive Backups in the last X Days 

        [Parameter(Mandatory=$False)] [bool]$ExportUnlicensed = $true,                        ## Export Unlicensed Mailboxes list
        [Parameter(Mandatory=$False)] [bool]$CleanupUnlicensed = $false,                       ## Prompt to Cleanup Deleted Mailboxes
        [Parameter(Mandatory=$False)] [int]$UnlicensedAge = 45,                                 ## Ignore Deleted Mailboxes where in the last X Days

        [Parameter(Mandatory=$False)] [bool]$ExportDeleted = $true,                           ## Export Deleted Mailboxes list
        [Parameter(Mandatory=$False)] [bool]$CleanupDeletedBillable = $false,                  ## Prompt to Cleanup Deleted Mailboxes
        [Parameter(Mandatory=$False)] [bool]$CleanupDeletedShared = $false,                    ## Prompt to Cleanup Deleted Mailboxes
        [Parameter(Mandatory=$False)] [bool]$CleanupDeletedBillableShared = $false,            ## Prompt to Cleanup Deleted Mailboxes
        [Parameter(Mandatory=$False)] [int]$DeletedAge = 45,                                    ## Ignore Deleted Mailboxes where in the last X Days

        [Parameter(Mandatory=$False)] [bool]$ForceClean = $false,                              ## if ($true) Skip GridView selection and proceed with cleanup automatically
        [Parameter(Mandatory=$False)] [bool]$CDPDebug = $false,                               ## Show detailed output tables for debugging

        [Parameter(Mandatory=$False)] [string]$ExportPath = "$PSScriptRoot",                    ## Export base path, dated sub folders will be created
        [Parameter(Mandatory=$False)] [string]$delimiter = ",",                                 ## Delimiter
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials                                 ## Remove Stored API Credentials at start of script
    )

    # Do Not Delete these emails or domains
    $DoNotDeleteEmails = @(
        "user1@example.com",
        "user2@example.com",
        "user3@example.com",
        "user4@example.com",
        "user5@example.com"
    )

    $DoNotDeleteDomains = @(
        "example1.com",
        "example2.org",
        "example3.net",
        "example4.edu",
        "example5.io"
    )


#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    #Requires -Version 5.1
    $scriptStartTime = Get-Date  # Set the start time at the beginning of the script
    $ConsoleTitle = "Cove Data Protection | M365 User Cleanup Offboarding "
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

    # Initialize execution log file path (header created after partner info is retrieved)
    $Script:executionLogFile = "$ExportPath\$($CurrentDate)_M365_Execution_Log_$($Partnername -replace ' \(.*\)','' -replace '[^a-zA-Z_0-9]','')_$($PartnerId).log"

    $culture = get-culture; $delimiter = $culture.TextInfo.ListSeparator

    Write-output "  Current Parameters:"
    Write-output "  -AllPartners                    = $AllPartners"
    Write-output "  -AllDevices                     = $AllDevices"
    Write-output "  -DeviceCount                    = $DeviceCount"
    Write-output "  -ExportCombined                 = $ExportCombined"
    Write-output "  -ExportShared                   = $ExportSharedBillable"
    Write-output "  -CleanupShared                  = $CleanupSharedBillable"
    Write-output "  -SharedAge                      = $OneDriveAge"
    Write-output "  -ExportUnlicensed               = $ExportUnlicensed"
    Write-output "  -CleanupUnlicensed              = $CleanupUnlicensed"
    Write-output "  -UnlicensedAge                  = $UnlicensedAge"
    Write-output "  -ExportDeleted                  = $ExportDeleted"
    Write-output "  -CleanupDeletedBillable         = $CleanupDeletedBillable"
    Write-output "  -CleanupDeletedShared           = $CleanupDeletedShared"
    Write-output "  -CleanupDeletedBillableShared   = $CleanupDeletedBillableShared"
    Write-output "  -DeletedAge                     = $DeletedAge"
    Write-output "  -ExportPath                     = $ExportPath"
    Write-output "  -Delimiter                      = $delimiter"
    
#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
Function Write-ExecutionLog {
    param(
        [string]$Operation,
        [string]$Email,
        [string]$UserGuid,
        [string]$Result,
        [string]$Details = "",
        [switch]$WhatIf
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $whatIfPrefix = if ($WhatIf) { "WHATIF - " } else { "" }
    $logEntry = "$timestamp | $whatIfPrefix$Operation | $Email | $UserGuid | $Result | $Details"
    
    # Append to log file
    Add-Content -Path $Script:executionLogFile -Value $logEntry -ErrorAction SilentlyContinue
}

Function Test-ProtectedEmail {
    param(
        [string]$EmailAddress
    )
    
    # Check if email matches any protected email patterns
    foreach ($pattern in $DoNotDeleteEmails) {
        if ($EmailAddress -like "*$pattern*") {
            return $true
        }
    }
    
    # Check if email domain matches any protected domains
    foreach ($domain in $DoNotDeleteDomains) {
        if ($EmailAddress -like "*@$domain") {
            return $true
        }
    }
    
    return $false
}

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

        <#
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }
        #>  ## Increment file name

        $xl.DisplayAlerts = $False  ## Block overwrite file warning
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

Function Send-EnumeratePartners {
    # ----- Get Partners via EnumeratePartners -----
    
    # (Create the JSON object to call the EnumeratePartners function)
        $objEnumeratePartners = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
            Add-Member -PassThru NoteProperty visa $Script:visa |
            Add-Member -PassThru NoteProperty method 'EnumeratePartners' |
            Add-Member -PassThru NoteProperty params @{
                                                        parentPartnerId = $PartnerId 
                                                        fetchRecursively = "true"
                                                        fields = (0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22) 
                                                        } |
            Add-Member -PassThru NoteProperty id '1')| ConvertTo-Json -Depth 5
    
    # (Call the JSON Web Request Function to get the EnumeratePartners Object)
            [array]$Script:EnumeratePartnersSession = CallJSON $urlJSON $objEnumeratePartners
    
            $Script:visa = $EnumeratePartnersSession.visa
    
            #Write-Output    $Script:strLineSeparator
            #Write-Output    "  Using Visa:" $Script:visa
            #Write-Output    $Script:strLineSeparator
    
    # (Added Delay in case command takes a bit to respond)
            Start-Sleep -Milliseconds 100
    
    # (Get Result Status of EnumerateAccountProfiles)
            $EnumeratePartnersSessionErrorCode = $EnumeratePartnersSession.error.code
            $EnumeratePartnersSessionErrorMsg = $EnumeratePartnersSession.error.message
    
    # (Check for Errors with EnumeratePartners - Check if ErrorCode has a value)
            if ($EnumeratePartnersSessionErrorCode) {
                Write-Output    $Script:strLineSeparator
                Write-Output    "  EnumeratePartnersSession Error Code:  $EnumeratePartnersSessionErrorCode"
                Write-Output    "  EnumeratePartnersSession Message:  $EnumeratePartnersSessionErrorMsg"
                Write-Output    $Script:strLineSeparator
                Write-Output    "  Exiting Script"
    # (Exit Script if there is a problem)
    
                #Break Script
            }else{
    # (No error)
    
            $Script:EnumeratePartnersSessionResults = $EnumeratePartnersSession.result.result | select-object Name,@{l='Id';e={($_.Id).tostring()}},Level,ExternalCode,ParentId,LocationId,* -ExcludeProperty Company -ErrorAction Ignore
            
            $Script:EnumeratePartnersSessionResults | ForEach-Object {$_.CreationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.CreationTime))}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialExpirationTime  -ne "0") { $_.TrialExpirationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialExpirationTime))}}
            $Script:EnumeratePartnersSessionResults | ForEach-Object { if ($_.TrialRegistrationTime -ne "0") {$_.TrialRegistrationTime = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.TrialRegistrationTime))}}
        
            $Script:SelectedPartners = $EnumeratePartnersSessionResults | Select-object * | Where-object {$_.name -notlike "001???????????????- Recycle Bin"} | Where-object {$_.Externalcode -notlike '`[??????????`]* - ????????-????-????-????-????????????'}
                    
            $Script:SelectedPartner = $Script:SelectedPartners += @( [pscustomobject]@{Name=$PartnerName;Id=[string]$PartnerId;Level='<ParentPartner>'} ) 
                            
            if ($AllPartners) {
                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name
                Write-Output    $Script:strLineSeparator
                Write-Output    "  All Partners Selected"
            }else{
                $script:Selection = $Script:SelectedPartners |Select-object id,Name,Level,CreationTime,State,TrialRegistrationTime,TrialExpirationTime,Uid | sort-object Level,name | out-gridview -Title "Current Partner | $partnername | Select the ParentPartner, individual EndCustomer or Site to list protected M365 domains." -OutputMode Single
        
                if($null -eq $Selection) {
                    # Cancel was pressed
                    # Run cancel script
                    Write-Output    $Script:strLineSeparator
                    Write-Output    "  No Partners Selected"
                    Break
                }else{
                    # OK was pressed, $Selection contains what was chosen
                    # Run OK script
                    [int]$script:PartnerId = $script:Selection.Id
                    [String]$script:PartnerName = $script:Selection.Name
                }
            }
    }
    
}  ## EnumeratePartners API Call

Function Send-GetDevices {
    $ConsoleTitle = "M365 User Cleanup | $Partnername"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle

    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $script.visa
    $data.method = 'EnumerateAccountStatistics'
    $data.params = @{}
    $data.params.query = @{}
    $data.params.query.PartnerId = [int]$PartnerId
    $data.params.query.Filter = "AT==2"
    $data.params.query.Columns = @("AR","PF","AN","MN","AL","AU","CD","TS","TL","T3","US","TB","T7","TM","D19F21","GM","D9F27","JM","D5F20","D5F22","LN","AA843")
    $data.params.query.OrderBy = "CD DESC"
    $data.params.query.StartRecordNumber = 0
    $data.params.query.RecordsCount = $devicecount
    $data.params.query.Totals = @("COUNT(AT==2)","SUM(T3)","SUM(US)")

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }  

    $Script:DeviceResponse = Invoke-RestMethod @params 
            
    $Script:DeviceDetail = @()

    ForEach ( $DeviceResult in $DeviceResponse.result.result ) {
        
        $AccountID = [Int]$DeviceResult.AccountId
        Get-AccountInfoById $AccountID

        $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ 
            AccountID        = [Int]$DeviceResult.AccountId;
            PartnerID        = [string]$DeviceResult.PartnerId;
            PartnerName      = $DeviceResult.Settings.AR -join '';
            Reference        = $DeviceResult.Settings.PF -join '';
            Account          = $DeviceResult.Settings.AU -join '';  
            DeviceName       = $DeviceResult.Settings.AN -join '';
            ComputerName     = $DeviceResult.Settings.MN -join '';
            DeviceAlias      = $DeviceResult.Settings.AL -join '';
            Creation         = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '');
            TimeStamp        = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '');  
            LastSuccess      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TL -join '');
            SelectedGB       = [math]::Round([Decimal](($DeviceResult.Settings.T3 -join '') /1GB),2);  
            UsedGB           = [math]::Round([Decimal](($DeviceResult.Settings.US -join '') /1GB),2);
            Last28Days       = ($DeviceResult.Settings.TB -join '')[-1..-28] -join '' -replace "8",[char]0x26a0 -replace "7",[char]0x23f9 -replace "6",[char]0x23f9 -replace "5",[char]0x2611 -replace "2",[char]0x274e -replace "1",[char]0x2BC8 -replace "0",[char]0x274c;
            #Last28Days       = ($DeviceResult.Settings.TB -join '')[-1..-28] -join '' -replace "8",[char]0x26A0 -replace "7",[char]0x23F9 -replace "6",[char]0x23F9 -replace "5",[char]0x2705 -replace "2",[char]0x274C -replace "1",[char]0x27A1 -replace "0",[char]0x274C;
            Last28           = ($DeviceResult.Settings.TB -join '')[-1..-28] -join '' -replace "8","!" -replace "7","!" -replace "6","?" -replace "5","+" -replace "2","-" -replace "1",">" -replace "0","X";
            Errors           = [int]($DeviceResult.Settings.T7 -join '');
            Billable         = [int]($DeviceResult.Settings.TM -join '');
            Shared           = [int]($DeviceResult.Settings.D19F21 -join '');
            OffBoarded       = [int]($DeviceResult.Settings.D9F27 -join '');
            MailBoxes        = [int]($DeviceResult.Settings.GM -join '');
            OneDrive         = [int]($DeviceResult.Settings.JM -join '');
            SPusers          = [int]($DeviceResult.Settings.D5F20 -join '');
            SPsites          = [int]($DeviceResult.Settings.D5F22 -join '');
            TeamsOwners      = [int]($DeviceResult.Settings.D23F20 -join '');
            ExchangeAutoAdd  = ($DeviceResult.Settings.AA3147 -join '') -eq 'true';
            OneDriveAutoAdd  = ($DeviceResult.Settings.AA3148 -join '') -eq 'true';
            SharePointAutoAdd = ($DeviceResult.Settings.AA3149 -join '') -eq 'true';
            TeamsAutoAdd     = ($DeviceResult.Settings.AA3150 -join '') -eq 'true';
            StorageLocation  = $DeviceResult.Settings.LN -join '';
            AccountToken     = $Script:AccountInfoResponse.result.result.Token;
            Notes            = $DeviceResult.Settings.AA843 -join ''; 
        }
    }

} ## EnumerateAccountStatistics API Call

Function GetM365Stats {
    Param([Parameter(Mandatory=$true)][Int]$DeviceId) #end param

    $url2 = "https://api.backup.management/c2c/statistics/devices/id/$deviceid"
    $method = 'GET'

    $params = @{
        Uri         = $url2
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        WebSession  = $websession
        ContentType = 'application/json; charset=utf-8'
    }   

    $Script:M365response = Invoke-RestMethod @params 

    # Show API URL only in debug mode
    if ($CDPDebug) {
        Write-Host "  API URL: $url2" -ForegroundColor DarkGray
    }
    
    # Show user-friendly Cove console link
    $coveLink = "https://backup.management/#/device/$deviceid/cloud-properties/office365/protected-users"
    Write-Host "  $coveLink" -ForegroundColor Cyan

    Get-AccountInfoById $deviceID

    $script:devicestatistics = $M365response.deviceStatistics | Select-object @{N="Partner";E={$device.partnername}},
                                                                                @{N="Account";E={$device.DeviceName}},
                                                                                DisplayName,
                                                                                EmailAddress,
                                                                                Billable,
                                                                                @{N="Shared";E={$_.shared[0] -replace("TRUE","Shared") -replace("FALSE","") }},
                                                                                @{N="MailBox";E={$_.datasources.status[0] -replace("unprotected","") }},
                                                                                @{N="OneDrive";E={$_.datasources.status[1]  -replace("unprotected","")  }},
                                                                                @{N="SharePoint";E={$_.datasources.status[2]  -replace("unprotected","")  }},
                                                                                @{N="UserGuid";E={$_.UserId}},
                                                                                @{N="AccountToken";E={$device.accountToken}} 
                                    

    $script:devicestatistics | foreach-object { if((($_.Mailbox -eq "protected") -and ($_.shared -ne "shared")) -or ($_.OneDrive -eq "protected") -or ($_.SharePoint -eq "protected")) {$_.Billable = "Billable"}} 
    if ($CDPDebug) {
        $devicestatistics | Select-object * | format-table
        if ($gridview) {$devicestatistics | out-gridview -title "$($device.partnername) | $($device.DeviceName)" }
    }
} 

Function Get-AccountInfoById ([int]$AccountId) {
    $url = "https://api.backup.management/jsonapi"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.visa = $Script:Visa
    $data.method = 'GetAccountInfoById'
    $data.params = @{}
    $data.params.accountId = [int]$AccountId

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        WebSession  = $websession
        ContentType = 'application/json; charset=utf-8'
    }   

    $Script:AccountInfoResponse = Invoke-RestMethod @params 
}

Function EnumerateM365Users ($AccountToken) {
    $url = "https://api.backup.management/management_api"
    $method = 'POST'
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'EnumerateUsers'
    $data.params = @{}
    $data.params.accountToken = $accountToken

    $jsondata = (ConvertTo-Json $data -depth 6)

    $params = @{
        Uri         = $url
        Method      = $method
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        WebSession  = $websession
        ContentType = 'application/json; charset=utf-8'
        TimeoutSec  = 10000
    }   

    $Script:EnumerateM365UsersResponse = Invoke-RestMethod @params 

    $script:M365UserStatistics = $EnumerateM365UsersResponse.result.result.Users | Select-Object `
        @{N="Partner"; E={$device.partnername}}, `
        @{N="Account"; E={$device.DeviceName}}, `
        DisplayName, `
        EmailAddress, `
        @{N="MailBoxSelection"; E={$_.ExchangeInfo.Selection}}, `
        @{N="MailboxType"; E={$_.ExchangeInfo.MailboxType}}, `
        @{N="New"; E={$_.IsNew[0] -replace "True", "New" -replace "False", ""}}, `
        @{N="Deleted"; E={$_.IsDeleted[0] -replace "True", "Deleted" -replace "False", ""}}, `
        @{N="Shared"; E={$_.IsShared[0] -replace "True", "Shared" -replace "False", ""}}, `
        @{N="MailBoxLastBackupStatus"; E={$_.ExchangeInfo.LastBackupStatus}}, `
        @{N="MailBoxLastBackupTimestamp"; E={Convert-UnixTimeToDateTime ($_.ExchangeInfo.LastBackupTimestamp)}}, `
        @{N="ExchangeAutoInclude"; E={$Script:EnumerateM365UsersResponse.result.result.ExchangeAutoInclusionType}}, `
        @{N="OneDriveSelection"; E={$_.OneDriveInfo.Selection}}, `
        @{N="OneDriveStatus"; E={$_.OneDriveInfo.LicenseStatus}}, `
        @{N="OneDriveSelectedGib"; E={[math]::Round([Decimal](($_.OneDriveInfo.SelectedSize) / 1GB), 3)}}, `
        @{N="OneDriveLastBackupStatus"; E={$_.OneDriveInfo.LastBackupStatus}}, `
        @{N="OneDriveLastBackupTimestamp"; E={Convert-UnixTimeToDateTime ($_.OneDriveInfo.LastBackupTimestamp)}}, `
        @{N="OneDriveAutoInclude"; E={$Script:EnumerateM365UsersResponse.result.result.OneDriveAutoInclusionType}}, `
        UserGuid, `
        @{N="AccountToken"; E={$accounttoken}}

    if ($CDPDebug) {
        $Script:M365UserStatistics | Select-object * | format-table
        if ($gridview) {$Script:M365UserStatistics | out-gridview -title "$($device.partnername) | $($device.DeviceName)" }
    } else {
        Write-Host "`n  Domain: $($device.DeviceName) | Users: $($Script:M365UserStatistics.Count)" -ForegroundColor Cyan
    }
}

Function Join-Reports {

    $Script:M365Users = Join-Object -left $script:M365UserStatistics -LeftJoinProperty UserGuid -right $script:devicestatistics -RightJoinProperty UserGuid -Prefix 'Protected_' -Type AllInLeft -RightProperties Billable,Mailbox,OneDrive,SharePoint

    $Script:M365UserStatistics = $Script:M365Users | Select-Object  Partner,
                                                                    Account,
                                                                    DisplayName,
                                                                    EmailAddress,
                                                                    Protected_Billable,
                                                                    New,
                                                                    Deleted,
                                                                    Shared,
                                                                    Protected_Mailbox,
                                                                    MailBoxSelection,
                                                                    MailboxType,
                                                                    MailBoxLastBackupStatus,
                                                                    MailBoxLastBackupTimestamp,
                                                                    ExchangeAutoInclude,
                                                                    Protected_OneDrive,
                                                                    OneDriveSelection,
                                                                    OneDriveStatus,
                                                                    OneDriveSelectedGib,
                                                                    OneDriveLastBackupStatus,
                                                                    OneDriveLastBackupTimestamp,
                                                                    OneDriveAutoInclude,
                                                                    Protected_SharePoint,
                                                                    UserGuid,
                                                                    AccountToken	
}

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


Function UpdateDataSource {
    param (
        [string]$AccountToken,
        [string]$EntityId,
        [string]$MailBoxSelection,
        [string]$ExchangeAutoInclude,
        [string]$OneDriveSelection,
        [string]$OneDriveAutoInclude
    )

    $headers = @{
        "Authorization" = "Bearer $Script:visa"
        "Content-Type"  = "application/json"
    }

    $body = @{
        method = "UpdateDataSources"
        params = @{
            accountToken = $AccountToken
            dataSources = @{
                DataSourceEntityAutoInclusions = @(
                    @{
                        DataSourceType = "Exchange"
                        EntityAutoInclusionType = $ExchangeAutoInclude
                    },
                    @{
                        DataSourceType = "OneDrive"
                        EntityAutoInclusionType = $OneDriveAutoInclude
                    }
                )
                Selections = @(
                    @{
                        EntityId = $EntityId
                        DataSourceSelections = @(
                            @{
                                DataSourceType = "Exchange"
                                SelectionType = $MailBoxSelection
                            },
                            @{
                                DataSourceType = "OneDrive"
                                SelectionType = $OneDriveSelection
                            }
                        )
                    }
                )
            }
        }
        jsonrpc = "2.0"
        id      = "jsonrpc"
    }

    $jsonBody = $body | ConvertTo-Json -Depth 6

    $Script:UpdateDataSource = Invoke-RestMethod -Uri 'https://api.backup.management/management_api' -Method 'POST' -Headers $headers -Body $jsonBody

    if ($Script:UpdateDataSource.error) {
        Write-Output $Script:UpdateDataSource.error.message $EntityId
    } else {
        Write-Output "Updating UserGuid $EntityId"
    }
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

Function ExitRoutine {
    Write-Output $Script:strLineSeparator
    Write-Output "  Secure credential file found here:"
    Write-Output $Script:strLineSeparator
    Write-Output "  & $APIcredfile"
    Write-Output ""
    Write-Output $Script:strLineSeparator
    Start-Sleep -Seconds 3
}

Send-APICredentialsCookie

# Use PartnerName parameter if provided, otherwise use credential file partner name
if ($PartnerName) {
    Send-GetPartnerInfo $PartnerName
} else {
    Send-GetPartnerInfo $Script:cred0
}

# Create execution log file with header now that partner info is available
if ($exportcombined) {
    $logHeader = "M365 User Cleanup Execution Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $logHeader | Set-Content -Path $Script:executionLogFile -ErrorAction SilentlyContinue
    "Partner: $Script:PartnerName (ID: $Script:PartnerId)" | Add-Content -Path $Script:executionLogFile -ErrorAction SilentlyContinue
    "="*80 | Add-Content -Path $Script:executionLogFile -ErrorAction SilentlyContinue
    "Timestamp | Operation | Email | UserGuid | Result | Details" | Add-Content -Path $Script:executionLogFile -ErrorAction SilentlyContinue
    "="*80 | Add-Content -Path $Script:executionLogFile -ErrorAction SilentlyContinue
}

if ($AllPartners) {}else{Send-EnumeratePartners}

if (Get-Module -ListAvailable -Name Join-Object) {
    Write-Host "Module Join-Object Already Installed"
} 
else {
    try {
        Install-Module -Name Join-Object   -Confirm:$True -Force
    }
    catch [Exception] {
        $_.message
        Write-Warning "PS Module 'Join-Object' not found, run with Administrator rights to install" 
        exit
    }
}

Send-GetDevices $partnerId

if ($AllDevices) {
    $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,SelectedGB,UsedGB,Last28Days,Errors,Billable,OffBoarded,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,AccountToken,Notes
}else{
    $script:SelectedDevices = $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,SelectedGB,UsedGB,Last28Days,Errors,Billable,OffBoarded,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,AccountToken,Notes  | Out-GridView -title "Current Partner | $partnername | Select Domains to analyze." -OutputMode Multiple}
    Visa-Check
if ($null -eq $SelectedDevices) {
    # Cancel was pressed
    # Run cancel script
    Write-Output    $Script:strLineSeparator
    Write-Output    "  No Devices Selected"
    Break
}else{
    # OK was pressed, $Selection contains what was chosen
    # Run OK script
    if ($CDPDebug) {
        $DeviceDetail | Select-Object PartnerId,PartnerName,Reference,AccountID,DeviceName,SelectedGB,UsedGB,Last28,Errors,Billable,OffBoarded,MailBoxes,Shared,OneDrive,SPusers,SPsites,StorageLocation,TimeStamp,LastSuccess,Creation,Notes  | Sort-object PartnerName,AccountId | format-table
    }

    If ($Script:Exportcombined) {
        $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_M365_Devices_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $SelectedDevices | Select-object * | Export-CSV -Delimiter $delimiter -path "$csvoutputfile" -NoTypeInformation -Encoding UTF8
    }
}

foreach ($device in $SelectedDevices) {
    Visa-Check
    GetM365Stats $Device.AccountID
    Show-ElapsedTime
    Visa-Check
    EnumerateM365Users $device.AccountToken
    Show-ElapsedTime
    Visa-Check
    join-reports
    
    # Display comprehensive summary matching Cove console
    if (-not $CDPDebug) {
        Write-Host "`n  === Protected Users Statistics ===" -ForegroundColor Cyan
        
        # Get device-level statistics from current device
        $currentDevice = $SelectedDevices | Where-Object { $_.AccountID -eq $device.AccountID }
        
        # Total Billable
        $totalBillable = $currentDevice.Billable
        Write-Host "  Total billable: $totalBillable" -ForegroundColor White
        
        # Datasource counts (Exchange, OneDrive, SharePoint, Teams)
        $exchangeProtected = ($Script:M365UserStatistics | Where-Object { $_.Protected_Mailbox -eq 'protected' }).Count
        $oneDriveProtected = ($Script:M365UserStatistics | Where-Object { $_.Protected_OneDrive -eq 'protected' }).Count
        $sharePointUsers = $currentDevice.SPusers
        $sharePointSites = $currentDevice.SPsites
        $teamsOwners = $currentDevice.TeamsOwners
        
        Write-Host "  Exchange: $exchangeProtected | OneDrive: $oneDriveProtected | SharePoint: $sharePointUsers users, $sharePointSites sites | Teams: $teamsOwners owners" -ForegroundColor Gray
        
        # License status (from MailboxType field in user data)
        $licensedUsers = ($Script:M365UserStatistics | Where-Object { $_.MailboxType -eq 'Licensed' }).Count
        $unlicensedUsers = ($Script:M365UserStatistics | Where-Object { $_.MailboxType -eq 'Unlicensed' }).Count
        Write-Host "  Licensed: $licensedUsers | Unlicensed: $unlicensedUsers" -ForegroundColor Gray
        
        # Mailbox types
        $regularMailboxes = ($Script:M365UserStatistics | Where-Object { $_.Shared -ne 'Shared' -and $_.Deleted -ne 'Deleted' }).Count
        $sharedMailboxes = ($Script:M365UserStatistics | Where-Object { $_.Shared -eq 'Shared' }).Count
        Write-Host "  Regular mailboxes: $regularMailboxes | Shared mailboxes: $sharedMailboxes" -ForegroundColor Gray
        
        # Billing status
        $billableCount = ($Script:M365UserStatistics | Where-Object { $_.Protected_Billable -eq 'Billable' }).Count
        $freeCount = ($Script:M365UserStatistics | Where-Object { $_.Protected_Billable -ne 'Billable' -or $null -eq $_.Protected_Billable }).Count
        Write-Host "  Billable: $billableCount | Free: $freeCount" -ForegroundColor Gray
        
        # User status
        $deletedUsers = ($Script:M365UserStatistics | Where-Object { $_.Deleted -eq 'Deleted' }).Count
        $activeUsers = ($Script:M365UserStatistics | Where-Object { $_.Deleted -ne 'Deleted' }).Count
        Write-Host "  Active: $activeUsers | Deleted: $deletedUsers" -ForegroundColor Gray
        
        Write-Host "`n  === Action Items ===" -ForegroundColor Cyan
        
        # Unlicensed Billable (costing money unnecessarily)
        $unlicensedBillable = ($Script:M365UserStatistics | Where-Object { 
            ($_.MailboxType -eq 'Unlicensed') -and ($_.Protected_Billable -eq 'Billable') 
        }).Count
        if ($unlicensedBillable -gt 0) {
            Write-Host "  Unlicensed Billable: $unlicensedBillable" -ForegroundColor Yellow
        }
        
        # Unprotected users - separated by reason
        # Unlicensed users (automatically unprotected due to no license)
        $mailUnprotectedUnlicensed = ($Script:M365UserStatistics | Where-Object { 
            $_.MailboxType -eq 'Unlicensed' -and 
            (($_.MailBoxSelection -eq 'Excluded') -or ($_.MailBoxSelection -eq 'Unselected'))
        }).Count
        
        $oneDriveUnprotectedUnlicensed = ($Script:M365UserStatistics | Where-Object { 
            $_.OneDriveStatus -eq 'Unlicensed' -and 
            (($_.OneDriveSelection -eq 'Excluded') -or ($_.OneDriveSelection -eq 'Unselected'))
        }).Count
        
        # Not selected (licensed but excluded/unselected)
        $mailUnprotectedNotSelected = ($Script:M365UserStatistics | Where-Object { 
            $_.MailboxType -eq 'Licensed' -and 
            (($_.MailBoxSelection -eq 'Excluded') -or ($_.MailBoxSelection -eq 'Unselected'))
        }).Count
        
        $oneDriveUnprotectedNotSelected = ($Script:M365UserStatistics | Where-Object { 
            $_.OneDriveStatus -ne 'Unlicensed' -and 
            (($_.OneDriveSelection -eq 'Excluded') -or ($_.OneDriveSelection -eq 'Unselected'))
        }).Count
        
        # Total unprotected
        $mailUnprotectedTotal = $mailUnprotectedUnlicensed + $mailUnprotectedNotSelected
        $oneDriveUnprotectedTotal = $oneDriveUnprotectedUnlicensed + $oneDriveUnprotectedNotSelected
        
        if ($mailUnprotectedTotal -gt 0 -or $oneDriveUnprotectedTotal -gt 0) {
            Write-Host "  Unprotected - Mail: $mailUnprotectedTotal (unlicensed: $mailUnprotectedUnlicensed, not selected: $mailUnprotectedNotSelected) | OneDrive: $oneDriveUnprotectedTotal (unlicensed: $oneDriveUnprotectedUnlicensed, not selected: $oneDriveUnprotectedNotSelected)" -ForegroundColor Yellow
        }
        
        # Offboarding Statistics
        $offboarded = ($Script:M365UserStatistics | Where-Object { 
            ($_.Deleted -eq 'Deleted') -and 
            (($_.MailBoxSelection -eq 'Excluded') -or ($_.MailBoxSelection -eq 'Unselected')) -and
            (($_.OneDriveSelection -eq 'Excluded') -or ($_.OneDriveSelection -eq 'Unselected')) -and
            (-not $_.MailBoxLastBackupTimestamp -or $_.MailBoxLastBackupTimestamp -eq '') -and
            (-not $_.OneDriveLastBackupTimestamp -or $_.OneDriveLastBackupTimestamp -eq '')
        }).Count
        
        $offboardable = ($Script:M365UserStatistics | Where-Object { 
            (
                (
                    ($_.Deleted -eq 'Deleted') -and 
                    (
                        (($_.MailBoxSelection -eq 'Selected') -or ($_.MailBoxSelection -eq 'Included')) -or
                        (($_.OneDriveSelection -eq 'Selected') -or ($_.OneDriveSelection -eq 'Included')) -or
                        ($_.MailBoxLastBackupTimestamp -and $_.MailBoxLastBackupTimestamp -ne '') -or
                        ($_.OneDriveLastBackupTimestamp -and $_.OneDriveLastBackupTimestamp -ne '')
                    )
                ) -or
                (
                    ($_.Shared -eq 'Shared') -and 
                    ($_.Protected_Billable -eq 'Billable') -and
                    (
                        (($_.OneDriveSelection -eq 'Selected') -or ($_.OneDriveSelection -eq 'Included')) -or
                        ($_.OneDriveLastBackupTimestamp -and $_.OneDriveLastBackupTimestamp -ne '')
                    )
                ) -or
                (
                    ($_.MailboxType -eq 'Unlicensed') -and 
                    (
                        ($_.MailBoxLastBackupTimestamp -and $_.MailBoxLastBackupTimestamp -ne '') -or
                        ($_.OneDriveLastBackupTimestamp -and $_.OneDriveLastBackupTimestamp -ne '')
                    )
                )
            )
        }).Count
        
        if ($offboarded -gt 0 -or $offboardable -gt 0) {
            Write-Host "  Offboarded: $offboarded | Offboardable: $offboardable" -ForegroundColor Magenta
        }
        
        # Auto-add configuration
        Write-Host "`n  === Auto-Add Configuration ===" -ForegroundColor Cyan
        $exchangeAutoAddStatus = if ($currentDevice.ExchangeAutoAdd) { "Enabled" } else { "Disabled" }
        $oneDriveAutoAddStatus = if ($currentDevice.OneDriveAutoAdd) { "Enabled" } else { "Disabled" }
        $sharePointAutoAddStatus = if ($currentDevice.SharePointAutoAdd) { "Enabled" } else { "Disabled" }
        $teamsAutoAddStatus = if ($currentDevice.TeamsAutoAdd) { "Enabled" } else { "Disabled" }
        
        $autoAddColor = "Gray"
        if ($currentDevice.ExchangeAutoAdd -or $currentDevice.OneDriveAutoAdd -or $currentDevice.SharePointAutoAdd -or $currentDevice.TeamsAutoAdd) {
            $autoAddColor = "Green"
        }
        
        Write-Host "  Exchange: $exchangeAutoAddStatus | OneDrive: $oneDriveAutoAddStatus | SharePoint: $sharePointAutoAddStatus | Teams: $teamsAutoAddStatus" -ForegroundColor $autoAddColor
        
        Write-Host ""
    }
    
    Visa-Check
    If ($Exportcombined) {
        $Script:csvoutputfile2b = "$ExportPath\$($CurrentDate)_M365_User_Export_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
        $Script:M365UserStatistics | Select-object * | Export-CSV -Delimiter $delimiter -path "$csvoutputfile2b" -NoTypeInformation -Encoding UTF8 -append
    }

    If ($csvoutputfile3b) {
        $xlsoutputfile3b = $csvoutputfile3b.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile3b
    }

    If ($csvoutputfile4) {
        $xlsoutputfile4 = $csvoutputfile4.Replace("csv","xlsx")
        Save-CSVasExcel $csvoutputfile4
    }
}

## Generate XLS from CSV

if ($csvoutputfile) {
    $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
    Save-CSVasExcel $csvoutputfile
}

If ($csvoutputfile2b) {
    $xlsoutputfile2b = $csvoutputfile2b.Replace("csv","xlsx")
    Save-CSVasExcel $csvoutputfile2b
}

Write-output $Script:strLineSeparator

## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
    
# This section handles the export and cleanup of shared billable OneDrive backups.
# If the $ExportSharedBillable flag is set, it imports the shared billable data from a CSV file,
# filters the data based on the OneDrive last backup timestamp, and exports the filtered data to a new CSV file.
# The export file path is constructed using the current date, partner name, and partner ID.
# The script then outputs the path of the generated CSV file.
#
# If the $CleanupSharedBillable flag is set, it converts the OneDriveLastBackupTimestamp field to a [datetime] object,
# filters the shared billable data again, and displays the filtered data in an Out-GridView window for selection.
# The user can select accounts to purge OneDrive backup history.
# For each selected account, the script updates the data sources to exclude OneDrive for the user
# and attempts to remove the OneDrive backup history for the shared mailbox.
if ($ExportSharedBillable) {
    $SharedBillable = Import-Csv $Script:csvoutputfile2b
    $filterdate1 = (Get-Date).AddDays(-$OneDriveAge)
    $Script:csvoutputfile2bs = "$ExportPath\$($CurrentDate)_M365_User_Export_Shared_Billable_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
    $SharedBillable = $SharedBillable | Where-Object {$_.OneDriveLastBackupTimestamp} 
    #$SharedBillable | Where-Object {($_.shared -eq "Shared") -and ($_.Protected_Billable -eq "Billable") -and ([datetime]$_.OneDriveLastBackupTimestamp -lt $filterdate1)} | Export-Csv -Path $csvoutputfile2bs -NoTypeInformation
    $SharedBillable = $SharedBillable | Where-Object {($_.shared -eq "Shared") -and ($_.Protected_Billable -eq "Billable") -and ([datetime]$_.OneDriveLastBackupTimestamp -lt $filterdate1) -and ([datetime]$_.OneDriveLastBackupTimestamp -gt $(Get-Date "2000/01/01"))} ## Thanks to morgan@shift.support for the suggestion

    $SharedBillable  | Export-Csv -Path $csvoutputfile2bs -NoTypeInformation 

    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $Script:csvoutputfile2bs"
    Write-output $Script:strLineSeparator

    if ($CleanupSharedBillable) {
        # Convert the OneDriveLastBackupTimestamp field to a [datetime] object
        $SharedBillable = $SharedBillable | ForEach-Object {
            if ($_.OneDriveLastBackupTimestamp) {$_.OneDriveLastBackupTimestamp = [datetime]$_.OneDriveLastBackupTimestamp}
            if ($_.MailBoxLastBackupTimestamp) {$_.MailBoxLastBackupTimestamp = [datetime]$_.MailBoxLastBackupTimestamp}
            $_
        }
        
        #$SharedBillableCount = ($SharedBillable | Where-Object {($_.shared -eq "Shared") -and ($_.Protected_Billable -eq "Billable")}).Count
        #$Script:OneDriveToClean = $SharedBillable | Where-Object {($_.shared -eq "Shared") -and ($_.Protected_Billable -eq "Billable") -and ([datetime]$_.OneDriveLastBackupTimestamp -lt $filterdate1) -and ([datetime]$_.OneDriveLastBackupTimestamp -gt $(Get-Date "2000/01/01"))} | Out-GridView -Title "$($SharedBillable.count) Billable Shared Mailboxes Due to Historic OneDrive Retention > $OneDriveAge days | Select Accounts to Purge OneDrive Backup History" -OutputMode Multiple

        # Filter out protected emails before cleanup
        $SharedBillable = $SharedBillable | Where-Object { -not (Test-ProtectedEmail -EmailAddress $_.emailAddress) }
        
        if ($SharedBillable.count -eq 0) {
            Write-Host "  No unprotected shared billable mailboxes to clean (all are in DoNotDelete list)" -ForegroundColor Yellow
        } else {
            if ($ForceClean) {
                Write-Host "  ForceClean enabled - automatically selecting all $($SharedBillable.count) shared billable mailboxes for cleanup" -ForegroundColor Yellow
                $Script:OneDriveToClean = $SharedBillable
            } else {
                $Script:OneDriveToClean = $SharedBillable | Out-GridView -Title "$($SharedBillable.count) Billable Shared Mailboxes Due to Historic OneDrive Retention > $OneDriveAge days | Select Accounts to Purge OneDrive Backup History" -OutputMode Multiple
            }
        }

        foreach ($item in $OneDriveToClean) {
            # Double-check protection (in case user manually selected protected email)
            if (Test-ProtectedEmail -EmailAddress $item.emailAddress) {
                Write-Host "  SKIPPED: $($item.emailAddress) is in DoNotDelete list" -ForegroundColor Yellow
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Shared Billable - Email is protected (DoNotDelete list)"
                continue
            }
            visa-check
            Write-output "`nAttempting to remove OneDrive selection from Shared Mailbox"

            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Exclude OneDrive from backup selection")) {
                UpdateDataSources $item.AccountToken $item.UserGuid $item.MailBoxSelection $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Maintain current Mailbox selection, Exclude Onedrive for user
                if ($Script:UpdateDataSource.error) {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Shared Billable - Error: $($Script:UpdateDataSource.error.message)"
                } else {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Shared Billable - OneDrive excluded"
                }
            } else {
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Shared Billable - WhatIf mode, operation not executed" -WhatIf
            }

            Write-output "Attempting OneDrive removal for Shared Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove OneDrive backup history")) {
                Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Shared Billable - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Shared Billable - OneDrive history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Shared Billable - WhatIf mode, operation not executed" -WhatIf
            }
        }
    }
}

#region ----- Export Unlicensed ----

# This section of the script handles the export and optional cleanup of unlicensed billable M365 users.
# 
# If the $ExportUnlicensed flag is set to $true, the script performs the following actions:
# 1. Imports the CSV file specified by $Script:csvoutputfile2b into the $UnlicensedBillable variable.
# 2. Sets a filter date ($filterdate1) to the current date minus the number of days specified by $UnlicensedAge.
# 3. Constructs the output file path for the unlicensed billable users CSV file.
# 4. Filters the $UnlicensedBillable data to include only users who are unlicensed, billable, and meet the specified conditions for mailbox and OneDrive backup timestamps.
# 5. Exports the filtered data to the CSV file specified by $Script:csvoutputfileUnlicensedBillable.
# 6. Outputs the path of the generated CSV file.
#
# If the $CleanupUnlicensed flag is set to $true, the script performs the following additional actions:
# 1. Displays a grid view of the unlicensed billable users, allowing the user to select accounts to purge.
# 2. Iterates through the selected accounts and performs the following actions for each:
#    a. Calls the visa-check function (assumed to be defined elsewhere in the script).
#    b. Outputs a message indicating the attempt to remove mail and OneDrive selections from the unlicensed mailbox.
#    c. Calls the UpdateDataSources function to exclude the mailbox and OneDrive for the user.
#    d. Outputs a message indicating the attempt to remove the unlicensed mailbox.
#    e. Calls the Remove-BackupHistory function to remove the backup history for Exchange and OneDrive for the user.

if ($ExportUnlicensed) {
    $UnlicensedBillable = Import-Csv $Script:csvoutputfile2b
    $filterdate2 = (Get-Date).AddDays(-$UnlicensedAge)
    $Script:csvoutputfileUnlicensedBillable = "$ExportPath\$($CurrentDate)_M365_User_Export_Unlicensed_Billable_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"

    $UnlicensedBillable = $UnlicensedBillable | Where-Object {
        ($_.MailboxType -eq "Unlicensed") -and 
        ($_.Protected_Billable -eq "Billable") -and 
        (
            (
                ($_.MailBoxLastBackupTimestamp -ne "") -and ([datetime]$_.MailBoxLastBackupTimestamp -lt $Filterdate2) -or 
                ($_.MailBoxLastBackupTimestamp -eq "")
            ) -and
            (
                ($_.OneDriveLastBackupTimestamp -ne "") -and ([datetime]$_.OneDriveLastBackupTimestamp -lt $filterDate2) -or 
                ($_.OneDriveLastBackupTimestamp -eq "")
            )
        )
    }

    $UnlicensedBillable | Export-Csv -Path $Script:csvoutputfileUnlicensedBillable -NoTypeInformation

    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $Script:csvoutputfileUnlicensedBillable"
    Write-output $Script:strLineSeparator

    if ($CleanupUnlicensed) {

        # Convert the OneDriveLastBackupTimestamp field to a [datetime] object
        $UnlicensedBillable = $UnlicensedBillable | ForEach-Object {
            if ($_.OneDriveLastBackupTimestamp) {$_.OneDriveLastBackupTimestamp = [datetime]$_.OneDriveLastBackupTimestamp}
            if ($_.MailBoxLastBackupTimestamp) {$_.MailBoxLastBackupTimestamp = [datetime]$_.MailBoxLastBackupTimestamp}
            $_
        }
        
        # Filter out protected emails before cleanup
        $UnlicensedBillable = $UnlicensedBillable | Where-Object { -not (Test-ProtectedEmail -EmailAddress $_.emailAddress) }
        
        if ($UnlicensedBillable.count -eq 0) {
            Write-Host "  No unprotected unlicensed mailboxes to clean (all are in DoNotDelete list)" -ForegroundColor Yellow
        } else {
            if ($ForceClean) {
                Write-Host "  ForceClean enabled - automatically selecting all $($UnlicensedBillable.count) unlicensed mailboxes for cleanup" -ForegroundColor Yellow
                $Script:UnlicensedToClean = $UnlicensedBillable
            } else {
                $Script:UnlicensedToClean = $UnlicensedBillable | Out-GridView -Title "$($UnlicensedBillable.count) Unlicensed Mailboxes Due to Historic Mail or OneDrive Retention > $UnlicensedAge days | Select Accounts to Purge Mail and OneDrive Backup History" -OutputMode Multiple
            }
        }

        foreach ($item in $UnlicensedToClean) {
            # Double-check protection (in case user manually selected protected email)
            if (Test-ProtectedEmail -EmailAddress $item.emailAddress) {
                Write-Host "  SKIPPED: $($item.emailAddress) is in DoNotDelete list" -ForegroundColor Yellow
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Unlicensed - Email is protected (DoNotDelete list)"
                continue
            }
            visa-check
            Write-output "`nAttempting to remove Mail and OneDrive selections from Unlicensed Mailbox"

            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Exclude Exchange and OneDrive from backup selection")) {
                UpdateDataSources $item.AccountToken $item.UserGuid "Excluded" $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Exclude Mailbox selection & Onedrive for user
                if ($Script:UpdateDataSource.error) {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Unlicensed - Error: $($Script:UpdateDataSource.error.message)"
                } else {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Unlicensed - Exchange and OneDrive excluded"
                }
            } else {
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Unlicensed - WhatIf mode, operation not executed" -WhatIf
            }

            Write-output "Attempting removal for Unlicensed Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove Exchange backup history")) {
                Remove-BackupHistory "Exchange" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Unlicensed - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Unlicensed - Exchange history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Unlicensed - WhatIf mode, operation not executed" -WhatIf
            }
            
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove OneDrive backup history")) {
                Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Unlicensed - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Unlicensed - OneDrive history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Unlicensed - WhatIf mode, operation not executed" -WhatIf
            }
        }
    }
} 

#endregion ----- Export Unlicensed ----

#region ----- Export Deleted ----

# This section handles the export and optional cleanup of deleted M365 users.
# If the $ExportDeleted flag is set to $true, it imports the CSV file containing deleted users' data.
# It then filters the deleted users based on the $DeletedAge parameter and their last backup timestamps.
# The filtered data is exported to a new CSV file.
# If the $CleanupDeletedBillable flag is set to $true, it allows the user to select billable, non-shared deleted users for cleanup.
# If the $CleanupDeletedBillableShared flag is set to $true, it allows the user to select billable, shared deleted users for cleanup.
# If the $CleanupDeletedShared flag is set to $true, it allows the user to select non-billable, shared deleted users for cleanup.
# For each selected user, it attempts to remove their mail and OneDrive selections and purge their backup history.

if ($ExportDeleted) {
    $Deleted = Import-Csv $Script:csvoutputfile2b
    $filterdate3 = (Get-Date).AddDays(-$DeletedAge)
    $Script:csvoutputfileDeleted = "$ExportPath\$($CurrentDate)_M365_User_Export_Deleted_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
    
    #$Deleted = $Deleted |  Where-Object {$_.MailBoxLastBackupTimestamp -or $_.OneDriveLastBackupTimestamp} 
    
    $Deleted = $Deleted | Where-Object {
        ($_.deleted -eq "Deleted") -and 
        (
            (
                ($_.MailBoxLastBackupTimestamp -ne "") -and ([datetime]$_.MailBoxLastBackupTimestamp -lt $Filterdate3) -or 
                ($_.MailBoxLastBackupTimestamp -eq "")
            ) -and
            (
                ($_.OneDriveLastBackupTimestamp -ne "") -and ([datetime]$_.OneDriveLastBackupTimestamp -lt $filterDate3) -or 
                ($_.OneDriveLastBackupTimestamp -eq "")
            )
        )
    }
    
    $Deleted | Export-Csv -Path $csvoutputfileDeleted -NoTypeInformation
    
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $Script:csvoutputfileDeleted"
    Write-output $Script:strLineSeparator
    
    if ($CleanupDeletedBillable) {
        # Filter deleted billable users, excluding protected emails
        $DeletedBillableFiltered = $Deleted | Where-Object {
            ($_.Protected_Billable -eq "Billable") -and 
            ($_.Shared -ne "Shared") -and
            (-not (Test-ProtectedEmail -EmailAddress $_.emailAddress))
        }
        
        $BillableNonSharedCount = $DeletedBillableFiltered.Count
        
        if ($BillableNonSharedCount -eq 0) {
            Write-Host "  No unprotected deleted billable (non-shared) mailboxes to clean" -ForegroundColor Yellow
        } else {
            if ($ForceClean) {
                Write-Host "  ForceClean enabled - automatically selecting all $BillableNonSharedCount deleted billable (non shared) mailboxes for cleanup" -ForegroundColor Yellow
                $Script:DeletedToClean = $DeletedBillableFiltered
            } else {
                $Script:DeletedToClean = $DeletedBillableFiltered | Out-GridView -Title "$BillableNonSharedCount Deleted Billable (Non Shared) Mailboxes with Retention > $DeletedAge days | Select Accounts to Purge Backup History" -OutputMode Multiple
            }
        }

        foreach ($item in $DeletedToClean) {
            # Double-check protection
            if (Test-ProtectedEmail -EmailAddress $item.emailAddress) {
                Write-Host "  SKIPPED: $($item.emailAddress) is in DoNotDelete list" -ForegroundColor Yellow
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable - Email is protected (DoNotDelete list)"
                continue
            }
            visa-check
            Write-output "Attempting to remove Mail and OneDrive selections from Deleted Mailbox"

            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Exclude Exchange and OneDrive from backup selection")) {
                UpdateDataSources $item.AccountToken $item.UserGuid "Excluded" $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Exclude Mailbox selection & Onedrive for user
                if ($Script:UpdateDataSource.error) {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable - Error: $($Script:UpdateDataSource.error.message)"
                } else {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable - Exchange and OneDrive excluded"
                }
            } else {
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable - WhatIf mode, operation not executed" -WhatIf
            }

            Write-output "`nAttempting removal for Deleted Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove Exchange backup history")) {
                Remove-BackupHistory "Exchange" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable - Exchange history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable - WhatIf mode, operation not executed" -WhatIf
            }
            
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove OneDrive backup history")) {
                Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable - OneDrive history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable - WhatIf mode, operation not executed" -WhatIf
            }
        }
    }

    if ($CleanupDeletedBillableShared) {
        # Filter deleted billable shared users, excluding protected emails
        $DeletedBillableSharedFiltered = $Deleted | Where-Object {
            ($_.Protected_Billable -eq "Billable") -and 
            ($_.Shared -eq "Shared") -and
            (-not (Test-ProtectedEmail -EmailAddress $_.emailAddress))
        }
        
        $BillableSharedCount = $DeletedBillableSharedFiltered.Count
        
        if ($BillableSharedCount -eq 0) {
            Write-Host "  No unprotected deleted billable shared mailboxes to clean" -ForegroundColor Yellow
        } else {
            if ($ForceClean) {
                Write-Host "  ForceClean enabled - automatically selecting all $BillableSharedCount deleted billable shared mailboxes for cleanup" -ForegroundColor Yellow
                $Script:DeletedToClean = $DeletedBillableSharedFiltered
            } else {
                $Script:DeletedToClean = $DeletedBillableSharedFiltered | Out-GridView -Title "$BillableSharedCount Deleted Billable Shared Mailboxes with Retention > $DeletedAge days | Select Accounts to Purge Backup History" -OutputMode Multiple
            }
        }

        foreach ($item in $DeletedToClean) {
            # Double-check protection
            if (Test-ProtectedEmail -EmailAddress $item.emailAddress) {
                Write-Host "  SKIPPED: $($item.emailAddress) is in DoNotDelete list" -ForegroundColor Yellow
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable Shared - Email is protected (DoNotDelete list)"
                continue
            }
            visa-check
            Write-output "`nAttempting to remove Mail and OneDrive selections from Deleted Mailbox"

            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Exclude Exchange and OneDrive from backup selection")) {
                UpdateDataSources $item.AccountToken $item.UserGuid "Excluded" $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Exclude Mailbox selection & Onedrive for user
                if ($Script:UpdateDataSource.error) {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable Shared - Error: $($Script:UpdateDataSource.error.message)"
                } else {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable Shared - Exchange and OneDrive excluded"
                }
            } else {
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable Shared - WhatIf mode, operation not executed" -WhatIf
            }

            Write-output "Attempting removal for Deleted Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove Exchange backup history")) {
                Remove-BackupHistory "Exchange" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable Shared - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable Shared - Exchange history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable Shared - WhatIf mode, operation not executed" -WhatIf
            }
            
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove OneDrive backup history")) {
                Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Billable Shared - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Billable Shared - OneDrive history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Billable Shared - WhatIf mode, operation not executed" -WhatIf
            }
        }
    }

    if ($CleanupDeletedShared) {
        # Filter deleted shared (non-billable) users, excluding protected emails
        $DeletedSharedFiltered = $Deleted | Where-Object {
            ($_.Protected_Billable -ne "Billable") -and 
            ($_.Shared -eq "Shared") -and
            (-not (Test-ProtectedEmail -EmailAddress $_.emailAddress))
        }
        
        $NonBillableSharedCount = $DeletedSharedFiltered.Count
        
        if ($NonBillableSharedCount -eq 0) {
            Write-Host "  No unprotected deleted shared (non-billable) mailboxes to clean" -ForegroundColor Yellow
        } else {
            if ($ForceClean) {
                Write-Host "  ForceClean enabled - automatically selecting all $NonBillableSharedCount deleted shared (non billable) mailboxes for cleanup" -ForegroundColor Yellow
                $Script:DeletedToClean = $DeletedSharedFiltered
            } else {
                $Script:DeletedToClean = $DeletedSharedFiltered | Out-GridView -Title "$NonBillableSharedCount Deleted Shared (Non Billable) Mailboxes with Retention > $DeletedAge days | Select Accounts to Purge Backup History" -OutputMode Multiple
            }
        }

        foreach ($item in $DeletedToClean) {
            # Double-check protection
            if (Test-ProtectedEmail -EmailAddress $item.emailAddress) {
                Write-Host "  SKIPPED: $($item.emailAddress) is in DoNotDelete list" -ForegroundColor Yellow
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Shared (Non-Billable) - Email is protected (DoNotDelete list)"
                continue
            }
            visa-check
            Write-output "`nAttempting to remove Mail and OneDrive selections from Deleted Mailbox"

            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Exclude Exchange and OneDrive from backup selection")) {
                UpdateDataSources $item.AccountToken $item.UserGuid "Excluded" $item.ExchangeAutoInclude "Excluded" $item.OneDriveAutoInclude ## Exclude Mailbox selection & Onedrive for user
                if ($Script:UpdateDataSource.error) {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Shared (Non-Billable) - Error: $($Script:UpdateDataSource.error.message)"
                } else {
                    Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Shared (Non-Billable) - Exchange and OneDrive excluded"
                }
            } else {
                Write-ExecutionLog -Operation "UpdateDataSources (Exclude Exchange/OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Shared (Non-Billable) - WhatIf mode, operation not executed" -WhatIf
            }

            Write-output "Attempting removal for Deleted Mailbox | $($item.emailAddress) | $($item.UserGuid) | $($item.accounttoken)"
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove Exchange backup history")) {
                Remove-BackupHistory "Exchange" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Shared (Non-Billable) - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Shared (Non-Billable) - Exchange history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (Exchange)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Shared (Non-Billable) - WhatIf mode, operation not executed" -WhatIf
            }
            
            if ($PSCmdlet.ShouldProcess("$($item.emailAddress) ($($item.UserGuid))", "Remove OneDrive backup history")) {
                Remove-BackupHistory "OneDrive" $item.accounttoken $item.UserGuid
                if ($Script:RemovedBackupHistory.error) {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "FAILED" -Details "Deleted Shared (Non-Billable) - Error: $($Script:RemovedBackupHistory.error.message)"
                } else {
                    Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SUCCESS" -Details "Deleted Shared (Non-Billable) - OneDrive history removed"
                }
            } else {
                Write-ExecutionLog -Operation "Remove-BackupHistory (OneDrive)" -Email $item.emailAddress -UserGuid $item.UserGuid -Result "SKIPPED" -Details "Deleted Shared (Non-Billable) - WhatIf mode, operation not executed" -WhatIf
            }
        }
    }
}


#endregion ----- Export Deleted ----

if ($Launch -and $ExportCombined) {
    If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
        Start-Process "$xlsoutputfile"
        #Start-Process "$xlsoutputfile2a"
        Start-Process "$xlsoutputfile2b"
        Write-output $Script:strLineSeparator
        Write-Output "  Opening XLS file"
        }else{
        Start-Process "$csvoutputfile"
        #Start-Process "$csvoutputfile2a"
        Start-Process "$csvoutputfile2b"
        Write-output $Script:strLineSeparator
        Write-Output "  Opening CSV file"
        Write-output $Script:strLineSeparator            
        }
}

If ($Exportcombined) {
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $Script:csvoutputfile"
    #Write-Output "  CSV Path = $Script:csvoutputfile2a"
    Write-Output "  CSV Path = $Script:csvoutputfile2b"
    Write-Output "  XLS Path = $Script:xlsoutputfile"
    #Write-Output "  XLS Path = $Script:xlsoutputfile2a"
    Write-Output "  XLS Path = $Script:xlsoutputfile2b"
    Write-output $Script:strLineSeparator
    Write-Output "  Execution Log Path = $Script:executionLogFile"
    Write-Output ""
}


ExitRoutine
  