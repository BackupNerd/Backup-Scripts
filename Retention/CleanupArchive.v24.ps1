﻿<# ----- About: ----
    # N-able Cleanup Backup Archives
    # Revision v24 - 2022-07-07
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
    # For use with Cove Data Protection from N-able (Formerly the Standalone editon of N-able Backup)
    # For use with N-central integrated Backup editions
    # Sample scripts may contain non-public API calls which are subject to change without notification
    # Requires Local Administrator or equivilent access
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Script Parameters:
    #
    # [-Months ##] <Int>                                         ## Clean Archives older than ## number of months (Default is 48)
    # [-Force] <switch>                                          ## Do not prompt for confirmation before Archive Cleaning
    # [CheckArchive] <switch>                                    ## Check backup client for an active Archive rule
    # [AddArchive] <switch>                                      ## Add a Monthly Archive rule if no active Archive rules exist 
    # 
    # Run locally as an Adminstrator or equivilent on each Backup Manager Client 
    # N-central and Cove Data Protection compatible
    # Self Copy Script to Backup %ProgramData% Path
    # Check Date Time via NTP Server
    # Get/ Log Current Device/Archive Settings using ClientTool.exe
    # Request Client Authentication Visa 
    # Get Historic Archive Session Times using ClientTool.exe
    # Get Historic Archive Session IDs from SessionReport.xml
    # Normalize DateTime/Match Archive Session Time/ID
    # List ArchiveSessionIDsToClean that are older than -Months ## parameter 
    # Prompt for Confirmation to Clean if not using -Force parameter
    # Pass List to JSON CleanupAchiving Function
    # Log actions
    # (PENDING)Create Recurring SCHTASK
    # (PENDING)AMP FILE  
# -----------------------------------------------------------#>

[CmdletBinding()]
    Param (

        [Parameter(Mandatory=$False)] [ValidateRange(12,84)] [int]$Months=48,       ## Clean Archives older than X number of months
        [Parameter(Mandatory=$False)] $NTPServer="Pool.ntp.org",                    ## NTP Server to validate current Date Time
        [Parameter(Mandatory=$False)] [switch]$Force,                               ## Do nont prompt for confirmation before Archive Cleaning
        [Parameter(Mandatory=$False)] [switch]$CheckArchive,                        ## Check backup client for an Active Archive rule
        [Parameter(Mandatory=$False)] [switch]$AddArchive                           ## Add a Monthly Archive rule if no active Archive rules exist 
               
    )

    Clear-Host

#region ----- Environment, Variables, Names and Paths ----
    $Script:strLineSeparator = "  ---------"
    $True_path = "C:\ProgramData\MXB\Backup Manager"
    $ScriptFull = $myInvocation.MyCommand.path
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ScriptFile = Split-Path -Leaf $MyInvocation.MyCommand.Path
    $ScriptVer = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
    $ScriptBase = $ScriptVer -replace '\..*'
    $ScriptFinal = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath $ScriptFile
    $ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.log"
    $ScriptLogParent = Split-path -path $ScriptLog
    mkdir -Force $ScriptLogParent

    $SetCompressed = Invoke-WmiMethod -Path "Win32_Directory.Name='$ScriptLogParent'" -Name compress
    If (($SetCompressed.returnvalue) -eq 0) { "Items successfully compressed" } else { "Something went wrong!" }

    If ($ScriptFull -eq $ScriptFinal) {
        Write-Output $Script:strLineSeparator 
        Write-Output 'Script Already Running from Target Location'
        Write-Output $Script:strLineSeparator
        } Else {
        Write-Output $Script:strLineSeparator
        Write-Output 'Copying Script to Target Location'
        Write-Output $Script:strLineSeparator
        Copy-item -Path $ScriptFull -Destination $ScriptLogParent -Force
        } ## ----- Self Copy & Logging Logic ----

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

    if (Get-Module -ListAvailable -Name NtpTime) {
        Write-Host "  Module NtpTime Already Installed"
    } 
    else {
        try {
            Install-Module -Name NtpTime -Confirm:$False -Force      ## https://www.powershellgallery.com/packages/NtpTime/1.1
        }
        catch [Exception] {
            $_.message 
            exit
        }
    }

    Function Get-LocalTime($UTCTime) {
        $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
        $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
        $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
        Return $LocalTime
    } ## Get Local Time for comparison vs UTC Time
    
    Function Get-TimeStamp {
        return "[{0:M/d/yyyy} {0:HH:mm:ss}]" -f (Get-Date)
    }  ## Get Timestamp for logging
    
    Function Get-FPvisa {

        $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"
        $config = "C:\Program Files\Backup Manager\config.ini"
    
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { 
            Write-Warning "Backup Manager Not Running"
            Break 
        }else{ 
            try { $ErrorActionPreference = 'Stop'; $UIToken = & $clienttool in-agent-authentication-token.get -config-path $config | convertfrom-json 
            }catch{ 
                Write-Warning "Oops: $_"
                Break 
            }
        }
    
        $UIToken = $UIToken.inagentauthenticationToken
    
        $url = "http://localhost:5000/jsonrpcv1"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = 'jsonrpc'
        $data.method = 'InAgentAuthenticationTokenLogin'
        $data.params = @{}
        $data.params.token = $UIToken
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -TimeoutSec 180 `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
    
        $FPvisa = $webrequest | convertfrom-json
        $script:BackupFPvisa = $FPvisa.result.result

        $script:BackupFPvisa
    }

    Function GetArchiveSettings {

    # ----- Clienttool Get Archive Settings From Local Client ----
        $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"
        & $clienttool -machine-readable control.setting.list > "$ScriptLogParent\DeviceSettings.tsv"
        $Script:DeviceSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\DeviceSettings.tsv"
        $Script:DeviceName = $DeviceSettings | Where-Object { $_.Name -eq "Device" } | Select-Object Value 
        $Script:ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.$($devicename.value).log"
                
        if (($CheckArchive) -or ($AddArchive)) {
        
            & $clienttool -machine-readable control.archiving.list > "$ScriptLogParent\ArchiveSettings.tsv"
            $ArchiveSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\ArchiveSettings.tsv"
        
            Write-Output "  Getting Archive Settings from Clienttool for $($ENV:COMPUTERNAME) | $($DeviceName.value)"
            Write-Output $Script:strLineSeparator
            $ArchiveSettings | format-table -autosize
        
            Write-Output "  $(Get-TimeStamp) Getting current archive settings for $($ENV:COMPUTERNAME) | $($DeviceName.value)" | Out-file $ScriptLog -append
            $ArchiveSettings | format-table | Out-file $ScriptLog -append
    
            $Script:ArchiveCheck = $ArchiveSettings | Where-object{ ($_.ACTV -eq "yes")} | Select-Object -Property ID,ACTV,NAME 
            }
    
    # (Check for Archive Rule / Create if missing)
        If (($CheckArchive) -and ($ArchiveCheck.ACTV) -eq "yes") {
            Write-Output $Script:strLineSeparator
            Write-Output 'Active Archive Schedule Found'
            Write-Output $Script:strLineSeparator
            #$Archivecheck | Where-object{ $_.name -eq "#Monthly - Archive#"} | Select-Object -Property ID,ACTV,NAME | format-table
            $Archivecheck | Select-Object -Property ID,ACTV,NAME | format-table
            
        } 
        
        if (($AddArchive) -and ($ArchiveSettings.ACTV) -notcontains "yes") {
            Write-Output 'Creating Monthly Archive Schedule/s'
            & $clienttool control.archiving.add -name "#Monthly - Archive#" -days-of-month Last
        }
    
    # ----- End Clienttool Get Archive Settings From Local Client ----    
    }  ## Get current Archive settings

    Function ClienttoolGetArchiveSessions {

        & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > "$ScriptLogParent\AllSessions.tsv"
        $Sessions = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\AllSessions.tsv"
    
    # Clienttool Sessions - All
    
        #Write-Output "  All Session History from Clienttool.exe"    
        #$Sessions | format-table
    
    # Clienttool Sessions - Filtered
    
        $Inventory = $Sessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TyPE -eq "Backup") } | Select-Object * | Sort-Object START
        Write-output "`n  10 oldest archive sessions"
        $inventory[0..9]  | sort-object DSRC,START |  format-table 
        Write-output "  10 newest archive sessions`n"   
        $inventory[-10..-1]  | sort-object DSRC,START |  format-table 

        $Script:CleanSessions = $Sessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TyPE -eq "Backup") -and ([datetime]$_.START -lt $filterDate) } | Select-Object -Property START,DSRC,STATE,TYPE,FLAGS,SELS | Sort-Object START 
    
        $Script:CleanSessions | ForEach-Object {$_.START = [datetime]$_.START}
    
        #Write-Output "  All Archive Session History from Clienttool.exe older than $filterdate [datetime]"
        #$Script:CleanSessions |Select-Object -Property DSRC,START,STATE,TYPE,FLAGS,SELS | Sort-Object START | format-table
    }  ## Get Archive session times via Clienttool.exe

    Function XMLLookupArchiveSessionIds {

    # Read SessionReport.XML for all Historical Archives
       
        [xml]$XmlDocument = Get-Content -Path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"
    
        #Write-Output "  Non archive sessions from SessionReport.xml (includes Cleaned Sessions)"

        #$XmlDocument.SessionStatistics.Session | Where-Object{ ($_.Archived -eq "0") } | Select-Object -Property Id,Plugin, StartTimeUTC,Status,Type,Archived | Sort-Object StartTime | format-table
    
        #Write-Output "  All archive sessions from SessionReport.xml older than $Filterdate (includes Cleaned Sessions)"
        #$XmlDocument.SessionStatistics.Session | Where-Object{ ($_.Archived -eq "1") -and ([datetime]$_.StartTimeUTC -lt $filterDate) } | Select-Object -Property Id,Plugin, StartTimeUTC,Status,Type,Archived | Sort-Object StartTime | format-table
    
        $Script:CleanArchive = $XmlDocument.SessionStatistics.Session | Where-Object{ ($_.Archived -eq "1") } | Select-Object -Property @{Name="StartTime";Expression={$_.StartTimeUTC}},Id,Plugin,Status,Type,Archived | Sort-Object StartTime 
        
        $Script:CleanArchive | ForEach-Object {$_.StartTime = get-localtime ([datetime]$_.StartTime)} | Where-Object{ ([datetime]$_.StartTime -lt $filterDate) } | Select-Object -Property Id,Plugin,StartTime,Status,Type,Archived | Sort-Object StartTime |format-table
        
        #$Script:CleanArchive | Where-Object{ ([datetime]$_.StartTime -lt $filterDate) } | Select-Object -Property Id,Plugin,StartTime,Status,Type,Archived | Sort-Object StartTime |format-table
    
    }  ## Get Archive session IDs from Session.xml

    Function Clean-Archives {
        $url = "http://localhost:5000/jsonrpcv1"
        $script:data = @{}
        $data.jsonrpc = '2.0'
        $data.id = 'jsonrpc'
        $data.visa = $BackupFPvisa
        $data.method = 'CleanupArchiving'
        $data.params = @{}
        $data.params.sessionsToCleanup = $ArchiveSessionIDsToClean
    
        $webrequest = Invoke-WebRequest -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body (ConvertTo-Json $data -depth 6) `
        -Uri $url `
        -SessionVariable Script:websession `
        -TimeoutSec 180 `
        -UseBasicParsing
        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession
      
        $script:CleanArchives = $webrequest | convertfrom-json
    
        [void]::$CleanArchives | convertto-json -depth 6
        if ($CleanArchives.error) {$CleanArchives.error.message}
        else {
            
            #Debug $data.method
            #Debug $data.params
            #Debug $data | ConvertTo-Json -depth 6
        }
       
    
    }

#endregion ----- Functions ----

#region ----- Body ----

    $ntptime =  Get-NtpTime $NTPServer -MaxOffset 500000000 | Select-Object NtpTime
    Write-output "  $(get-date) Current system date/time"
    Write-output "  $($NtpTime.ntptime) Current NTP-Server date/time"
    Write-output "  Adjusting system clock to match"
    [datetime]$NtpTime.ntptime | set-date | out-null

    if ($months -lt 12) { 
        Write-Warning "Archive Retention Expiration of < 12 Months Is Not Allowed"; Break
    }else{
        $filterDate = (Get-Date).Addmonths($months * -1)
        Write-Output "  Looking for archive sessions older than $filterdate"
    }
   
    Get-FPvisa
    GetArchiveSettings
    ClienttoolGetArchiveSessions
    XMLLookupArchiveSessionIds
    
    $ArchiveSessionsToClean = Foreach ($row in $Script:CleanSessions) {
    [pscustomobject]@{
        StartTime = $row.Start
        DataSource = $row.DSRC
        State = $row.STATE
        Type = $row.TYPE
        SelectedSize = $row.SELS
        Flags = $row.FLAGS
        ID = $CleanArchive | Where-Object {$_.StartTime -eq $row.Start} | Select-Object -ExpandProperty Id
    }
}
    Write-Output $Script:strLineSeparator
    Write-Output "  Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"

    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags | Sort-Object StartTime | format-table

    Write-Output "  $(Get-TimeStamp) Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate" | Out-file $ScriptLog -append
    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags| Sort-Object StartTime | format-table | Out-file $ScriptLog -append

    [int32[]]$Script:ArchiveSessionIDsToClean = $ArchiveSessionsToClean.id

    Write-Output $Script:strLineSeparator
    Write-Output "  Sessions to Clean" $ArchiveSessionIDsToClean
    Write-Output $Script:strLineSeparator

    if ($($ArchiveSessionsToClean.ID.count) -eq 0) { break }

# (Check for Confirmation)
    If ($force -ne $true) { $strAreYouSure = Read-Host -Prompt "Do you want to clean all Archive Sessions prior to [$Filterdate]? This can not be undone! Type (Y/N)" } 
    
    If (($force) -or ($strAreYouSure -eq "Y")) {
        Write-Output $Script:strLineSeparator
        Write-Output "  Performing Archive Cleaning"
        Write-Output $Script:strLineSeparator
# (If Y, then Clean Archives)

    Clean-Archives
  
# (pause and wait for return)
    Start-Sleep -Milliseconds 200
        
        If ($ArchiveSessionIDsToClean -eq $null) { 
            Write-Output "  Get Cleanup Archiving Description:  Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
            Write-Output $Script:strLineSeparator
    # (Exit Script if there is a problem)
            Break Script
        } else {
        Write-Output "  Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
        Write-Output "  $(Get-TimeStamp) Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate`n" | Out-file $ScriptLog -append
        }
    } else {
    Write-Output $Script:strLineSeparator
    Write-Output "  You chose No, exiting script."
    Write-Output "  $(Get-TimeStamp) You chose Not to Clean Archive Sessions" | Out-file $ScriptLog -append
    Write-Output $Script:strLineSeparator
    #Read-Host -Prompt "Press Enter to exit"
# (If N, End the script and Exit)
    Break Script
}