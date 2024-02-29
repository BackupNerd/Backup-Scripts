<# ----- About: ----
    # N-able Cleanup Backup Archives
    # Revision v24.02.06 - 2024-02-06
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
    # For use with Cove Data Protection from N-able
    # Sample scripts may contain non-public API calls which are subject to change without notification
    # Requires Local Administrator or equivilent access
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Script Parameters:
    #
    # [-Days ##] <Int>                                           ## Clean Archives older than ## number of days (Default is 365)
    # [-Force] <switch>                                          ## Do not prompt for confirmation before Archive Cleaning
    # 
    # Run locally as an Adminstrator or equivilent on each Backup Manager Client 
    # Get/ Log Current Device/Archive Settings using ClientTool.exe
    # Request Client Authentication Visa 
    # Get Historic Archive Session Times using ClientTool.exe
    # Get Historic Archive Session IDs from SessionReport.xml
    # Normalize DateTime/Match Archive Session Time/ID
    # List ArchiveSessionIDsToClean that are older than -Days ## parameter 
    # Prompt for Confirmation to Clean if not using -Force parameter
    # Pass List to JSON CleanupAchiving Function
    # Log actions
# -----------------------------------------------------------#>

#region ----- Environment, Variables, Names and Paths ----
    [int]$days=365
    [switch]$Force=$false
    $Confirm="CONFIRM!"

    $Script:strLineSeparator = "  ---------"
    $ScriptLog = "c:\programdata\MXB\CleanArchive\CleanArchive.log"
    $ScriptLogParent = Split-path -path $ScriptLog
    mkdir -Force $ScriptLogParent

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

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
    
        if ($Null -eq (get-process "BackupFP" -ea SilentlyContinue)) { 
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
    }

    Function GetArchiveSettings {
        $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"
        & $clienttool -machine-readable control.setting.list > "$ScriptLogParent\DeviceSettings.tsv"
        $Script:DeviceSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\DeviceSettings.tsv"
        $Script:DeviceName = $DeviceSettings | Where-Object { $_.Name -eq "Device" } | Select-Object Value 
    }  ## Clienttool Get Archive Settings From Local Client

    Function ClienttoolGetArchiveSessions {
        & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > "$ScriptLogParent\AllSessions.tsv"
        $Sessions = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\AllSessions.tsv"
        $Inventory = $Sessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TyPE -eq "Backup") } | Select-Object * | Sort-Object START

        Write-output "  10 newest archive sessions`n"   
        $inventory[-10..-1]  | sort-object DSRC,START |  format-table
        Write-output "`n  10 oldest archive sessions"
        $inventory[0..9]  | sort-object DSRC,START |  format-table 

        $Script:CleanSessions = $Sessions | Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and ($_.TyPE -eq "Backup") -and ([datetime]$_.START -lt $filterDate) } | Select-Object -Property START,DSRC,STATE,TYPE,FLAGS,SELS | Sort-Object START 
        $Script:CleanSessions | ForEach-Object {$_.START = [datetime]$_.START}
    }  ## Get Archive session times via Clienttool.exe

    Function XMLLookupArchiveSessionIds {
        [xml]$XmlDocument = Get-Content -Path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"
        $Script:CleanArchive = $XmlDocument.SessionStatistics.Session | Where-Object{ ($_.Archived -eq "1") } | Select-Object -Property @{Name="StartTime";Expression={$_.StartTimeUTC}},Id,Plugin,Status,Type,Archived | Sort-Object StartTime 
        $Script:CleanArchive | ForEach-Object {$_.StartTime = get-localtime ([datetime]$_.StartTime)} | Where-Object{ ([datetime]$_.StartTime -lt $filterDate) } | Select-Object -Property Id,Plugin,StartTime,Status,Type,Archived | Sort-Object StartTime |format-table
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
    if ($days -lt 30) { 
        Write-Warning "Archive Retention Expiration of < 3 Months Is Not Allowed With This Script"; Break
    }else{
        $filterDate = (Get-Date).Adddays($days * -1)
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
    #If ($force -ne $true) { $strAreYouSure = Read-Host -Prompt "Do you want to clean $($ArchiveSessionsToClean.ID.count) Archive Sessions prior to [$Filterdate]? This can not be undone! Type (Y/N)" } 
    
    If (($confirm -CEQ "CONFIRM") -or ($strAreYouSure -eq "Y")) {
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
    Write-warning "  Archive cleaning was not CONFIRMED in Automation Policy "
    Write-warning "  $(Get-TimeStamp) You chose not to clean these archive sessions" | Out-file $ScriptLog -append
    Write-Output $Script:strLineSeparator
    #Read-Host -Prompt "Press Enter to exit"
# (If N, End the script and Exit)
    Break Script
}