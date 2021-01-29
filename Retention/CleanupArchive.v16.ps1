<# ----- About: ----
    # SolarWinds Backup Cleanup Archive
    # Revision v16 - 2020-08-31
    # Author: Eric Harless, HeadBackupNerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Script Parameters:
    #
    # [-Months ##] <Int>                                         ## Clean Archives older than ## number of months (Default is 84)
    # [-Force] <switch>                                          ## Do not prompt for confirmation before Archive Cleaning
    # [CheckArchive] <switch>                                    ## Check backup client for an active Archive rule
    # [AddArchive] <switch>                                      ## Add a Monthly Archive rule if no active Archive rules exist 
    # 
    # RMM / Ncentral and Standalone Backup compatible
    # Self Copy Script to Backup %ProgramData% Path
    # Get/ Log Current Device/Archive Settings using ClientTool.exe
    # 
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

        [ValidateRange(12,84)] [int]$Months=84, ## Clean Archives older than X number of months
        [switch]$Force,                         ## Do not prompt for confirmation before Archive Cleaning
        [switch]$CheckArchive,                  ## Check backup client for an Active Archive rule
        [switch]$AddArchive                     ## Add a Monthly Archive rule if no active Archive rules exist 
               
    )

    Clear-Host

# ----- Define Variables ----
    $Global:strLineSeparator = "  ---------"
    $ArchiveValue = ($months * -1)
    $filterDate = (Get-Date).Addmonths($ArchiveValue)
# ----- End Define Variables ----

    Function Get-LocalTime($UTCTime) {
        $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
        $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
        $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
        Return $LocalTime
    } ## Get Local Time for comparison vs UTC Time
    
    Function Get-TimeStamp {
        return "[{0:M/d/yyyy} {0:HH:mm:ss}]" -f (Get-Date)
    }  ## Get Timestamp for logging
    
    Function CheckInstallationType {

# ----- Check for RMM & Standalone Backup Installation Type ----
    $MOB_path = "$env:ALLUSERSPROFILE\Managed Online Backup\"
    $MOB_XMLpath = Join-Path -Path $MOB_path -ChildPath "\Backup Manager\StatusReport.xml"
    $MOB_clientpath = "$env:PROGRAMFILES\Managed Online Backup\"
    $SA_path = "$env:ALLUSERSPROFILE\MXB\"
    $SA_XMLpath = Join-Path -Path $SA_path -ChildPath "\Backup Manager\StatusReport.xml"
    $SA_clientpath = "$env:PROGRAMFILES\Backup Manager\"

# ----- Boolean vars to indicate if each exists

    $test_MOB = Test-Path $MOB_XMLpath
    $test_SA = Test-Path $SA_XMLpath

# ----- If both exist, get last modified time and set path of most recent as true_path

    If ($test_MOB -eq $True -And $test_SA -eq $True) {
	    $lm_MOB = [datetime](Get-ItemProperty -Path $MOB_XMLpath -Name LastWriteTime).lastwritetime
	    $lm_SA =  [datetime](Get-ItemProperty -Path $SA_XMLpath -Name LastWriteTime).lastwritetime
	    if ((Get-Date $lm_MOB) -gt (Get-Date $lm_SA)) {
		    $global:true_XMLpath = $MOB_XMLpath
            $global:true_path = $MOB_path
            $global:true_clientpath = $MOB_clientpath
            Write-Output $Global:strLineSeparator
            Write-Output "  Multiple Installations Found - RMM Managed Online Backup is Newest"
	    } else {
		    $global:true_XMLpath = $SA_XMLpath
            $global:true_path = $SA_path
            $global:true_clientpath = $SA_clientpath
            Write-Output $Global:strLineSeparator
            Write-Output "  Multiple Installations Found - Standalone/N-central Backup is Newest"
	    }

# ----- If one exists, set it as true_path

    } elseif ($test_SA -eq $True) {
    	$global:true_XMLpath = $SA_XMLpath
        $global:true_path = $SA_path
        $global:true_clientpath = $SA_clientpath
        Write-Output $Global:strLineSeparator
        Write-Output "  Standalone or N-central Backup Installation Found"
    } elseif ($test_MOB -eq $True) {
    	$global:true_XMLpath = $MOB_XMLpath
        $global:true_path = $MOB_path
        $global:true_clientpath = $MOB_clientpath
        Write-Output $Global:strLineSeparator
        Write-Output "  RMM Managed Online Backup Installation Found"

# ----- If none exist, report & fail check

    } else {
        Write-Output $Global:strLineSeparator
    	Write-Output "  Backup Manager Installation Type Not Found"
    	$global:failed = 1
    }
# ----- End Check for RMM & Standalone Backup Installation Type ----


}  ## Check for RMM / Ncentral / Standalone Backup Installation Type

    CheckInstallationType

# ----- Self Copy & Logging Logic ----

    $ScriptFull = $myInvocation.MyCommand.path
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ScriptFile = Split-Path -Leaf $MyInvocation.MyCommand.Path
    $ScriptVer = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
    $ScriptBase = $ScriptVer -replace '\..*'
    $ScriptFinal = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath $ScriptFile
    $ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.log"
    $ScriptLogParent = Split-path -path $ScriptLog
    md -Force $ScriptLogParent

    $SetCompressed = Invoke-WmiMethod -Path "Win32_Directory.Name='$ScriptLogParent'" -Name compress
    If (($SetCompressed.returnvalue) -eq 0) { "Items successfully compressed" } else { "Something went wrong!" }
  
    If ($ScriptFull -eq $ScriptFinal) {
        Write-Output $Global:strLineSeparator 
        Write-Output 'Script Already Running from Target Location'
        Write-Output $Global:strLineSeparator
        } Else {
        Write-Output $Global:strLineSeparator
        Write-Output 'Copying Script to Target Location'
        Write-Output $Global:strLineSeparator
        Copy-item -Path $ScriptFull -Destination $ScriptLogParent -Force
        }

# ----- End Self Copy Logic ----

    Function GetArchiveSettings {

# ----- Clienttool Get Archive Settings From Local Client ----

        & "$true_clientpath\ClientTool.exe" -machine-readable control.setting.list > "$ScriptLogParent\DeviceSettings.tsv"
        $global:DeviceSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\DeviceSettings.tsv"
        $global:DeviceName = $DeviceSettings | Where-Object { $_.Name -eq "Device" } | Select-Object Value 
        $global:ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.$($devicename.value).log"
          
    if (($CheckArchive) -or ($AddARchive)) {
    
        & "$true_clientpath\ClientTool.exe" -machine-readable control.archiving.list > "$ScriptLogParent\ArchiveSettings.tsv"
        $ArchiveSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\ArchiveSettings.tsv"
    
        Write-Output "  Getting Archive Settings from Clienttool for $($ENV:COMPUTERNAME) | $($DeviceName.value)"
        Write-Output $Global:strLineSeparator
        $ArchiveSettings | format-table -autosize
   
        Write-Output "  $(Get-TimeStamp) Getting current archive settings for $($ENV:COMPUTERNAME) | $($DeviceName.value)" | Out-file $ScriptLog -append
        $ArchiveSettings | format-table | Out-file $ScriptLog -append

        $global:ArchiveCheck = $ArchiveSettings | Where-object{ ($_.ACTV -eq "yes")} | Select-Object -Property ID,ACTV,NAME 
        }

# (Check for Archive Rule / Create if missing)
    If (($CheckArchive) -and ($ArchiveCheck.ACTV) -eq "yes") {
        Write-Output $Global:strLineSeparator
        Write-Output 'Active Archive Schedule Found'
        Write-Output $Global:strLineSeparator
        #$Archivecheck | Where-object{ $_.name -eq "#Monthly - Archive#"} | Select-Object -Property ID,ACTV,NAME | format-table
        $Archivecheck | Select-Object -Property ID,ACTV,NAME | format-table
        
    } 
    
    if (($AddArchive) -and ($ArchiveSettings.ACTV) -notcontains "yes") {
        Write-Output 'Creating Monthly Archive Schedule/s'
        & "$true_clientpath\ClientTool.exe" control.archiving.add -name "#Monthly - Archive#" -days-of-month Last
    }

# ----- End Clienttool Get Archive Settings From Local Client ----    
    }  ## Get current Archive settings

    GetArchiveSettings

    Function ClienttoolGetArchiveSessions {

    & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > "$ScriptLogParent\AllSessions.tsv"
    $Sessions = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\AllSessions.tsv"

# Clienttool Sessions - All

    #Write-Output "  All Session History from Clienttool.exe"    
    #$Sessions | format-table

# Clienttool Sessions - Filtered

    $Global:CleanSessions = $Sessions | 
    Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and
                  ($_.TyPE -eq "Backup") -and 
                  ([datetime]$_.START -lt $filterDate) } | 
              Select-Object -Property START,DSRC,STATE,TYPE,FLAGS,SELS | Sort-Object START 

    $Global:CleanSessions| ForEach-Object {$_.START = [datetime]$_.START}

    #Write-Output "  All Archive Session History from Clienttool.exe older than $filterdate [datetime]"
    #$Global:CleanSessions |Select-Object -Property DSRC,START,STATE,TYPE,FLAGS,SELS | Sort-Object START | format-table
    }  ## Get Historic Archive sessions

    ClienttoolGetArchiveSessions

    Function XMLLookupArchiveSessionIds {

# Read SessionReport.XML for all Historical Archives
   $SessionReport = Join-path -path(split-path $true_XMLpath) -ChildPath "SessionReport.xml"
    [xml]$XmlDocument = Get-Content -Path $SessionReport

    #Write-Output "  Non archive sessions from SessionReport.xml (includes Cleaned Sessions)"
    #$XmlDocument.SessionStatistics.Session |
        Where-Object{ ($_.Archived -eq "0") } | 
        Select-Object -Property Id,Plugin, StartTimeUTC,Status,Type,Archived | Sort-Object StartTime | format-table

    #Write-Output "  All archive sessions from SessionReport.xml older than $Filterdate (includes Cleaned Sessions)"
    #$XmlDocument.SessionStatistics.Session |
        Where-Object{ ($_.Archived -eq "1") -and
                  ([datetime]$_.StartTimeUTC -lt $filterDate) } | 
              Select-Object -Property Id,Plugin, StartTimeUTC,Status,Type,Archived | Sort-Object StartTime | format-table

    $global:CleanArchive = $XmlDocument.SessionStatistics.Session |
        Where-Object{ ($_.Archived -eq "1") } | 
              Select-Object -Property @{Name="StartTime";Expression={$_.StartTimeUTC}},Id,Plugin,Status,Type,Archived | Sort-Object StartTime 
    
    $global:CleanArchive | ForEach-Object {$_.StartTime = get-localtime ([datetime]$_.StartTime)} |
         Where-Object{ ([datetime]$_.StartTime -lt $filterDate) } | 
              Select-Object -Property Id,Plugin,StartTime,Status,Type,Archived | Sort-Object StartTime |format-table
   
    #$global:CleanArchive |
         Where-Object{ ([datetime]$_.StartTime -lt $filterDate) } | 
              Select-Object -Property Id,Plugin,StartTime,Status,Type,Archived | Sort-Object StartTime |format-table

    }  ## Get Historic Archive session IDs

    XMLLookupArchiveSessionIds

    $ArchiveSessionsToClean = Foreach ($row in $CleanSessions) {
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
    Write-Output $Global:strLineSeparator
    Write-Output "  Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"

    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags | Sort-Object StartTime | format-table

    Write-Output "  $(Get-TimeStamp) Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate" | Out-file $ScriptLog -append
    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags| Sort-Object StartTime | format-table | Out-file $ScriptLog -append

    [int32[]]$global:ArchiveSessionIDsToClean = $ArchiveSessionsToClean.id

    Write-Output $Global:strLineSeparator
    Write-Output "  Sessions to Clean" $ArchiveSessionIDsToClean
    Write-Output $Global:strLineSeparator

    if ($($ArchiveSessionsToClean.ID.count) -eq 0) { break }

# (Check for Confirmation)
    If ($force -ne $true) { $strAreYouSure = Read-Host -Prompt "Are you sure you want to Clean Archives prior to [$Filterdate] (Y/N)" } 
    
    If (($force) -or ($strAreYouSure -eq "Y")) {
        Write-Output $Global:strLineSeparator
        Write-Output "  Performing Archive Cleaning"
        Write-Output $Global:strLineSeparator
# (If Y, then Clean Archives)

# (URL to JSON-RPC API)
    $urlJSON = 'http://localhost:5000/jsonrpcv1'

# (Function to call the JSON-RPC web request)
    Function CallJSON($url,$object) {
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
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

# (Create the JSON object to call the CleanupArchiving function)
    $objCleanupArchiving = (New-Object PSObject | 
    Add-Member -PassThru NoteProperty jsonrpc ‘2.0’ |
    Add-Member -PassThru NoteProperty method ‘CleanupArchiving’ |
    Add-Member -PassThru NoteProperty params @{
                                               sessionsToCleanup=$ArchiveSessionIDsToClean 
                                              }|
    Add-Member -PassThru NoteProperty id ‘jsonrpc’)| ConvertTo-Json -Depth 4

# (Check JSON Syntax - for Debugging)
    #Write-Output $objCleanupArchiving

# (Call the JSON Web Request Function to get the CleanupArchiving Object)
    $CleanupArchivingSession = CallJSON $urlJSON $objCleanupArchiving
    
# (pause and wait for return)
    Start-Sleep -Milliseconds 200

# (Get Result Status of Cleanup Archiving)
    $CleanupArchivingErrorCode = $CleanupArchivingSession.error.code
    $CleanupArchivingErrorMsg = $CleanupArchivingSession.error.message

# (Check for Errors with Cleanup Archiving)
    If ($CleanupArchivingErrorCode) {
        Write-Output $Global:strLineSeparator
        Write-Output "  Get Cleanup Archiving Error Code: " $CleanupArchivingErrorCode
        Write-Output "  Get Cleanup Archiving Error Message: " $CleanupArchivingErrorMsg
        
        If ($ArchiveSessionIDsToClean -eq $null) { 
  
        Write-Output "  Get Cleanup Archiving Description:  Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
        }
        Write-Output $Global:strLineSeparator
 # (Exit Script if there is a problem)
        Break Script
    } else {

    Write-Output "  Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
    Write-Output "  $(Get-TimeStamp) Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"| Out-file $ScriptLog -append
    Write-Output "" | Out-file $ScriptLog -append
    }
    } Else {
    Write-Output $Global:strLineSeparator
    Write-Output "  You chose No, exiting script."
    Write-Output "  $(Get-TimeStamp) You chose No to Clean Archive Sessions" | Out-file $ScriptLog -append
    Write-Output $Global:strLineSeparator
    #Read-Host -Prompt "Press Enter to exit"
# (If N, End the script and Exit)
    Break Script
}
