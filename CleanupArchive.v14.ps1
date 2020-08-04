Clear-Host

<# ----- About: ----
    # SolarWinds Backup Cleanup Archive
    # Revision v14- 2020-05-06
    # Author: Eric Harless, HeadBackupNerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>
    $ArchiveValue = "-84"

    $filterDate = (Get-Date).Addmonths($ArchiveValue)
    #$filterDate = (Get-Date).Addyears(-7)

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # (VARIABLE) Set FilterDate
    # (PENDING) Script Parameters
    # Determine if RMM, Ncentral or Standalone Installation
    # Self Copy Script to correct %ProgramData% Path
    # (PENDING)Create Recurring SCHTASK
    # (PENDING)AMP FILE 
    # Get/ Log Current Device/Archive Settings from ClientTool
    # Check/ Add Default Archive Rule
    # Get Archive Session Times from ClientTool
    # Get Archive Session IDs from SessionReport.xml
    # Normalize DateTime/Match Archive Session Time/ID
    # List ArchiveSessionIDsToClean
    # Prompt for Confirmation to Clean 
    # Pass List to JSON CleanupAchiving Function
    # Log actions 
# -----------------------------------------------------------#>

# ----- Define Variables ----
    $Global:strLineSeparator = "---------"
# ----- End Define Variables ----

    Function Get-LocalTime($UTCTime) {
        $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
        $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
        $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
        Return $LocalTime
    }
    
    Function Get-TimeStamp {
        return "[{0:M/d/yyyy} {0:HH:mm:ss}]" -f (Get-Date)
    }
    
    Function CheckInstallationType {

# ----- Check for RMM & Standalone Backup Installation Type ----
    $MOB_path = "$env:ALLUSERSPROFILE\Managed Online Backup\"
    $MOB_XMLpath = Join-Path -Path $MOB_path -ChildPath "\Backup Manager\StatusReport.xml"
    $MOB_clientpath = "$env:PROGRAMFILES\Managed Online Backup\"
    $SA_path = "$env:ALLUSERSPROFILE\MXB\"
    $SA_XMLpath = Join-Path -Path $SA_path -ChildPath "\Backup Manager\StatusReport.xml"
    $SA_clientpath = "$env:PROGRAMFILES\Backup Manager\"

# (Boolean vars to indicate if each exists)

    $test_MOB = Test-Path $MOB_XMLpath
    $test_SA = Test-Path $SA_XMLpath

# (If both exist, get last modified time and set path of most recent as true_path)

    If ($test_MOB -eq $True -And $test_SA -eq $True) {
	    $lm_MOB = [datetime](Get-ItemProperty -Path $MOB_XMLpath -Name LastWriteTime).lastwritetime
	    $lm_SA =  [datetime](Get-ItemProperty -Path $SA_XMLpath -Name LastWriteTime).lastwritetime
	    if ((Get-Date $lm_MOB) -gt (Get-Date $lm_SA)) {
		    $global:true_XMLpath = $MOB_XMLpath
            $global:true_path = $MOB_path
            $global:true_clientpath = $MOB_clientpath
            Write-Host $Global:strLineSeparator
            Write-Host "Multiple Installations Found - RMM Managed Online Backup is Newest"
	    } else {
		    $global:true_XMLpath = $SA_XMLpath
            $global:true_path = $SA_path
            $global:true_clientpath = $SA_clientpath
            Write-Host $Global:strLineSeparator
            Write-Host "Multiple Installations Found - Standalone/N-central Backup is Newest"
	    }

# (If one exists, set it as true_path)

    } elseif ($test_SA -eq $True) {
    	$global:true_XMLpath = $SA_XMLpath
        $global:true_path = $SA_path
        $global:true_clientpath = $SA_clientpath
        Write-Host $Global:strLineSeparator
        Write-Host "Standalone or N-central Backup Installation Found"
    } elseif ($test_MOB -eq $True) {
    	$global:true_XMLpath = $MOB_XMLpath
        $global:true_path = $MOB_path
        $global:true_clientpath = $MOB_clientpath
        Write-Host $Global:strLineSeparator
        Write-Host "RMM Managed Online Backup Installation Found"

# (If none exist, report & fail check)

    } else {
        Write-Host $Global:strLineSeparator
    	Write-Host "Backup Manager Installation Type Not Found"
    	$global:failed = 1
    }
# ----- End Check for RMM & Standalone Backup Installation Type ----


}

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

    #Test-Path -Path $ScriptFull,$ScriptFinal,$scriptLog
    
    If ($ScriptFull -eq $ScriptFinal) {
        Write-Host $Global:strLineSeparator 
        Write-host 'Script Already Running from Target Location'
        Write-Host $Global:strLineSeparator
        } Else {
        Write-Host $Global:strLineSeparator
        Write-host 'Copying Script to Target Location'
        Write-Host $Global:strLineSeparator
        Copy-item -Path $ScriptFull -Destination $ScriptLogParent -Force
        }

# ----- End Self Copy Logic ----


    Function GetArchiveSettings {

# ----- Clienttool Get Archive Settings From Local Client ----

    & "$true_clientpath\ClientTool.exe" -machine-readable control.setting.list > "$ScriptLogParent\DeviceSettings.tsv"
    $global:DeviceSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\DeviceSettings.tsv"
    $global:DeviceName = $DeviceSettings | Where-Object { $_.Name -eq "Device" } | Select-Object Value 
    $global:ScriptLog = Join-Path -Path $True_path -ChildPath $ScriptBase | Join-Path -ChildPath "$ScriptBase.$($devicename.value).log"
    
    & "$true_clientpath\ClientTool.exe" -machine-readable control.archiving.list > "$ScriptLogParent\ArchiveSettings.tsv"
    $ArchiveSettings = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\ArchiveSettings.tsv"
    
    Write-Host "Getting Archive Settings from Clienttool for $($ENV:COMPUTERNAME) | $($DeviceName.value)"
    Write-Host $Global:strLineSeparator
    $ArchiveSettings | format-table -autosize
   
    Write-Output "$(Get-TimeStamp) Getting current archive settings for $($ENV:COMPUTERNAME) | $($DeviceName.value)" | Out-file $ScriptLog -append
    $ArchiveSettings | format-table | Out-file $ScriptLog -append

    $ArchiveCheck = $ArchiveSettings | Where-object{ ($_.NAME -eq "#Monthly - Archive#")} | Select-Object -Property ID,ACTV,NAME 


# (Check for Archive Rule / Create if missing)
    If (($ArchiveCheck.name -eq "#Monthly - Archive#") -and ($ArchiveCheck.ACTV -eq "yes")) {
        Write-Host $Global:strLineSeparator
        write-host 'Exisiting Monthly Archive Schedule Found'
        Write-Host $Global:strLineSeparator
        #$Archivecheck | Where-object{ $_.name -eq "#Monthly - Archive#"} | Select-Object -Property ID,ACTV,NAME | format-table
        $Archivecheck | Select-Object -Property ID,ACTV,NAME | format-table
        } else {
        write-host 'Creating Monthly Archive Schedule'
        & "$true_clientpath\ClientTool.exe" control.archiving.add -name "#Monthly - Archive#" -days-of-month Last
        }

# ----- End Clienttool Get Archive Settings From Local Client ----    
    }

    GetArchiveSettings

    Function ClienttoolGetArchiveSessions {

    & "C:\Program Files\Backup Manager\clienttool.exe" -machine-readable control.session.list > "$ScriptLogParent\AllSessions.tsv"
    $Sessions = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\AllSessions.tsv"

# Clienttool Sessions - All

    #Write-host "All Session History from Clienttool.exe"    
    #$Sessions | format-table

# Clienttool Sessions - Filtered

    $Global:CleanSessions = $Sessions | 
    Where-Object{ (($_.FLAGS -eq'A---?') -or ($_.FLAGS -eq 'A---?---?')) -and
                  ($_.TyPE -eq "Backup") -and 
                  ([datetime]$_.START -lt $filterDate) } | 
              Select-Object -Property START,DSRC,STATE,TYPE,FLAGS,SELS | Sort-Object START 

    $Global:CleanSessions| ForEach-Object {$_.START = [datetime]$_.START}

    #Write-host "All Archive Session History from Clienttool.exe older than $filterdate [datetime]"
    #$Global:CleanSessions |Select-Object -Property DSRC,START,STATE,TYPE,FLAGS,SELS | Sort-Object START | format-table
    }

    ClienttoolGetArchiveSessions

    Function XMLLookupArchiveSessionIds {

# Read SessionReport.XML for all Historical Archives
   $SessionReport = Join-path -path(split-path $true_XMLpath) -ChildPath "SessionReport.xml"
    [xml]$XmlDocument = Get-Content -Path $SessionReport

    #Write-Host "Non archive sessions from SessionReport.xml (includes Cleaned Sessions)"
    #$XmlDocument.SessionStatistics.Session |
        Where-Object{ ($_.Archived -eq "0") } | 
        Select-Object -Property Id,Plugin, StartTimeUTC,Status,Type,Archived | Sort-Object StartTime | format-table

    #Write-Host "All archive sessions from SessionReport.xml older than $Filterdate (includes Cleaned Sessions)"
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

    }

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
    Write-Host $Global:strLineSeparator
    Write-Host "Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags | Sort-Object StartTime | format-table

    Write-Output "$(Get-TimeStamp) Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate" | Out-file $ScriptLog -append
    $ArchiveSessionsToClean | Select-Object -Property ID,DataSource,StartTime,State,Type,SelectedSize,Flags| Sort-Object StartTime | format-table | Out-file $ScriptLog -append

    [int32[]]$global:ArchiveSessionIDsToClean = $ArchiveSessionsToClean.id

    Write-Host $Global:strLineSeparator
    Write-Host "Sessions to Clean" $ArchiveSessionIDsToClean
    Write-Host $Global:strLineSeparator

# (Check for Confirmation)
    $strAreYouSure = Read-Host -Prompt "Are you sure you want to Clean Archives prior to [$Filterdate] (Y/N)"
    If ($strAreYouSure -eq "Y") {
        Write-Host $Global:strLineSeparator
        Write-Host "Performing Archive Cleaning"
        Write-Host $Global:strLineSeparator
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
    #Write-host $objCleanupArchiving

# (Call the JSON Web Request Function to get the CleanupArchiving Object)
    $CleanupArchivingSession = CallJSON $urlJSON $objCleanupArchiving
    
# (pause and wait for return)
    Start-Sleep -Milliseconds 300

# (Get Result Status of Cleanup Archiving)
    $CleanupArchivingErrorCode = $CleanupArchivingSession.error.code
    $CleanupArchivingErrorMsg = $CleanupArchivingSession.error.message

# (Check for Errors with Cleanup Archiving)
    If ($CleanupArchivingErrorCode) {
        Write-Host $Global:strLineSeparator
        Write-Host "Get Cleanup Archiving Error Code: " $CleanupArchivingErrorCode
        Write-Host "Get Cleanup Archiving Error Message: " $CleanupArchivingErrorMsg
        If ($ArchiveSessionIDsToClean -eq $null) { 
        Write-host "Get Cleanup Archiving Description:  Found $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
        }
        Write-Host $Global:strLineSeparator
 # (Exit Script if there is a problem)
        Break Script
    } else {

    Write-Host "Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"
    Write-Output "$(Get-TimeStamp) Removed $($ArchiveSessionsToClean.ID.count) archive session(s) older than $Filterdate"| Out-file $ScriptLog -append
    Write-Output "" | Out-file $ScriptLog -append
    }
    } Else {
    Write-Host $Global:strLineSeparator
    Write-Host "You chose No, exiting script."
    Write-Output "$(Get-TimeStamp) You chose No to Clean Archive Sessions" | Out-file $ScriptLog -append
    Write-Host $Global:strLineSeparator
    #Read-Host -Prompt "Press Enter to exit"
# (If N, End the script and Exit)
    Break Script
}