Clear-Host

<# ----- About: ----
    # SolarWinds Backup Custom Bandwidth Throttle  
    # Revision v05 - 2020-04-27
    # Author: Eric Harless, Head Backup Nerd - SolarWinds 
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
    # Requires Run as Adminstrator or System
    # Determine if RMM, Ncentral or Standalone Installation
    # Self Copy Script to correct %ProgramData% Path
    # Create Recurring SCHTASK
    # Create default local throttle values if txt file not found
    # Read local throttle values from exisiting txt file
    #
    # CustomBackupThrottle.txt Example below: 

    limit	OnAt	OffAt	KbUp	KbDn	Weekday
    true	08:01	16:01	4096	4096	Monday
    true	08:02	16:02	4096	4096	Tuesday
    true	08:03	16:03	4096	4096	Wednesday
    true	08:04	16:04	4096	4096	Thursday
    true	08:05	16:05	4096	4096	Friday
    false	08:06	16:06	4096	4096	Saturday
    false	08:07	16:07	4096	4096	Sunday

# -----------------------------------------------------------#>

# ----- Define Variables ----
    $Global:strLineSeparator = "---------"
# ----- End Define Variables ----

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

    Write-Host $Global:strLineSeparator 
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

# ----- End Self Copy & Logging Logic ----

# ----- Determine Day of Week Logic ----
    $Today = (Get-Date).DayOfWeek

# (Test Values)    
    #$today = "Monday"
    #$today = "Tuesday"
    #$today = "Wednesday" 
    #today = "Thursday"
    #$today = "Friday"
    #$today = "Saturday"
    #$today = "Sunday"

# ----- End Determine Day of Week Logic ----

# ----- Windows Task Scheduler Logic ----
    $ScriptDesc = "SolarWinds MSP\$scriptbase"
    $ScriptStart = "00:15"
    $ScriptSched = "HOURLY"
    $ScriptSchedMod = "3"
    
<# ----- Usage: ----
    # $ScriptSched & $ScriptSchedMod Supported Parameters
    # "MINUTE"  1 - 1439  (Not Recommend with "FilesNotToBackup" Reg Key)
    # "HOURLY"  1 - 23 
    # "DAILY"   1 - 365
    # "WEEKLY"  1 - 52
    # "MONTHLY" 1 - 12
    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/schtasks#
# ------------------------------------------------------------#>

    Write-Host "Creating Scheduled Task to Run every $ScriptSchedMod $ScriptSched"
    Write-Host $Global:strLineSeparator
    SCHTASKS.EXE /Create /RU "SYSTEM" /RP /SC $ScriptSched /MO $ScriptSchedMod /TN $ScriptDesc /TR "Powershell $ScriptFinal" /ST $ScriptStart /RL HIGHEST /F
    Write-Host $Global:strLineSeparator
# ----- End Windows TaskSCheduler Logic ----

# ----- Check for CustomBackupThrottle.txt / Create if not Exist ----

    Function CreateDefaultThrottle {

$CreateThrottle = @'
true,07:01,17:01,32,32,Monday
true,07:02,17:02,32,32,Tuesday
true,07:03,17:03,32,32,Wednesday
true,07:04,17:04,32,32,Thursday
true,07:05,17:05,32,32,Friday
true,07:06,17:06,32,32,Saturday
true,07:07,17:07,32,32,Sunday
'@ -split "`n" | % { $_.trim() }


$Global:DefaultThrottle = @()
Foreach ($line in $CreateThrottle) 
{
[array]$linedata = $line.Split(",")
$ThrottleData = New-Object PSObject
$ThrottleData | Add-Member -MemberType NoteProperty -Name "limit" -Value $LineData[0]
$ThrottleData | Add-Member -MemberType NoteProperty -Name "OnAt" -Value $LineData[1]
$ThrottleData | Add-Member -MemberType NoteProperty -Name "OffAt" -Value $LineData[2]
$ThrottleData | Add-Member -MemberType NoteProperty -Name "KbUp" -Value $LineData[3]
$ThrottleData | Add-Member -MemberType NoteProperty -Name "KbDn" -Value $LineData[4]
$ThrottleData | Add-Member -MemberType NoteProperty -Name "Weekday" -Value $LineData[5]
$Global:DefaultThrottle += $ThrottleData
}
}
    
    #test-path "$ScriptLogParent\$Scriptbase.txt"
    $TestThrottleTxt = test-path "$ScriptLogParent\$Scriptbase.txt"

    If ($TestThrottleTxt -eq "True" ) {
        Write-host "Exisiting Throttle Settings found at $ScriptLogParent"
        Write-Host $Global:strLineSeparator
        } Else {
        Write-host "Exisiting Throttle Settings not found at $ScriptLogParent"
        Write-Host $Global:strLineSeparator
        Write-host "Creating DEFAULT Throttle Settings at $ScriptLogParent\$Scriptbase.txt"
        Write-Host $Global:strLineSeparator
        CreateDefaultThrottle
        #$DefaultThrottle | format-table
        $DefaultThrottle | export-csv -delimiter "`t" -path "$ScriptLogParent\$Scriptbase.txt" -NoType | % {$_ -replace '"',''}
        }

    #Write-Host $Global:strLineSeparator
    Write-Host "Reading Throttle Settings from $ScriptLogParent\$Scriptbase.txt"
    Write-Host $Global:strLineSeparator

    $global:CustomThrottle = Import-Csv -Delimiter "`t" -Path "$ScriptLogParent\$Scriptbase.txt"
    $CustomThrottle | format-table

    Write-Host $Global:strLineSeparator
    Write-Host "$Today's throttle settings are:"
    Write-Host $Global:strLineSeparator
    $CurrentThrottle = $CustomThrottle | Where-object {$_.Weekday -eq $Today} 
    $CurrentThrottle | format-table

    [String]$Par1 = $CurrentThrottle.Limit
    [String]$Par2 = $CurrentThrottle.OnAt
    [String]$Par3 = $CurrentThrottle.OffAt
    [String]$Par4 = $CurrentThrottle.KbUp
    [String]$Par5 = $CurrentThrottle.KbDn
   
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $body = "{`"id`":`"jsonrpc`",`"jsonrpc`": `"2.0`",`"method`":`"SaveBandwidthOptions`",`"params`": {`"limitBandWidth`":$par1 ,`"turnOnAt`":`"$par2`",`"turnOffAt`":`"$par3`",`"maxUploadSpeed`":$par4,`"maxDownloadSpeed`":$par5,`"dataThroughputUnits`":`"KBits`",`"unlimitedDays`":[],`"pluginsToCancel`":[]} }"
    $response = Invoke-RestMethod 'http://localhost:5000/jsonrpcv1' -Method 'POST' -Headers $headers -Body $body
    $response | ConvertTo-Json


