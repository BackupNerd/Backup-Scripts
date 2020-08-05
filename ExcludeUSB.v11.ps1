cls

<# ----- About: ----
    # Get/ Exclude Disks with USB BUSTYPE From Backup  
    # Revision v11 - 2020-04-29
    # Author: Eric Harless, HeadBackupNerd - SolarWinds 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>

# ----- Set Script Filter Method (Yes/No) ----
    $ScriptFilterClient = "Yes"
    $ScriptFilterReg = "Yes"
# ----- End Set ScriptFilter Method (Yes/No) ----

<# ----- Behavior: ----
    # Required Run as Administrator
    # Determine if RMM Ncentral or SA Installation
    # Self Copy to correct %ProgramData% Path
    # Create Recurring SCHTASK 
    # Method 1 - Set Backup Filter for USB Disk Volume 
    # Method 2 - Add USB Disk Volume to "FilesNotToBackup" Reg Key
# -----------------------------------------------------------#>

# ----- Check for RMM & Standalone Backup Installation Type ----
    $MOB_path = "$env:ALLUSERSPROFILE\Managed Online Backup\"
    $MOB_XMLpath = Join-Path -Path $MOB_path -ChildPath "\Backup Manager\StatusReport.xml"
    $SA_path = "$env:ALLUSERSPROFILE\MXB\"
    $SA_XMLpath = Join-Path -Path $SA_path -ChildPath "\Backup Manager\StatusReport.xml"

# (Boolean vars to indicate if each exists)

    $test_MOB = Test-Path $MOB_XMLpath
    $test_SA = Test-Path $SA_XMLpath

# (If both exist, get last modified time and set path of most recent as true_path)

    If ($test_MOB -eq $True -And $test_SA -eq $True) {
	    $lm_MOB = [datetime](Get-ItemProperty -Path $MOB_XMLpath -Name LastWriteTime).lastwritetime
	    $lm_SA =  [datetime](Get-ItemProperty -Path $SA_XMLpath -Name LastWriteTime).lastwritetime
	    if ((Get-Date $lm_MOB) -gt (Get-Date $lm_SA)) {
		    $true_XMLpath = $MOB_XMLpath
            $true_path = $MOB_path
            Write-Host "Multiple Installations Found - RMM Managed Online Backup is Newest"
	    } else {
		    $true_XMLpath = $SA_XMLpath
            $true_path = $SA_path
            Write-Host "Multiple Installations Found - Standalone/N-central Backup is Newest"
	    }

# (If one exists, set it as true_path)

    } elseif ($test_SA -eq $True) {
    	$true_XMLpath = $SA_XMLpath
        $true_path = $SA_path
        Write-Host "Standalone or N-central Backup Installation Found"
    } elseif ($test_MOB -eq $True) {
    	$true_XMLpath = $MOB_XMLpath
        $true_path = $MOB_path
        Write-Host "RMM Managed Online Backup Installation Found"

# (If none exist, report & fail check)

    } else {
    	Write-Host "Backup Manager Installation Type Not Found"
    	$global:failed = 1
    }
# ----- End Check for RMM & Standalone Backup Installation Type ----

# ----- Self Copy Logic ----
    $ScriptTargetPath = $True_path
    $ScriptFull = $myInvocation.MyCommand.path
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ScriptFile = Split-Path -Leaf $MyInvocation.MyCommand.Path
    $ScriptVer = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
    $ScriptBase = $ScriptVer -replace '\..*'
    $ScriptFinal = Join-Path -Path $ScriptTargetPath -ChildPath $ScriptFile
    Test-Path -Path $ScriptFull,$ScriptTargetPath,$ScriptFinal
    If ($ScriptFull -eq $ScriptFinal) {
        'Script Already Running from Target Location'
        } Else {
        'Copying Script to Target Location'
        Copy-item -Path $ScriptFull -Destination $ScriptTargetPath -Force
        }
# ----- End Self Copy Logic ----

# ----- Windows Task Scheduler Logic ----
    $ScriptDesc = "SolarWinds MSP\Exclude USB From Backup"
    $ScriptSched = "HOURLY"
    $ScriptSchedMod = "12"

<# ----- Usage: ----
    # $ScriptSched & $ScriptSchedMod Supported Parameters
    # "MINUTE"  1 - 1439  (Not Recommend with "FilesNotToBackup" Reg Key)
    # "HOURLY"  1 - 23 
    # "DAILY"   1 - 365
    # "WEEKLY"  1 - 52
    # "MONTHLY" 1 - 12
    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/schtasks#
# ------------------------------------------------------------#>
    
    Write-Host "Creating Task to Run every $ScriptSchedMod $ScriptSched"
    SCHTASKS.EXE /Create /RU "SYSTEM" /RP /SC $ScriptSched /MO $ScriptSchedMod /TN $ScriptDesc /TR "Powershell $ScriptFinal" /RL HIGHEST /F
# ----- End Windows TaskSCheduler Logic ----

# ----- Get Disk Partion Letter for USB / Non USB Bustype ----
    $AllDisk = Get-Disk | Select-Object Number
    Update-Disk $AllDisk.Number
    Get-Disk | Sort-Object Number | Where-Object -FilterScript {$_.Bustype -Ne "USB"}
    Write-Host ""
    Write-Host "----- USB Bus Type ----"
    Write-Host ""
    Get-Disk | Sort-Object Number | Where-Object -FilterScript {$_.Bustype -Eq "USB"}
    $Disk = Get-Disk | Where-Object -FilterScript {$_.Bustype -Eq "USB"} | Select-Object Number

# (Exclude Null Partition Drive Letters)
   
    $Partition = Get-Partition -DiskNumber $Disk.Number | Where-Object {$_.DriveLetter -ne "`0"} | Select-Object @{name="DriveLetter"; expression={$_.DriveLetter+":\"}}
# ----- End Get Disk Partion Letter for USB / Non USB Bustypes ----

# ----- Set Backup Filter for Drive Letter with ClientTool ---- 
    If ($ScriptFilterClient -eq 'Yes')
    {
        Write-Host ""
        Write-Host "Adding USB Drive Letter/s to Local Backup Filter if Not Present"
        Write-Host ""
        Write-Host $Partition.DriveLetter
        Write-Host ""

        foreach ($Driveletter in $partition.DriveLetter) {
            & 'C:\Program Files\Backup Manager\clienttool.exe' control.filter.modify -add $DriveLetter
        }

<# ----- Documentation: ----     
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/command-line.htm?Highlight=control.filter.modify%20-add
# ------------------------------------------------------------#>

    }
    Remove-Variable -name ScriptFilterClient
# ----- End Set Backup Filter for Drive Letter with ClientTool ----

# ----- Add USB Drive Letter to [FilesNotToBackup] Registry Key ----
    If ($ScriptFilterReg -eq 'Yes')
    {
        Write-Host ""
        Write-Host "Adding USB Drive Letter/s to The [FilesNotToBackup] Registry Key"
        Write-Host ""
        Write-Host $Partition.DriveLetter
        Write-Host ""
        
        $USBpath = ""
        $partition.Driveletter | %{$USBPath += ($(if($USBPath){"\0"}) + $_)}
        & REG DELETE "HKLM\SYSTEM\ControlSet001\Control\BackupRestore\FilesNotToBackup" /v "ExcludeUSB" /f
        & REG ADD "HKLM\SYSTEM\ControlSet001\Control\BackupRestore\FilesNotToBackup" /v "ExcludeUSB" /t REG_MULTI_SZ /d $USBPath /f

<# ----- Documentation: ----     
    # https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/reg-add
    # https://docs.microsoft.com/en-us/windows/win32/backup/registry-keys-for-backup-and-restore#filesnottobackup 
# ------------------------------------------------------------#>
        
    } else {
        & REG DELETE "HKLM\SYSTEM\ControlSet001\Control\BackupRestore\FilesNotToBackup" /v "ExcludeUSB" /f
    }
# ----- End USB Drive Letter to [FilesNotToBackup] Registry Key ----

# ----- Restart BackupFP if Idle ----
    $BuService = get-service -name 'Backup Service Controller'
    $BackupFPstatus = & 'C:\Program Files\Backup Manager\clienttool.exe' status.get
    
    If (($ScriptFilterReg -eq 'Yes') -and ($BuService.status -eq 'Running') -and ($BackupFPstatus -eq 'Idle')) {
        & 'C:\Program Files\Backup Manager\clienttool.exe' shutdown
        }
# ----- End Restart BackupFP if Idle ----

# ----- Debug Section ----
    #Get-Variable Script*,*Path,Test* | Select-object Name,value | Format-table
    #Write-Host $ScriptFilterReg $backupFPstatus $BuService.status
    #Read-Host -Prompt "Press Enter to exit"
# ----- End Debug Section ----

# ----- Cleanup Section ----     
    #Remove-Variable -name ScriptFilterReg,BackupFPStatus,USBPath
# ----- End Cleanup Section ----    
