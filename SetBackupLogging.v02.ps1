<# # ----- About: ----
    # SolarWinds Backup Set Logging Level
    # Revision v02 - 2020-08-20
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
    # [-LogLevel]                                                  ## Sets logging level in Config.ini to Log
    # [-LogLevel] {-Restart]                                       ## Sets logging level in Config.ini to Log and Restart the Backup Service 
    # [-ErrorLevel]                                                ## Sets logging level in Config.ini to Error
    # [-WarningLevel]                                              ## Sets logging level in Config.ini to Warning (Default)
    # [-DebugLevel] [-Restart]                                     ## Sets logging level in Config.ini to Debug and Restart the Backup Service 
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/logging.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/debug-logs.htm

# -----------------------------------------------------------#>

[CmdletBinding(DefaultParameterSetName="WarningLevel")]    
    Param (

        [parameter(ParameterSetName="LogLevel",Mandatory=$false)] [switch]$LogLevel,
        [parameter(ParameterSetName="ErrorLevel",Mandatory=$false)] [switch]$ErrorLevel,
        [parameter(ParameterSetName="WarningLevel",Mandatory=$false)] [switch]$WarningLevel,        
        [parameter(ParameterSetName="DebugLevel",Mandatory=$false)] [switch]$DebugLevel,

        [parameter(ParameterSetName="LogLevel",Mandatory=$false)]
        [parameter(ParameterSetName="ErrorLevel",Mandatory=$false)]
        [parameter(ParameterSetName="WarningLevel",Mandatory=$false)] 
        [parameter(ParameterSetName="DebugLevel",Mandatory=$false)] [switch]$Restart  ## Restart Backup Process and Backup Service after update
    )
    
    clear-host
    
    Function Stop-BackupProcess {
        stop-process -name "BackupFP" -Force -ErrorAction SilentlyContinue
    }

    Function Stop-BackupService {
        $BackupService = get-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
        
        if ($BackupService.Status -eq "Stopped") {
        Write-Output "  Backup Service : $($BackupService.status)"
        }else{
        Write-Output "  Backup Service : $($BackupService.status)"
        stop-service -name "Backup Service Controller" -force -ErrorAction SilentlyContinue
        Get-BackupService
        }
    }

    Function Start-BackupService {
        $BackupService = get-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
        
        if ($BackupService.Status -eq "Running") {
        Write-Output "  Backup Service : $($BackupService.status)"
        Get-InitError
        }else{
        Write-Output "  Backup Service : $($BackupService.status)"
        start-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
        Get-BackupService
        Get-InitError
        }
    }

    Function Get-BackupService {
        $BackupService = get-service -name "Backup Service Controller" -ErrorAction SilentlyContinue
        
        if ($backupservice.status -eq "Stopped") {
        Write-Output "  Backup Service : $($BackupService.status)"
        }
        elseif ($backupservice.status -eq "Running") {
            Write-Output "  Backup Service : $($BackupService.status)"
            #start-sleep -seconds 10
            #$initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
            #if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
        }
        else{
        Write-Output "  Backup Service : Not Present"
        }
    }

    Function Get-InitError {
        start-sleep -seconds 10
        $initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
    }


    # Check if script is running as Adminstrator
    
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if (-not $IsAdmin){
        Write-Host "The script is NOT running as Administrator, restart PowerShell as Administrator..."
                    
    }
    else{
        Write-Host "  The script is running as Administrator"
        Write-Host "" 
    }

    $ParamType = $PSCmdlet.ParameterSetName
    

    # define file for input/output

    $config = 'C:\Program Files\Backup Manager\config.ini'

    Write-Output "  Script parameter = $ParamType"

    if ($ParamType -eq "LogLevel") {$Level = "Log"}   
    if ($ParamType -eq "ErrorLevel") {$Level = "Error"}  
    if ($ParamType -eq "WarningLevel") {$Level = "Warning"}  
    if ($ParamType -eq "DebugLevel") {$Level = "Debug"}  
       
    $loggingheader = "\[Logging\]"
    $logginglevel = "LoggingLevel=*"
    $newlogginglevel = "[Logging]`r`nLoggingLevel=$level"

    # use temporary variable for enabling config overwrite

    $tmp = get-content $config | select-string -pattern ^$loggingheader -notmatch
    $tmp1 = $tmp | select-string -pattern ^$logginglevel -notmatch
    $tmp1 | Set-Content $config
    start-sleep -seconds 2
        
    # add new content at the end of the file

    $newlogginglevel | Out-File $config -Append ascii

    if ($Restart) {
        Stop-BackupProcess
        Stop-BackupService
        Start-BackupService
    }  
        
    Write-Output ""
    Write-Output "  Script location:"
    Write-Output "  cd `"$PSScriptRoot`""
    Write-Output ""
    Write-Output "  Script syntax:"
    Write-Output ""
    Get-Command $PSCommandPath -Syntax