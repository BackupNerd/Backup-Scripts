  
[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][ValidateLength (12,12)] [string]$installationkey   ## Replacement Installation Key 
    )

    clear-host

    #$scriptpath = $MyInvocation.MyCommand.Path
    #$dir = Split-Path $scriptpath
    #Push-Location $dir

    Function Update-Config {
        if (Get-Module -ListAvailable -Name PsIni) {
            Write-Host "  Module PsIni Already Installed"
        } 
        else {
            try {
                Install-Module -Name PsIni -Confirm:$False -Force      ## https://www.powershellgallery.com/packages/PsIni
            }
            catch [Exception] {
                $_.message 
                break
            }
        }

        $config = "C:\Program Files\Backup Manager\config.ini" 
        $General = @{}
        $General.Add("Password","$installationkey")
    
        Get-IniContent $config | Set-IniContent -Sections 'General' -NameValuePairs  $General | Out-IniFile $config  -Pretty -Force -Encoding ASCII
    }

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
        $clienttool = "c:\program files\backup manager\clienttool.exe"    
        start-sleep -seconds 10
        $initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
    }

    Function Get-AppStatus {
        $clienttool = "c:\program files\backup manager\clienttool.exe"    
        start-sleep -seconds 30
        & "$clienttool" control.application-status.get
        & "$clienttool" control.status.get  
    }

    Update-Config
    Stop-BackupProcess
    Stop-BackupService
    Start-BackupService
    Get-BackupService
    Get-InitError 
    Get-AppStatus
