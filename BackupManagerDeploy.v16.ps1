<# # ----- About: ----
    # SolarWinds Backup Universal Deployment Script
    # Revision v16 - 2020-07-23
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
    # [-Documents] [-Uid] <string>                                 ## D\L then deploy a NEW "Document" Backup device
    # [-AutoDeploy] [-Uid] <string>                                ## D\L then deploy a NEW Backup Manager as a Passphrase compatible device
    # [-AutoDeploy] [-Uid] <string> [-Takeover]                    ## D\L then take ownership of existing Backup Manager and convert to Passphrase compatible device
    # [-AutoDeploy] [-Uid] <string> [-Takeover] [-Alias <string>]  ## D\L take ownership, convert to passphrase and set a device name Alias for existing Backup Manager device
    # [-Upgrade]                                                   ## D\L then  upgrade existing Backup Manager installation to the latest Backup Manager release
    # [-Reuse]                                                     ## Copy previously stored Backup Manager Config.ini credentials to the Backup Manager installation director
    # [-Reuse] {-Restart]                                          ## Restart Backup Services after [-Reuse] command
    # [-Copy]                                                      ## Store a copy of the current Backup Manager Config.ini credentials
    # [-Copy] [-Ditto]                                             ## Store a primary and secondary copy of the current Backup Manager Config.ini credentials
    # {-Force]                                                     ## Force overwrite of existing Backup Manager installation or Config.ini credentials
    # {-Remove]                                                    ## Uninstall existing Backup Manager installation (supports [-Copy] [-Ditto] prior to removal)
    # {-Test]                                                      ## Returns configuration and settings information for the current Backup Manager installation 
    # {-Help]                                                      ## Displays Script Parameter Syntax
# -----------------------------------------------------------#>

[CmdletBinding(DefaultParameterSetName="Help")]
    Param (

        [parameter(ParameterSetName="Documents",Mandatory=$true)] [Alias("Docs")] [switch]$Documents,  ## AutoDeploy Documents backup device at location assigned by Customer UID
        
        [parameter(ParameterSetName="AutoDeploy")] [Alias("Auto")] [switch]$AutoDeploy,  ## AutoDeploy passphrase compatible backup device at location assigned by Customer UID
        
        [parameter(ParameterSetName="AutoDeploy", Mandatory=$true, Position=0)]
            [parameter(ParameterSetName="Documents", Mandatory=$true, Position=0)] [ValidateLength(36,36)] [string]$Uid,  ## Mandatory Customer UID - Found @ https://Backup.Management | Customer Management 
        
        [parameter(ParameterSetName="AutoDeploy", Position=1)] [string]$ProfileName,  ## Optional "Profile name"
        
        [parameter(ParameterSetName="AutoDeploy")] [int]$ProfileId,  ## Optional Profile ID# instead of "Profile name" 

        [parameter(ParameterSetName="AutoDeploy", Position=2)] [string]$ProductName,  ## Optional "Product name"
        
        [parameter(ParameterSetName="AutoDeploy")] [String]$Alias,  ## Optional "Device Name Alias" 
        
        [parameter(ParameterSetName="AutoDeploy")] [switch]$Takeover, ## Optional Takeover existing installation, covert to passphrase and move to location assigned by Customer UID
  
        [parameter(ParameterSetName="Redeploy", Mandatory=$false)] [Alias("Private")] [switch]$Redeploy,  ## Manual deploy with private key or redeploy device using exisiting credentials
        
        [parameter(ParameterSetName="Redeploy", Mandatory=$true, Position=0)] [string]$Devicename,  ## Mandatory for redeployment
        
        [parameter(ParameterSetName="Redeploy", Mandatory=$true, Position=1)] [ValidateLength(12,12)] [string]$Password,  ## Mandatory for redeployment
        
        [parameter(ParameterSetName="Redeploy")] [ValidateLength(36,36)] [string]$Passphrase,  ## Recommended for redeployment (can be entered via BM UI)
        
        [parameter(ParameterSetName="Redeploy")] [Alias("EncKey")] [string]$EncryptionKey,  ## Recommended for redeployment (can be entered via BM UI)
        
        [parameter(ParameterSetName="Redeploy")] [Alias("RO")] [switch]$RestoreOnly,  ## Place redeployed device in Restore Only mode
     
        [parameter(ParameterSetName="AutoDeploy")]
        [parameter(ParameterSetName="Documents")]
        [parameter(ParameterSetName="Redeploy")]
        [parameter(ParameterSetName="Copy")]
        [parameter(ParameterSetName="Reuse")]
            [parameter(ParameterSetName="Upgrade")] [Alias("F")] [switch]$Force,  ## Force overwrite of installation or config.ini 

        [parameter(ParameterSetName="Copy",Mandatory=$true)]
        [parameter(ParameterSetName="AutoDeploy")]
        [parameter(ParameterSetName="Documents")]
        [parameter(ParameterSetName="Redeploy")]
        [parameter(ParameterSetName="Upgrade")] 
            [parameter(ParameterSetName="Remove")] [switch]$Copy,  ## Store credentials from existing installation
            
        [parameter(ParameterSetName="Copy")]
        [parameter(ParameterSetName="AutoDeploy")]
        [parameter(ParameterSetName="Documents")]
        [parameter(ParameterSetName="Redeploy")]
        [parameter(ParameterSetName="Upgrade")]
            [parameter(ParameterSetName="Remove")] [switch]$Ditto,  ## Store a second set of credentials from Existing installation                
           
        [parameter(ParameterSetName="Reuse",Mandatory=$true)]
            [parameter(ParameterSetName="Upgrade")] [switch]$Reuse,  ## Reuse stored credentials 

        [parameter(ParameterSetName="Upgrade",Mandatory=$true)] [Alias("Upg")] [switch]$Upgrade,  ## Upgrade existing installation if config.ini is present

        [parameter(ParameterSetName="Reuse")] [switch]$Restart,  ## Restart Backup Process and Backup Service after Reuse of stored credentials

        [parameter(ParameterSetName="Remove",Mandatory=$true)] [Alias("Uninstall")] [switch]$Remove,  ## Remove existing installation  
             
        [parameter(ParameterSetName="Test",Mandatory=$true)] [switch]$Test,  ## Test Backup services and read configuration

        [parameter(ParameterSetName="Help",Mandatory=$false)] [switch]$Help  ## Displays script parameter syntax

    )

    clear-host
    
    # Check if script is running as Adminstrator and if not use RunAs
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if (-not $IsAdmin){
        Write-Host "The script is NOT running as Administrator, restart PowerShell as Administrator..."
                    
    }
    else{
        Write-Host "The script is running as Administrator"
    }

    $DeployType = $PSCmdlet.ParameterSetName
    $clienttool = "c:\program files\backup manager\clienttool.exe"    

    #cd (split-path -path $MyInvocation.MyCommand.source -parent)
    Write-Output ""
    Write-Output "  SolarWinds Backup Installer"
    Write-Output "  Use -Help for script parameter syntax"
    Write-Output ""
    Write-Output "  Script location:"
    Write-Output "  cd `"$PSScriptRoot`""
    Write-Output ""
 
    Function EncodeTo-Base64($InputString) {
        $BytesString = [System.Text.Encoding]::UTF8.GetBytes($InputString)
        $global:OutputBase64 = [System.Convert]::ToBase64String($BytesString)
        Write-Output $global:OutputBase64
        }
 
    Function DecodeFrom-Base64($InputBase64) {
        $BytesBase64 = [System.Convert]::FromBase64String($InputBase64)
        $global:OutputString = [System.Text.Encoding]::UTF8.GetString($BytesBase64)
        Write-Output $global:OutputString
        }

    Function Download-BackupManager {
        "  Downloading Backup Manager"
        (New-Object System.Net.WebClient).DownloadFile("http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")
    }

    Function Autodeploy-Passphrase {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"

        if (((Test-Path $BMConfig -PathType leaf) -eq $false) -or ($Force)) {
            
            if (($ProfileName) -and ($ProfileId)) {
                Write-Output ""
                Write-output "  AutoDeploy not supported with both -ProfileName and -ProfileId switches selected"
                Write-Output ""
                Break
                }
            if ($ProfileName) { $BackupProfile = "-profile-name `"$ProfileName`"" }
            if ($ProfileId) { $BackupProfile = "-profile-id `"$ProfileId`"" }
            if ($ProductName) { $BackupProduct = "-product-name `"$ProductName`"" }
            if ($Takeover) { $TakeoverParam = "-takeover" }
            if ($Alias) { $AliasParam = "-device-alias `"$(EncodeTo-Base64 $alias)`""} 

            Write-Output "  Profile param  : $backupprofile"
            Write-Output "  Product name   : $productname"
            Write-Output "  Alias name     : $Alias"
            Write-Output "  Alias param    : $AliasParam" 
            Write-Output "  Takeover       : $($Takeover.IsPresent)"
            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"
            Write-Output ""

            if ($Copy) { Copy-BackupConfig }
            
            Download-BackupManager
            Stop-BackupProcess
            Write-Output ""
            Write-Output "  Autodeploying Backup Manager instance"        
            Write-Output ""

            $process = start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-unattended-mode -silent -partner-uid $Uid $TakeoverParam $BackupProfile $BackupProduct $AliasParam" -PassThru

            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Installer" -Completed
                    Break
                }
            }
            
            Get-BackupService
            } else {
            
            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"
            Write-Output ""
            Write-Output "  Autodeploy aborted, existing Backup Manager CONFIG.INI found"
            Write-Output ""
            
            if ($Copy) { Copy-BackupConfig }  
            
            Write-Output "  Use -force switch to overwrite prior installation"
            Write-Output "  Use -copy to store prior installation credentials"
            Write-Output "  Use -force and -copy to overwrite prior installation and store prior credentials"
            Write-Output "" 
            Break
            }
    }

    Function Autodeploy-Documents {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"

        if (((Test-Path $BMConfig -PathType leaf) -eq $false) -or ($Force)) {

            Write-Output "  Profile name   : Documents"
            Write-Output "  Profile id     : 5038"
            Write-Output "  Product name   : Documents"
            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"   
            Write-Output ""

            if ($Copy) { Copy-BackupConfig }  
            Download-BackupManager
            Stop-BackupProcess
            Write-Output ""
            Write-Output "  Autodeploying a new Backup Manager (Documents) instance"       
            Write-Output ""
      
            $process = start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-unattended-mode -partner-uid $Uid -profile-name Documents" -PassThru

            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Installer" -Completed
                    Break
                }
            }

            Get-BackupService
            } else {

            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"
            Write-Output ""
            Write-Output "  Documents aborted, existing Backup Manager CONFIG.INI found"
            Write-Output ""
                        
            if ($Copy) { Copy-BackupConfig }
             
            Write-Output "  Use -force switch to overwrite prior installation"
            Write-Output "  Use -copy to store prior installation credentials"
            Write-Output "  Use -force and -copy to overwrite prior installation and store prior credentials"
            Write-Output "" 
            Break
            }
        }
     
    Function Upgrade-BackupManager {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"

        Write-Output "  Force          : $($Force.IsPresent)"
        Write-Output "  Copy Config    : $($Copy.IsPresent)"
        Write-Output "  Reuse Config   : $($reuse.IsPresent)"
        Write-Output ""
        
        if (($copy) -and ($reuse)) {
        Write-Output ""
        Write-output "  Upgrade not supported with both -copy and -reuse switches selected"
        Write-Output ""
        Break
        }
        elseif ($copy) { Copy-BackupConfig }
        elseif ($reuse) { Reuse-BackupConfig }
  
        if ((Test-Path $BMConfig -PathType leaf) -eq $true) {
        
        Download-BackupManager
        Stop-BackupProcess
        Write-Output "  Upgrading Backup Manager"
    
        $process = start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-silent -unattended " -PassThru

        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Installer" -Completed
                Break
            }
        }

        Get-BackupService
        } else {
        Write-Output ""
        Write-Output "  Upgrade aborted, no existing Backup Manager CONFIG.INI found"
        Write-Output "" 
        Break
        }
    }

    Function Redeploy-BackupManager {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"

        if (((Test-Path $BMConfig -PathType leaf) -eq $false) -or ($Force)) {            
  
            if (($Passphrase) -and ($EncryptionKey)) {
                Write-Output ""
                Write-output "  Redeploy not supported with both -Passphrase and -EncryptionKey switches selected"
                Write-Output ""
                Break
                }
 
            if ($Passphrase) { $SecurityCode = "-passphrase $Passphrase" }
            if ($EncryptionKey) { $SecurityCode = "-encryption-key `"$EncryptionKey`"" }
            if ($RestoreOnly) { $RestoreOnlyParam = "-restore-only" }

            Write-Output "  Device name    : $DeviceName"
            Write-Output "  Password       : $Password"
            Write-Output "  Passphrase     : $Passphrase"
            Write-Output "  EncryptionKey  : $EncryptionKey"
            Write-Output "  Security Param : $SecurityCode"
            Write-Output "  Restore Only   : $($RestoreOnly.IsPresent)"
            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"
            Write-Output ""

            if ($Copy) { Copy-BackupConfig } 
            Download-BackupManager
            Stop-BackupProcess
            Write-Output "  Redeploying Backup Manager with provided credentials"        
            Write-Output ""

            $process = start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-silent -user $DeviceName -password $Password $SecurityCode $RestoreOnlyParam" -PassThru

            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Installer" -Completed
                    Break
                }
            }           
           
            Get-BackupService
            Get-InitError

            } else {
            
            Write-Output "  Force Install  : $($Force.IsPresent)"
            Write-Output "  Copy Config    : $($Copy.IsPresent)"
            Write-Output ""
            Write-Output "  Redeploy aborted, existing Backup Manager CONFIG.INI found"
            Write-Output ""
            
            if ($Copy) { Copy-BackupConfig }  
            
            Write-Output "  Use -force switch to overwrite prior installation"
            Write-Output "  Use -copy to store prior installation credentials"
            Write-Output "  Use -force and -copy to overwrite prior installation and store prior credentials"
            Write-Output "" 
            Break
            }
    }

    Function Remove-BackupManager {
        
        $BackupIP="C:\Program Files\Backup Manager\BackupIP.exe"

        If ((Test-Path $BackupIP -PathType leaf) -eq $true) {
            Stop-BackupProcess
            Write-Output "  Uninstalling Backup Manager"
           
            $process = start-process -FilePath "C:\Program Files\Backup Manager\BackupIP.exe" -ArgumentList "uninstall -interactive -path `"C:\Program Files\Backup Manager`" -sure" -PassThru

            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Uninstalling"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Installer" -Completed
                    Break
                }
            }               

            Get-BackupService
            } else {
        Write-Output ""
        Write-Output "  Uninstall aborted, existing Backup Manager install not found"
        Write-Output ""
            }
        }

    Function Reuse-BackupConfig {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"
        $BMConfigCopy = "C:\programdata\MXB\config.ini.copy"

        if ((Test-Path $BMConfigCopy -PathType leaf) -eq $false) {

            Write-Output "" 
            Write-OUtput "  No Prior Copy of the Config.ini File Exists, Please reinstall using credentials from the Mangement Console"
            Write-Output "" 
            Break
            }

        elseif (((Test-Path $BMConfigCopy -PathType leaf) -eq $true) -and ($Force)) {

            Write-Output "" 
            Write-OUtput "  Reusing stored Backup Manager CONFIG.INI, copying to $BMConfig"
            Write-Output "" 
            New-Item -itemtype Directory -path "c:\Program Files\Backup Manager\" -force | Out-Null
            Copy-Item $BMConfigCopy -Destination $BMConfig

            if ($Restart) {
                Stop-BackupProcess
                Stop-BackupService
                Start-BackupService
                Get-InitError 
                }
            }

        elseif ((Test-Path $BMConfig -PathType leaf) -eq $true) {

            Write-Output ""     
            Write-Output "  Credential reuse aborted, Prior Backup Manager CONFIG.INI found"
            Write-Output "" 
            Write-Output "  Use -force switch to overwrite prior installation"
            Write-Output "  Use -restart switch to restart Backup Service"
            Write-Output "" 
            Break
            }

        elseif ((Test-Path $BMConfig -PathType leaf) -eq $false) {

            Write-Output "" 
            Write-OUtput "  Reusing stored Backup Manager CONFIG.INI, copying to $BMConfig"
            Write-Output "" 
            New-Item -itemtype Directory -path "c:\Program Files\Backup Manager\" -force | Out-Null
            Copy-Item $BMConfigCopy -Destination $BMConfig

            if ($Restart) {
                Stop-BackupProcess
                Stop-BackupService
                Start-BackupService
                Get-InitError 
                }
            }
        }   
      
    Function Copy-BackupConfig {

        $BMConfig = "C:\Program Files\Backup Manager\config.ini"
        $BMConfigCopy = "C:\programdata\MXB\config.ini.copy"

        if ((Test-Path $BMConfig -PathType leaf) -eq $false) {
            Write-Output ""
            Write-Output "  Backup Manager CONFIG.INI not found, Backup Manager may not be installed"
            Write-Output ""
            Break
            }
        elseif (((Test-Path $BMConfigCopy -PathType leaf) -eq $false) -or ($Force)) {
            Write-Output ""
            Write-Output "  Copying existing Backup Manager CONFIG.INI to $BMConfigCopy"
            Write-Output ""
            Copy-Item $BMConfig -Destination $BMConfigCopy
            if ($Ditto) { 
                Copy-Item $BMConfig -Destination "$BMConfigCopy.copy"
                Write-Output "  Storing second copy of existing Backup Manager CONFIG.INI"
                Write-Output ""
                }
            }
        elseif ((Test-Path $BMConfigCopy -PathType leaf) -eq $true) {
            Write-Output ""
            Write-Output "  Copy aborted, previously copied Backup Manager CONFIG.INI found"
            Write-Output "" 
            #Write-Output "  Use -force switch to overwrite prior installation"
            Write-Output "  Use -copy to store prior installation credentials"
            Write-Output "  Use -force and -copy to overwrite prior installation and store credentials"
            Write-Output "" 
            Break   
            } 
        }
    
    Function Test-BackupManager {

        Write-Output ""
        Write-Output "Getting Backup Manager Information"
        Write-Output ""
  
        Write-Output "Checking Backup Service Controller"
        get-service -name "Backup Service Controller"
        
        Write-Output "Checking Backup Process"
        get-process -name "BackupFP"

        Write-Output "Application status:"
            & "$clienttool" control.application-status.get
        Write-Output ""
        Write-Output "VSS Check:"        
            & "$clienttool" vss.check
        Write-Output ""
        Write-Output "Settings list:"
            & "$clienttool" control.setting.list
        Write-Output ""
        Write-Output "Selection list:"
            & "$clienttool" control.selection.list
        Write-Output ""
        Write-Output "Schedule list:"
            & "$clienttool" control.schedule.list
        Write-Output ""
        Write-Output "Filter list:"
            & "$clienttool" control.filter.list
        Write-Output ""
        Write-Output "Version info:"
            & "$clienttool" -version
        Write-Output ""
        Write-Output "Current job status:"
            & "$clienttool" control.status.get
        Write-Output ""
        Write-Output "Last error status:"
            & "$clienttool" control.initialization-error.get
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
        start-sleep -seconds 10
        $initmsg = & $clienttool control.initialization-error.get | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($($initmsg.message)) { Write-Output "  InitMsg Error  : $($initmsg.message)" }
    }

    Switch ($PSCmdlet.ParameterSetName) { 
            'Autodeploy' {
                "  Switch Type    : $DeployType"
                    Get-BackupService
                    Autodeploy-Passphrase
                }
            'Documents' { 
                "  Switch Type    : $DeployType"
                    Get-BackupService
                    Autodeploy-Documents 
                }
            'Redeploy' { 
                "  Switch Type    : $DeployType"
                   Redeploy-BackupManager 
                }
            'Upgrade' { 
                "  Switch Type    : $DeployType"
                    Get-BackupService
                    Upgrade-BackupManager 
                }
            'Remove' { 
                "  Switch Type    : $DeployType"
                   Get-BackupService
                   Remove-BackupManager 
                }
            'Reuse' { 
                "  Switch Type    : $DeployType"
                   Get-BackupService
                   Reuse-BackupConfig 
                }
            'Copy' { 
                "  Switch Type    : $DeployType"
                   Copy-BackupConfig 
                }
            'Test' { 
                "  Switch Type    : $DeployType"
                   Test-BackupManager 
                }
            'Help' { 
                "  Switch Type    : $DeployType"
                   Get-BackupService
                   Write-Output ""
                   #Get-Command "  $($MyInvocation.MyCommand.source)" -Syntax
                   Get-Command $PSCommandPath -Syntax  
                }
        }


    