<#
.SYNOPSIS
  Pre migration Prep script to allow a move from N-able N-central integrated backup to N-able | Cove Data Protection
.DESCRIPTION
  This script checks N-central integration status, Caches the existing Backup Manager configuration file, upgrades
  to the latest Backup Manager client and logs the success of the latest Prep attempt. 
.INPUTS
  None
.OUTPUTS
  Script results are output to the console and the local Backup Manager's Device Alias Name is updated to display
  the last Migration Prep timestamp in the https://backup.management console
.NOTES
  Revision:         v24.07.23
  Purpose/Change:   Updates to error handling and added proof of completion
  Author:           Eric Harless, Head Backup Nerd - N-able 
  Email:            eric.harless@n-able.com
  Twitter:          @backup_nerd
  Reddit:           https://www.reddit.com/r/Nable/
.NOTES 
  Sample scripts are not supported under any N-able support program or service.
  The sample scripts are provided AS IS without warranty of any kind.
  N-able expressly disclaims all implied warranties including, warranties
  of merchantability or of fitness for a particular purpose. 
  In no event shall N-able or any other party be liable for damages arising
  out of the use of or inability to use the sample scripts.
.EXAMPLE
  No parameters
.LINK
https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/command-line.htm?Highlight=clienttool

#>

Function Get-BackupIntegrationStatus {

    $IntegratedBackupXML = "C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\MSPBackupManagerConfig.xml"     ## Ncentral XML for Integrated Backup
    $StatusReportxml     = "C:\ProgramData\MXB\Backup Manager\StatusReport.xml" 
    $Script:PartnerName  = ([Xml] (get-content $StatusReportxml)).SelectSingleNode("//PartnerName")."#text"
    $Script:BackupIP = "c:\Program Files\Backup Manager\BackupIP.exe"
    
    Write-Output "`nChecking N-able Backup integration status"
    if ((test-path -PathType Leaf -Path $IntegratedBackupXML) -eq $true) {

        [xml]$BackupAgent = get-content -path $IntegratedBackupXML 

        $Script:IsIntegrated = $BackupAgent.MaxBackupManagerConfig.enabled
        $Script:IsInstalled  = $BackupAgent.MaxBackupManagerConfig.installationStatus
        $Script:CUID         = $BackupAgent.MaxBackupManagerConfig.partnerUuid
        $Script:ProfileID    = $BackupAgent.MaxBackupManagerConfig.profileId

        if ($IsIntegrated -eq "False") {
            Write-Output "N-central Integration status  : $IsIntegrated"
            Write-Output "Partner Name                  : $Script:PartnerName"
            Write-Output "Migration Prep not required."
            Break
        }
        if ($IsIntegrated -eq "True") {
            Write-Output "N-central Integration status  : $IsIntegrated"
            Write-Output "Partner Name                  : $Script:PartnerName"
            #Write-Output "Backup Install Status         : $IsInstalled"
            #Write-Output "Backup CUID                   : $CUID"
            #Write-Output "Backup ProfileId              : $ProfileID"
        }

    }
    else{
        Write-Warning "No N-able Backup integration found, N-central migration to Cove not required or previously completed"
    }
}  ## Check N-Central for Backup Integration Status

Function Copy-Config {

    $Script:BMConfig = "C:\Program Files\Backup Manager\config.ini"
    $Script:BMConfigCopy = "C:\programdata\MXB\config.ini.copy"
    Write-Output "`nMaking a copy of the Backup config file"
    If ($Script:IsIntegrated -eq "True" ) {

        if ((Test-Path $Script:BMConfig -PathType leaf) -eq $false) {
            Write-Output "`nBackup Manager CONFIG.INI not found, Backup Manager may not be installed"
            Break
        }
        elseif ((Test-Path $Script:BMConfig -PathType leaf) -eq $true) {
            Write-Output "`nCopying existing Backup Manager CONFIG.INI to $Script:BMConfigCopy"
            Copy-Item $Script:BMConfig -Destination $Script:BMConfigCopy
        }

    }
    else{Write-warning "No N-able Backup integration found, skipping copy of CONFIG.INI"}
}   ## Copy Config.ini to ProgramData foder

Function Update-BackupManager {

    $Script:BackupIP = "c:\Program Files\Backup Manager\BackupIP.exe"
    
    Write-Output "`nUpgrading the Backup Manager prior to migration"       
    If (((Test-Path $Script:BackupIp -PathType leaf) -eq $true) -and ($Script:IsIntegrated -eq "True" )) {

        Get-JobStatus
        Write-Output "`nChecking status of runnign backups jobs"
        if ($Script:JobStatus -ne "Idle") {
            Write-Warning "Backup Client upgrade deferred while job status is not Idle`nMigration Prep should be rerun later"
            Break
        }
        else{
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $Url = "https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
            $Path = "C:\Users\Public\Downloads\mxb-windows-x86_x64.exe"
            $Username = ""
            $Password = "" 

            Write-Output "`nDownloading Latest Backup Manager Client"
            $WebClient = New-Object System.Net.WebClient
            $WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password)
            $WebClient.DownloadFile( $url, $path )

            Write-Output "`nUpdating Backup Manager Client"
            & $Path -silent
            Start-Sleep -Seconds 180
        }
    }
    else{
        Write-warning "No N-able Backup integration found, skipping upgrade of Backup Manager"
    }
}  ## Update Backup Manager to latest version

Function Rename-BackupIP {

    $Script:BackupIP = "c:\Program Files\Backup Manager\BackupIP.exe"
    $Script:DisabledBackupIP = "c:\Program Files\Backup Manager\BackupIP.disabled.exe"
    Write-Output "`nRenaming BackupIP.exe to block N-central uninstall of Cove after migration"

    if ((Test-Path $Script:BackupIp -PathType leaf) -eq $false) {      
        Write-Warning "BackupIP.exe is not present,the file may already be renamed"
    }elseif ((Test-Path $Script:BackupIp -PathType leaf) -eq $true) {      
        Move-Item $Script:BackupIp -destination $Script:DisabledBackupIP -force
        Write-Output "`nBackupIP.exe was found and has been disabled"
    }

} ## Rename BackupIP.exe to block N-central uninstall

Function ConvertTo-Base64($InputString) {
    $BytesString = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $global:OutputBase64 = [System.Convert]::ToBase64String($BytesString)
    Write-Output $global:OutputBase64
} ## Covert string to Base64 encoding

Function ConvertFrom-Base64($InputBase64) {
    $BytesBase64 = [System.Convert]::FromBase64String($InputBase64)
    $global:OutputString = [System.Text.Encoding]::UTF8.GetString($BytesBase64)
    Write-Output $global:OutputString
} ## Covert Base64 encoding to string

Function Get-TimeStamp {
    return "[{0:yy/MM/dd} {0:HH:mm}]" -f (Get-Date)
} ## Get formated Timestamp for use in logging

Function Get-JobStatus {

    $clienttool = "C:\Program Files\Backup Manager\ClientTool.exe"

    try { $ErrorActionPreference = 'Stop'; $Script:JobStatus = & $clienttool control.status.get }catch{ Write-Warning "ERROR     : $_" }

    if ($JobStatus) {

        #Write-output "`nJob Status                    : $JobStatus"
        Write-output "`nBackup Manager Job Status     : $JobStatus"
    }

} ## Get Backup Job Status from ClientTool.exe

Function Set-Alias {

    $BMConfig = "C:\Program Files\Backup Manager\config.ini"
    $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"  

    if ((Test-Path $BMConfig -PathType leaf) -eq $true) {
        
        $script:alias = "$env:computername $(get-timestamp) Cove Prep"

        $AliasParam = "-device-alias `"$(ConvertTo-Base64 $alias)`""

        Write-Output "`nUpdating Backup Manager Device Alias Column"    
        Write-Output "`n  Migration Prep Time         : $Script:Alias"
        #Write-Output "Alias Base64   : $AliasParam" 
        #Write-Output "Partner UID    : $UID"
        Write-Output ""

        $global:process = start-process -FilePath "$clienttool" -ArgumentList "takeover -config-path `"$bmconfig`" -partner-uid $Script:CUID $AliasParam" -PassThru

        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "N-able Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Installer" -Completed
                Break
                }
            }
        }
        else{
            Write-Output "`nSet Alias aborted, existing Backup Manager deployment not found"
            Break
            }
} ## Set local Device Alias Name to display the last Migration Prep timestamp

Get-BackupIntegrationStatus
Copy-Config
Update-BackupManager
Rename-BackupIP

Write-Output "`nN-central to Cove Migration Prep Status "
if (((Test-Path $Script:BMConfigCopy -PathType leaf) -eq $true) -and ((Test-Path $Script:BackupIp -PathType leaf) -eq $false) -and ($Script:IsIntegrated -eq "True" )<# -and ($PartnerName -match '\d{2,4} - *')#>)  {
    Write-Output "`nCove Migration Prep appears successful"
    Set-Alias
}else{ Write-Warning "Cove Migration Prep may be partially complete but the Device Alias was not updated."}
