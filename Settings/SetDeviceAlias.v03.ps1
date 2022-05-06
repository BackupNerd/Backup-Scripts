<# ----- About: ----
    # 
    # Set Backup Device Alias
    # Revision v03 - 2022-05-05
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with the following N-able Backup product editions:
    #   N-able Cove Data Protection             .... Compatible
    #   N-able N-central MSP Backup integrated  .... Compatible
    #   N-able RMM Backup & Recovery integrated .... Not Compatible
    # 
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Specify the following parameters to update the Alias of an existing backup device.
    # [-Uid] <36 Character string>              ## Customer UID found @ https://backup.management | Customers | Edit                                
    # [-Alias] <string>                         ## New Alias name       
    #
    # WARNING: This command will will perform a Takeover and Convert your Private-Key encryption to System-Managed Passphrase based encryption

# -----------------------------------------------------------#>  ## Behavior

param(

    [parameter(Mandatory=$false)][ValidateLength(36,36)][string]$UID = "dec8a8eb-9bf6-4a15-afe2-9bb6e5b94aaf",   ## Customer UID Default unless overrided
    [parameter(Mandatory=$true)] [string]$Alias                                                                  ## Alias
        
)

Clear-Host

Function EncodeTo-Base64($InputString) {
    $BytesString = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $global:OutputBase64 = [System.Convert]::ToBase64String($BytesString)
    Write-Output $global:OutputBase64
    }

$AliasParam = "-device-alias `"$(EncodeTo-Base64 $Alias)`""
$ConfigParam = "-config-path `"c:\program files\backup manager\config.ini`""
    
start-process -FilePath "c:\program files\backup manager\clienttool.exe" -ArgumentList "takeover $configParam -partner-uid $UID $AliasParam"  


<#  ## ALt code to download the Backup Manager installer for the takeover instead of using the existing Clienttool.exe

Function Download-BackupManager {
    "  Downloading Backup Manager"
    (New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","c:\windows\temp\mxb-windows-x86_x64.exe")
    }

    if ((Test-Path "c:\windows\temp\mxb-windows-x86_x64.exe" -PathType leaf) -eq $false) {
        Download-BackupManager
    }

    start-process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-unattended-mode -silent -partner-uid $UID -takeover $AliasParam"




    
    Future Conciderations:

    Use Version 2.0 code to read alias text from local file or reg key to sync with third party PSA
    Option to grab net bios name
    test passing API credentials for existing device to authenticate and lookup UID


    #>