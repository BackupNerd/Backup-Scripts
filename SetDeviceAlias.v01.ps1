<# # ----- About: ----
    # SolarWinds Backup Set Device Alias
    # Revision v01 - 2020-08-31
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
    # [-Alias "<STRING>] [-Uid] <string>                          ## Sets "Alias" for local device, requires 36 character -UID of current partner
    #
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/auto-deployment.htm
    # https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-installation/convert-to-passphrase.htm?Highlight=uid
    #
    # Important Note: Use of this Script will convert a Private Key Encrypted Device to a Passphrase Managed device.
    #                 Please contact me above if you have questions before running this script.   
    #
# -----------------------------------------------------------#>

[CmdletBinding()]
    Param (

        [parameter(Mandatory=$true,position=1)] [String]$Alias,  ## Optional "Device Name Alias" 
        
        [parameter(Mandatory=$true,position=2)] [ValidateLength(36,36)] [string]$UID  

    )

    clear-host
    
    # Check if script is running as Adminstrator
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if (-not $IsAdmin){
        Write-Host "The script is NOT running as Administrator, restart PowerShell as Administrator..."
                    
    }
    else{
        Write-Host "The script is running as Administrator"
    }
      
    Write-Output ""
    Write-Output "  SolarWinds Backup Set Device Alias"
    Write-Output ""
    Write-Output "  Script location:"
    Write-Output "  cd `"$PSScriptRoot`""
    Write-Output ""
 
    Function EncodeTo-Base64($InputString) {
        $BytesString = [System.Text.Encoding]::UTF8.GetBytes($InputString)
        $global:OutputBase64 = [System.Convert]::ToBase64String($BytesString)
        Write-Output $global:OutputBase64
        }  ## Function to convert Alias text to Base64 encoding
 
    Function DecodeFrom-Base64($InputBase64) {
        $BytesBase64 = [System.Convert]::FromBase64String($InputBase64)
        $global:OutputString = [System.Text.Encoding]::UTF8.GetString($BytesBase64)
        Write-Output $global:OutputString
        }  ## Function to convert Base64 encoding to text

    Function Set-Alias {
    
    $BMConfig = "C:\Program Files\Backup Manager\config.ini"
    $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"  

        if ((Test-Path $BMConfig -PathType leaf) -eq $true) {
            

            if ($Alias) { $AliasParam = "-device-alias `"$(EncodeTo-Base64 $alias)`""} 

            Write-Output "  Alias name     : $Alias"
            Write-Output "  Alias Base64   : $AliasParam" 
            Write-Output "  Partner UID    : $UID"

            Write-Output ""
            Write-Output "  Updating Backup Manager Device Alias"        
            Write-Output ""

            $global:process = start-process -FilePath "$clienttool" -ArgumentList "takeover -config-path `"$bmconfig`" -partner-uid $Uid $AliasParam" -PassThru

            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Solarwinds Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
                Start-Sleep -Milliseconds 100
                if ($process.HasExited) {
                    Write-Progress -Activity "Installer" -Completed
                    Break
                }
            }
                    
            } else {
            
            Write-Output ""
            Write-Output "  Set Alias aborted, existing Backup Manager deployment not found"
            Write-Output ""
            
            Break
            }
    }  ## Function to Call Clienttool.exe to perform Takeover an Set Device Alias 

    Set-Alias

    