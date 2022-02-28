<## ----- About: ----
    # DATTO RMM - Deploy Documents N-able Backup Manager
    # Revision v03 - 2021-12-09
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Behavior: ----
    # D\L and deploy a NEW Backup Manager 'Documents' device for Window Workstations
    # Check for existing installation and exit without installing if found 
    #         
    # Copy this Script into DATTO RMM
    # Create the following varible in DATTO RMM to Pass through to the Script at run time
    #
    # DATTO RMM SCRIPT INPUT VARIABLES
    #
    # Name: Backup_UID
    # Type: Variable Value
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Description: Found @ https://Backup.Management | Customers
    # Description: 36-character Customer UID string
    #
    # Execute Script as Administrator 
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
# -----------------------------------------------------------#>

###START SCRIPT###

if ((Test-Path "C:\Program Files\Backup Manager\config.ini" -PathType leaf) -eq $true) { Write-output "Backup Manager already present"; Break }
else{

    ## Start Download Script

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    ## Remove above line if you have SSL/TLS issues on legacy OS installations

    $Url = "https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
    ## Download source url, change above line to HTTP if you have SSL/TLS issues on legacy OS installations

    $Path = "C:\windows\temp\mxb-windows-x86_x64.exe"
    ## Download target path
    
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile( $url, $path )

    ## Start Install Script

    $BackupProfile = "-profile-name `"Documents`""

    $BackupProduct = "-product-name `"Documents`""

    Start-Process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-unattended-mode -silent -partner-uid $env:Backup_UID $BackupProfile $BackupProduct" -passthru
    ## $env: prefix required to pass external variable from DATTO RMM to script

}
###END SCRIPT###