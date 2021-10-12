<## ----- About: ----
    # DATTO RMM - Upgrade N-able Backup Manager
    # Revision v02 - 2021-10-12
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
    #          
    # Copy this Script into DATTO RMM
    # Execute Script as Administrator 
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
# -----------------------------------------------------------#>

###START SCRIPT###

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Url = "https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe"
$Path = "C:\windows\temp\mxb-windows-x86_x64.exe"
$Username = ""
$Password = ""
 
$WebClient = New-Object System.Net.WebClient
$WebClient.Credentials = New-Object System.Net.Networkcredential($Username, $Password)
$WebClient.DownloadFile( $url, $path )

Start-Process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe"

###END SCRIPT###