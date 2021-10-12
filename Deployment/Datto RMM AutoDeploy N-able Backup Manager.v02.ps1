<## ----- About: ----
    # DATTO RMM - AutoDeploy N-able Backup Manager
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
    # D\L and deploy a NEW Backup Manager as a Passphrase compatible device
    # Optionally assign Profile (selections & schedules) and Product (retention) values
    #                  
    # Copy this Script into DATTO RMM
    # Create the following varibles in DATTO RMM to Pass through to the Script at run time
    #
    # Name: UID
    # Type: Variable
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ Backup.Management | Customers
    #
    # Name: Set_Profile
    # Type: Boolian
    # Value: True
    # Note: True to pass a Profile Name
    #
    # Name: Profile
    # Type: Variable
    # Value: Workstation
    # Note: Case sensitive, Found @ Backup.Management | Profiles
    #
    # Name: Set_Product
    # Type: Boolian
    # Value: True
    # Note: True to pass a Product Name
    #
    # Name: Profile
    # Type: Variable
    # Value: All-In
    # Note: Case sensitive, Found @ Backup.Management | Products
    #
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

if ( $env:Set_Profile -eq "true" ) { $BackupProfile = "-profile-name `"$env:Profile`"" }        ## $env: required to pass external varible from DATTO RMM to script

if ( $env:Set_Product -eq "true" ) { $BackupProduct = "-product-name `"$env:Product`"" }        ## $env: required to pass external varible from DATTO RMM to script

Start-Process -FilePath "c:\windows\temp\mxb-windows-x86_x64.exe" -ArgumentList "-unattended-mode -silent -partner-uid $env:UID $BackupProfile $BackupProduct" -passthru

###END SCRIPT###