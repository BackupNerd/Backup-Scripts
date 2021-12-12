<# ----- About: ----
    # Sync Offline LocalSpeedVault
    # Revision v02 - 2021-02-15
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
# -----------------------------------------------------------#>  ## About
<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with the Standalone edition of SolarWinds Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----

    # Run Servertool against a defined device list
    #
    # https://serverfault.com/questions/186030/how-to-use-a-config-file-ini-conf-with-a-powershell-script
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-guide/seed-restore.htm
    #
    # f:\servertool\servertool.exe seed.download -account laptop-53hul6dm_12345 -password 2687f08eb -path f:\seeddownload -thread-count 10
    # f:\servertool\servertool.exe seed.download -account laptop-53hul6dm_12345 -password 2687f08eb -path \\ivy\speedvault -thread-count 10
    #
    # Check script location for config.ini, else create one.
    # Read Config.in for Path to OffsiteSpeedVault, connect, diagnostics, size
    #   Else, Prompt for an existing root path for the OffsiteSpeedVault i.e P:\ or \\server\share
    #       if direct attached, check if exist, else create
    #       if network, check if connect, else prompt for credentials, net connect, check for path, store secure credentials
    #       Connect, Report size, read, write test
    # Read Config.ini for Backup Edition
    #   Else, Specify Edition (RMM,Standalone,Ncentral)
    #   if servertool not exist, download and extract
    #   if servertool exist, check date is older than 28 days,download, extract (Skip 7 day internval)
    #       Else, proceed
    #   If devicelist exist in script directory, Read list, check password for length, check RMM for length, else fail
    #   if no devicelist, prompt to create sample device list, read sample, prompt for user to update device list if sample found
    # Read config.ini for Threads count, else prompt (10 is default)

# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Help")]
    Param (
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [Switch]$Network,                                     ## Select a network target
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$NetworkPath = "\\Ivy\backup",                ## Set a network OffsiteSpeedVault target
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$NetworkTemp = "Z",                           ## Set a temp drive letter for Network share
        [Parameter(ParameterSetName="Local",Mandatory=$False)] [Switch]$Local,                                         ## Select a local target
        [Parameter(ParameterSetName="Local",Mandatory=$False)] [String]$LocalPath = "P:\Data\MyVault",                 ## Set a local OffsiteSpeedVault target
        [Parameter(ParameterSetName="Help",Mandatory=$False)] [Switch]$Help,                                           ## Display Help
        [Parameter(Mandatory=$False)][ValidateSet('RMM','N-central','Standalone')][String]$Product = "Standalone",     ## RMM, N-central, Standalone
        [Parameter(Mandatory=$False)][ValidateRange(1,200)][int]$threadcount = 5,                                      ## Seed download thread count
        [Parameter(Mandatory=$False)][ValidateRange(1,5)][int]$retrycount = 5,                                         ## Seed dowload error retry count
        [Parameter(Mandatory=$False)][string]$DeviceList = "Q:\devicelist.csv"                                         ## Path to CSV file with devicename and devicepassword
 
    )

#region ----- Functions ----

    Function Get-LogTimeStamp {
        #Get-Date -format s
        return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    }  ## Output proper timeStamp for log file 

    Function Download-ServerTool {
        Write-Output "  Downloading Latest Server Tool"
        (New-Object System.Net.WebClient).DownloadFile($ServerToolURl,$ServerToolDownloadPath)

        Function Expand-ServerTool {
            Write-Output "  Extracting Latest Server Tool"
            if (!(Test-Path -Path $ServerToolExtractPath)) {
                Write-output "  Creating ServerTool installation Path"
                [void] (New-Item -ItemType Directory -Force -Path $ServerToolExtractPath)
            }
            
            Expand-Archive -Path $ServerToolDownloadPath -DestinationPath $ServerToolExtractPath -Force

            }
        
            Expand-ServerTool

            if (Test-Path $ServerToolDownloadPath) {
                Remove-Item $ServerToolDownloadPath
        }
    }

    Function Get-FolderSize {
        $targetfolder = $LockedVaultStoragePath
        $dataColl = @()
        Get-ChildItem -force $targetfolder -ErrorAction SilentlyContinue | Where-Object { $_ -is [io.directoryinfo] } | ForEach-Object {
        $len = 0
        Get-ChildItem -recurse -force $_.fullname -ErrorAction SilentlyContinue | ForEach-Object { $len += $_.length }
        $foldername = $_.fullname
        $foldersize= '{0:N2}' -f ($len / 1Gb)
        $dataObject = New-Object PSObject
        Add-Member -inputObject $dataObject -memberType NoteProperty -name "Path" -value $foldername
        Add-Member -inputObject $dataObject -memberType NoteProperty -name "Size(GB)" -value $foldersize
        $dataColl += $dataObject
        }
        Write-output "`n  $(Get-LogTimeStamp) Current OffsiteSpeedVault Usage"       # Write timestamp    
        $dataColl | format-table
    }

    Function Get-FreeSpace {
        $Drives = Get-PSDrive -PSProvider FileSystem
        ForEach ($Drive in $Drives) {
        "`n {0,0} {1,21:N2} {2,10:P0}`n" -f $Drive.Name,
        ($Drive.Free/1gb + $Drive.Used/1gb),
        ($Drive.free / (($Drive.free +0.001) + $Drive.used))
        }
    }



#endregion ----- Functions ----

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Synchronize OffsiteSpeedVault`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "`n`n`n`n`n  Script Parameter Syntax:`n`n$Syntax"
    Write-output    "  Current Parameters:"
    Write-output    "  -Mode               = $($PsCmdlet.ParameterSetName)"
    If ($Network) { "  -OffsiteVaultPath   = $NetworkPath" }
    If ($Local) {   "  -OffsiteVaultPath   = $LocalPath" }
    Write-output    "  -Product            = SolarWinds $Product Backup"

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $ServerToolDownloadPath = "c:\windows\temp\ServerTool.zip"

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Main body ----    

    Switch ($PSCmdlet.ParameterSetName) { 
        'Network' {
            $OffsiteVaultRootPath = $NetworkPath
            $cred = Get-Credential -message "`nEnter network share user credentials for OffsiteSpeedVault`n`n  User examples:`n  [10.10.10.200\user]`n  [workgroup\user]`n  [domain\user]`n  [server\user]`n  [nas\user]"
            New-PSDrive -name "$NetworkTemp" -PSProvider "FileSystem" -Root $OffsiteVaultRootPath -scope script -Credential $cred -persist
            Get-FreeSpace
            Remove-psdrive -name "$NetworkTemp"
            }
        'Local' { 
            $OffsiteVaultRootPath = $LocalPath 
            Get-PSDrive -name "$($localpath[0])"
            Get-FreeSpace
            }
        'Help' { 
            #"  Switch Type    : $($PSCmdlet.ParameterSetName)"
            
            Exit
            }
    }

    if(!(Test-Path -Path $OffsiteVaultRootPath)) {
        Write-output "  OffsiteSpeedVault Root Path ( $OffsiteVaultRootPath ) does not exist or not accessible"
        [void](New-Item -ItemType Directory -Force -Path $OffsiteVaultRootPath)
    }

    Switch ($Product) { 
        'RMM' {
            $ServerToolExtractPath  = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\RMM\ServerTool\"
            $ServerTool             = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\RMM\ServerTool\Servertool.exe"
            $LockedVaultStoragePath = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\RMM\LockedVault\"
            $VaultStoragePath       = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\RMM\Vault\"
            $ServerToolURl          = "https://cdn.cloudbackup.management/mobdownloads/backup-and-recovery-st-windows-x64.zip"
            Download-ServerTool
            
            $lookup = Import-csv -path $DeviceList      ## Read CSV

            }
        'N-central' { 
            $ServerToolExtractPath  = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\NC\ServerTool\"
            $ServerTool             = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\NC\ServerTool\Servertool.exe"
            $LockedVaultStoragePath = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\NC\LockedVault\"
            $VaultStoragePath       = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\NC\Vault\"
            $ServerToolURl          = "https://cdn.cloudbackup.management/maxdownloads/mxb-st-windows-x64.zip"
            Download-ServerTool

            $lookup = Import-csv -path $DeviceList      ## Read CSV

            }
        'Standalone' {

            $ServerToolExtractPath  = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\MXB\ServerTool\"
            $ServerTool             = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\MXB\ServerTool\Servertool.exe"
            $LockedVaultStoragePath = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\MXB\LockedVault\"
            $VaultStoragePath       = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\MXB\Vault\"
            $ServerToolURl          = "https://cdn.cloudbackup.management/maxdownloads/mxb-st-windows-x64.zip"
            Download-ServerTool

            $lookup = Import-csv -path $DeviceList      ## Read CSV

            }
    }

# Select Devices
    $selected = $lookup | Out-GridView -OutputMode Multiple
    #$selected | Select-Object devicename,password

# Start Seed.Download

    if(!(Test-Path -Path $VaultStoragePath)) {
        Write-output "  Creating\Locking Vault Storage Path"
        [void](New-Item -ItemType Directory -Force -Path $VaultStoragePath)
        [void](Move-Item -path  $VaultStoragePath -destination $LockedVaultStoragePath -Force)
        
    }else{
        Write-output "  Moving\Locking Vault Storage Path"
        [void](Move-Item -path  $VaultStoragePath -destination $LockedVaultStoragePath -Force)
        Get-FolderSize}

        do {
            $dirs = Get-ChildItem $LockedVaultStoragePath -directory -recurse | Where-Object { (Get-ChildItem $_.fullName -Force).count -eq 0 } | Select-Object -expandproperty FullName
            $dirs | Foreach-Object { Remove-Item $_ }
          } while ($dirs.count -gt 0)

    foreach ($Device in $Selected) {
     
    Write-output "`n  $(Get-LogTimeStamp) Started Seed.Download for $($device.devicename)"     # Write start timestamp        
    
    $process = start-process -FilePath $ServerTool -ArgumentList "seed.download -account $($device.devicename) -password $($device.password) -path $($LockedVaultStoragePath) -thread-count $($threadcount) -retry-count $($retrycount)" -passthru -WindowStyle Minimized
    Write-output "`n  **Command** Servertool.exe seed.download -account $($device.devicename) -password $($device.password) -path $($LockedVaultStoragePath) -thread-count $($threadcount) -retry-count $($retrycount)"
   
# Loop device list
        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "Synchronizing $($device.devicename) to $($LockedVaultStoragePath)" -PercentComplete $i -Status "Processing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Synchronization" -Completed

                Write-output "`n  $(Get-LogTimeStamp) Ended Seed.Download for $($device.devicename)"       # Write end timestamp
                Get-FolderSize     
                Break
            }
        }     
    }

    do {
        $dirs = Get-ChildItem $LockedVaultStoragePath -directory -recurse | Where-Object { (Get-ChildItem $_.fullName -Force).count -eq 0 } | Select-Object -expandproperty FullName
        $dirs | Foreach-Object { Remove-Item $_ }
      } while ($dirs.count -gt 0)

    if((Test-Path -Path $VaultStoragePath)) {
        Write-output "  Cleaning Up\Unlocking Vault Storage Path"
        [void](Remove-Item -Path $VaultStoragePath -Force -Recurse)
        [void](move-Item -path  $LockedVaultStoragePath -Destination $VaultStoragePath -Force)
 
    }else{
        Write-output "  Moving\Unlocking Vault Storage Path"
        [void](move-Item -path  $LockedVaultStoragePath -Destination $VaultStoragePath -Force)
        Get-FolderSize}





