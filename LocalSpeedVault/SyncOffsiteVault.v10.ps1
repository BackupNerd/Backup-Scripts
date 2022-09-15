<# ----- About: ----
    # Sync Offline LocalSpeedVault
    # Revision v10 - 2022-09-09
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
    # For use with N-able | Cove Data Protection 22.6 and later
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----

    # This Script is a powershell wrapper for the ServerTool command line seed loading utility.
    # It downloads, extracts and runs Servertool sequentually against a CSV device list to build or update one or more offsite speedvaults
    #
    # Use the -Help switch parameter to view script parameters
    # Use the -Network switch parameter to designate Network Attached Storage (NAS)
    #   Use the -NetworkPath parameter to set the base network path, otherwise a default path is used
    #   Use the -NetworkTemp parameter to set a temp drive letter for the network path, otherwise a default parameter is used
    #   Use the -NetworkUser parameter to set the user name for the network path, otherwise you will be prompted
    #   Use the -NetworkPW parameter to set the password for the network path, otherwise you will be prompted
    #   ** Important: Care should be taken at any time network credentials are embeded or passed in a script parameter.
    #   ** Failure to pass credentials in the script or via parameter will cause a secure prompt for credentials
    #
    # Use the -Local switch parameter to designate Direct Attached Storage (DAS)
    #   Use the -LocalPath parameter to set the base network path, otherwise a default parameter path is used
    # Use the -ThreadCount parameter to set the number of download connections to use (1-50), otherwise a default parameter is used
    # Use the -RetryCount parameter to set the number of retry attempts to perform(1-5), otherwise a default parameter is used
    # Use the -DeviceList parameter to specify the path to a CSV file with "Device Name" and "Installation Key" for each device, otherwise a default path
    # Use the -SelectDevices switch parameter to force GUI device selection
    # 
    # Usage: 
    # .\SyncOffsiteVault.v10.ps1 -local 
    # .\SyncOffsiteVault.v10.ps1 -local -localpath "d:\vault"
    # .\SyncOffsiteVault.v10.ps1 -network -networkpath "\\nas\share" 
    # .\SyncOffsiteVault.v10.ps1 -network -networkpath "\\nas\share" -networkuser "nas\admin" -networkpw "sharepw" -networktemp Z -threadcount 20 -devicelist "c:\temp\devicelist.csv"
    # 
    # https://serverfault.com/questions/186030/how-to-use-a-config-file-ini-conf-with-a-powershell-script ##Future Ideas
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/seed-restore.htm
    #
    # Sample Servertool commands:
    # c:\servertool\servertool.exe seed.download -account laptop-53hul6dm_12345 -password 2687f08eb -path f:\OffsiteVault -thread-count 10 -retry-count 2
    # f:\servertool\servertool.exe seed.download -account laptop-53hul6dm_12345 -password 2687f08eb -path \\Server\share -thread-count 20 -retry-count 5
    #
    #  The -account & -password references above point to the "Device name" & "Installation key" columns found in the https://backup.management console
    #  You can generate a compatable CSV file by adding these colums in the https://backup.management console and then using the Export option to get and XLSX file
    #  Then same the XLS file as a CSV

    <# Sample CSV Format

    Device name,Installation key,
    2019-man_hpw2c,7f86da88730e
    2019-lis_drbpu,3828260aae8b
    2019-dus_hwf7d,7b7a6b322b3b
    2019-doh_gvsnj,6b0fa0e7f6e3
    2019-ayt_ki9yx,e3ad88720eaf
    2019-hel_ol09i,88b88ab668fa
    2019-klp_tzh5r,66a2c22a6c8b
    2019-mco_224ln,6b870c68822d
    2019-jfk_2g73d,f0fae23d827b
    
    #>
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Network")]                                                                     ## Change between Network, Local or Help to assign a default 
    Param (
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [Switch]$Network,                                     ## Select a network target
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$NetworkPath = "\\Ivy\backup",                ## Set a network OffsiteSpeedVault target
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$NetworkTemp = "Z",                           ## Set a temp drive letter for Network share
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$Networkuser = "Ivy\Administrator",           ## Set username for network path
        [Parameter(ParameterSetName="Network",Mandatory=$False)] [String]$NetworkPW,                                   ## Set Password for network path
        [Parameter(ParameterSetName="Local",Mandatory=$False)] [Switch]$Local,                                         ## Select a local target
        [Parameter(ParameterSetName="Local",Mandatory=$False)] [String]$LocalPath = "D:\Vault",                        ## Set a local OffsiteSpeedVault target
        [Parameter(ParameterSetName="Help",Mandatory=$False)] [Switch]$Help,                                           ## Display Help
        [Parameter(Mandatory=$False)][ValidateRange(1,100)][int]$threadcount = 100,                                    ## Seed download thread count
        [Parameter(Mandatory=$False)][ValidateRange(1,5)][int]$retrycount = 5,                                         ## Seed dowload error retry count
        [Parameter(Mandatory=$False)][switch]$SelectDevices = $true,                                                   ## $True to prompt for Device Selection
        [Parameter(Mandatory=$False)][string]$DeviceList = "D:\devicelist.csv"                                         ## Path to CSV file with devicename and devicepassword
 
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    
    $ConsoleTitle = "Synchronize OffsiteSpeedVault"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    Write-output "  $ConsoleTitle`n`n"
    $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "`n`n`n`n`n  Script Parameter Syntax:`n`n$Syntax"
    Write-output    "  Current Parameters:"
    Write-output    "  -Mode               = $($PsCmdlet.ParameterSetName)"
    If ($Network) { "  -OffsiteVaultPath   = $NetworkPath" }
    If ($Local) {   "  -OffsiteVaultPath   = $LocalPath" }

    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
    $ServerToolDownloadPath = "c:\windows\temp\ServerTool.zip"
    Remove-psdrive -name $NetworkTemp -ErrorAction SilentlyContinue -Force
    $templetter = $networktemp + ":"
    net use $templetter /d

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

    Function Get-LogTimeStamp {
        #Get-Date -format s
        return "[{0:yyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    }  ## Output proper timeStamp for log file 

    Function Download-ServerTool {
        $script:ServerToolURL          = "https://cdn.cloudbackup.management/maxdownloads/mxb-st-windows-x64.zip"

        Write-Output "  Downloading Latest Cove Data Protection Servertool Build`n"
        (New-Object System.Net.WebClient).DownloadFile($ServerToolURL,$ServerToolDownloadPath)

        Function Expand-ServerTool {
            Write-Output "  Extracting Latest Cove Data Protection Servertool Build`n"
            if (!(Test-Path -Path $ServerToolExtractPath)) {
                Write-output "  Creating ServerTool Installation Path"
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
        $targetfolder = $VaultStoragePath
        $dataColl = @()
        Get-ChildItem -force $targetfolder -ErrorAction SilentlyContinue | Where-Object { $_ -is [io.directoryinfo] } | ForEach-Object {
        $len = 0
        Get-ChildItem -recurse -force $_.fullname -ErrorAction SilentlyContinue | ForEach-Object { $len += $_.length }
        $foldername = $_.fullname
        $foldersize= '{0:N2}' -f ($len / 1Gb)
        $LastWrite= $_.LastWriteTime
        $dataObject = New-Object PSObject
        Add-Member -inputObject $dataObject -memberType NoteProperty -name "Path" -value $foldername
        Add-Member -inputObject $dataObject -memberType NoteProperty -name "Size(GB)" -value $foldersize
        Add-Member -inputObject $dataObject -memberType NoteProperty -name "LastWrite" -value $LastWrite
        $dataColl += $dataObject
        }
        Write-output "`n  $(Get-LogTimeStamp) Current OffsiteSpeedVault Usage"       # Write timestamp    
        $dataColl | Sort-Object LastWrite | format-table
    }

    Function Get-FreeSpace {
        $Drives = Get-PSDrive -PSProvider FileSystem
        
        ForEach ($Drive in $Drives) {
        "`n {0,0} {1,21:N2} {2,10:P0}`n" -f $Drive.Name,
        ($Drive.Free/1gb + $Drive.Used/1gb),
        ($Drive.free / (($Drive.free +0.001) + $Drive.used ))
        }
    }

    Function Start-SeedDownload {
        
        foreach ($Device in $Selected) {
        
        $ConsoleTitle = "Synchronize OffsiteSpeedVault for $($device.{Device name})"
        $host.UI.RawUI.WindowTitle = $ConsoleTitle
        
        Write-output "`n  $(Get-LogTimeStamp) Started Seed.Download for $($device.{Device name})"     # Write start timestamp        
        
        $process = start-process -FilePath $ServerTool -ArgumentList "seed.download -account $($device.{Device name}) -password $($device.{Installation key}) -path $($VaultStoragePath) -thread-count $($threadcount) -retry-count $($retrycount)" -passthru -WindowStyle Minimized
        Write-output "`n**Servertool Command**`n`nServertool.exe seed.download -account $($device.{Device name}) -password $($device.{Installation key}) -path $($VaultStoragePath) -thread-count $($threadcount) -retry-count $($retrycount)"
    
    # Loop device list
            for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
                Write-Progress -Activity "Synchronizing $($device.{Device name}) to $($VaultStoragePath)" -PercentComplete $i -Status "Processing"
                Start-Sleep -seconds 2
                
                #$a = Get-Process -Name servertool | ForEach-Object { Get-NetTCPConnection -OwningProcess $_.Id -ErrorAction SilentlyContinue }
                #$a | sort-object remoteaddress | where-object {($_.remoteaddress -ne "127.0.0.1") -and ($_.remoteaddress -ne "0.0.0.0")} | format-table
                
                if ($process.HasExited) {
                    Write-Progress -Activity "Synchronization" -Completed

                    Write-output "`n  $(Get-LogTimeStamp) Ended Seed.Download for $($device.{Device name})"       # Write end timestamp
                    Get-FolderSize     
                    Break
                }
            }     
        }
    } # Start Seed.Download

#endregion ----- Functions ----

#region ----- Main body ----    

    Switch ($PSCmdlet.ParameterSetName) { 
        'Network' {
            $OffsiteVaultRootPath = $NetworkPath

            if (($networkUser) -and ($networkPW)) {
                $pass = $networkPW | ConvertTo-SecureString -AsPlainText -Force
                $Cred = New-Object System.Management.Automation.PsCredential($networkUser,$pass)
                Remove-Variable -name NetworkPW
                New-PSDrive -name $NetworkTemp -PSProvider "FileSystem" -Root $OffsiteVaultRootPath -Credential $cred -Scope Script -Persist
                
            }else{
                $cred = Get-Credential -message "`nEnter network share user credentials for OffsiteSpeedVault`n`n  User examples:`n  [10.10.10.200\user]`n  [workgroup\user]`n  [domain\user]`n  [server\user]`n  [nas\user]"
                New-PSDrive -name $NetworkTemp -PSProvider "FileSystem" -Root $OffsiteVaultRootPath -Credential $cred -Scope Script -Persist
            }
            
            Get-FreeSpace
            
            if(!(Test-Path -Path $OffsiteVaultRootPath)) {
                Write-output "  OffsiteSpeedVault Root Path ( $OffsiteVaultRootPath ) does not exist or not accessible"
                [void](New-Item -ItemType Directory -Force -Path $OffsiteVaultRootPath)
            }

            $script:ServerToolExtractPath  = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\ServerTool\"
            $script:ServerTool             = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\ServerTool\Servertool.exe"
            $script:VaultStoragePath       = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\"
            
            Download-ServerTool

            $lookup = Import-csv -path $DeviceList      ## Read CSV

            # Select Devices
            
            if ($SelectDevices) {
                Write-warning "`n--== FIND THE POP-UP DEVICE LIST AND SELECT ONE OR MORE DEVICES TO BUILD/SYNC AS AN OFFSITE SPEED VAULT ==--"
                $script:selected = $lookup | Out-GridView -OutputMode Multiple -Title "SELECT ONE OR MORE DEVICES TO BUILD/SYNC AS AN OFFSITE SPEED VAULT"
            }
            else {
                $script:selected = $lookup
                $selected.count 
            }
        
            Start-SeedDownload

            Remove-psdrive -name $NetworkTemp
            $templetter = $networktemp + ":"
            net use $templetter /d

            }
        'Local' { 
            $OffsiteVaultRootPath = $LocalPath 
            Get-PSDrive -name "$($localpath[0])"
            
            Get-FreeSpace

            if(!(Test-Path -Path $OffsiteVaultRootPath)) {
                Write-output "  OffsiteSpeedVault Root Path ( $OffsiteVaultRootPath ) does not exist or not accessible"
                [void](New-Item -ItemType Directory -Force -Path $OffsiteVaultRootPath)
            }

            $script:ServerToolExtractPath  = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\ServerTool\"
            $script:ServerTool             = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\ServerTool\Servertool.exe"
            $script:VaultStoragePath       = join-path -path $OffsiteVaultRootPath -childpath "OffsiteSpeedVault\"
            
            Download-ServerTool

            $lookup = Import-csv -path $DeviceList      ## Read CSV

            # Select Devices
            $selected = $lookup | Out-GridView -OutputMode Multiple
        
            Start-SeedDownload

            }
        'Help' { 
            Exit
            }
    }










