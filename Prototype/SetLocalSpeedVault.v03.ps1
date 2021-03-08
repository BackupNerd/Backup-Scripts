<# ----- About: ----
    # Set LocalSpeedVault
    # Revision v03 - 2021-03-06
    # Author: Eric Harless, Head Backup Nerd - SolarWinds MSP | N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with all editions of SolarWinds | N-able Backup
    # LSV Compatible with all Windows clients except clients running a Documents Only license
    # Cannot modify LSV paths set via Profile
    
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----

    # Configure LocalSpeedVault Settings with Clienttool.exe
    # Specify -Local, -Network or -Disable Mode
    # Validate path / warn of errors
    # Optionally -Force apply LSV Path
    #
    # Use the -local switch parameter to Select a local LSV target
    # Use the -localPath parameter to specify a local path examples: [D:\SpeedVault][Q:\Data\Vault]
    # Use the -Network switch parameter to Select a network LSV target
    # Use the -NetworkPath parameter to specify a network path examples: [\\10.10.10.200\share][\\server\share]
    # Use the -NetworkUser parameter to specify a network user examples: [10.10.10.200\user][workgroup\user][domain\user][server\user][nas\user]
    # Use the -NetworkPassword parameter to specify a network password
    # Use the -Force switch parameter to force apply the local or network LSV target
    # Use the -Disable switch parameter to disable the LSV target
    # Use the -Help switch parameter to display parameter syntax
    #
    # Recommendations for using a NAS device for the LSV share:
    #
    #   Both NAS devices integrated with Active Directory and in a workgroup are supported for LSV.
    #   To limit access to the LSV share, a NAS device in a workgroup is recommended.
    #   For NAS devices in a workgroup, limited access is easier to manage.
    #   For each device use a unique username and password to access the LSV share.
    #   Do not reuse this username / password combo with other shares on the NAS.
    #   Do not use administrative logins / passwords.
    #   Do not map drives to the LSV Location
    #   Limit access to the defined LSV share for all other users and groups on the NAS.
    #
    #   https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/command-line.htm
    #   https://documentation.solarwindsmsp.com/backup/documentation/Content/backup-manager/backup-manager-guide/localspeedvault.htm

# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding(DefaultParameterSetName="Help")]
    Param (
        [Parameter(ParameterSetName="Network",Mandatory=$False)][Switch]$Network,                                                                   ## Select a network LSV target
        [Parameter(ParameterSetName="Network",Mandatory=$False,Position=0)][ValidateNotNullOrEmpty()][String]$NetworkPath,                          ## examples: [\\10.10.10.200\share][\\server\share]
        [Parameter(ParameterSetName="Network",Mandatory=$False,Position=1)][ValidateNotNullOrEmpty()][String]$NetworkUser,                          ## examples: [10.10.10.200\user][workgroup\user][domain\user][server\user][nas\user]
        [Parameter(ParameterSetName="Network",Mandatory=$False,Position=2)][ValidateNotNullOrEmpty()][String]$NetworkPassword,                      ## Specify network password
        [Parameter(ParameterSetName="Local",Mandatory=$False)][Switch]$Local,                                                                       ## Select a local LSV target
        [Parameter(ParameterSetName="Local",Mandatory=$False,Position=0)][String]$LocalPath = "Q:\Data\Vault",                                      ## examples: [D:\SpeedVault][Q:\Data\Vault]
        [Parameter(ParameterSetName="Disable",Mandatory=$False)][Switch]$Disable,                                                                   ## Disable LSV
        [Parameter(Mandatory=$False)][Switch]$Force,                                                                                                ## Force LSV Settings without validation
        [Parameter(ParameterSetName="Help",Mandatory=$False)][Switch]$Help                                                                          ## Help
 
    )

#region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    Write-output "  Set LocalSpeedVault`n"
    Write-output    "  Current Parameters:"
    Write-output    "  -Mode               = $($PsCmdlet.ParameterSetName)"
    if ($Network) { "  -LSV Path           = $NetworkPath" }
    if ($Network) { "  -LSV User           = $NetworkUser" }
    if ($Local) {   "  -LSV Path           = $LocalPath" }
    $clientool = "c:\program files\backup manager\clienttool.exe"
    $Global:ERRfound = $false
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
    Push-Location $dir

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Main body ----    

    Switch ($PSCmdlet.ParameterSetName) { 
        'Disable' { 
            $LSVEnabled = "-name LocalSpeedVaultEnabled -value 0"
            start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled" -WindowStyle Minimized
            }
        'Local' { 
            $LSVEnabled = "-name LocalSpeedVaultEnabled -value 1"
            if ($LocalPath){$LSVPath = "-name LocalSpeedVaultLocation -value $LocalPath"}
            if (Test-Path $LocalPath) { 
                start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled $LSVPath"
                Write-Output  "  PATH OK: LSV Path Set"
            }else{
                New-Item -type directory -path $LocalPath -Force -ErrorAction SilentlyContinue | Out-Null
                Write-Host  "  Attempting to Create LSV Path"
                if ($force) {
                    start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled $LSVPath" -WindowStyle Minimized
                    Write-Host  "  FORCED PATH: LSV Path Set"
                }else{
                    if (Test-Path $LocalPath) { 
                        start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled $LSVPath" -WindowStyle Minimized
                        Write-Host  "  PATH OK: LSV Path Set"
                    }else{
                        $Global:ERRfound = $true
                        Write-Host  "  PATH ERR: LSV Path not found or can't be created"
                        Write-Host "  Error"
                        exit 1001
                        }
                    }       
                }
            }            
        'Network' {
            $LSVEnabled = "-name LocalSpeedVaultEnabled -value 1"
            if ($NetworkPath){$LSVPath = "-name LocalSpeedVaultLocation -value $NetworkPath"}
            if ($NetworkUser){$LSVUser = "-name LocalSpeedVaultUser -value $NetworkUser"}
            if ($NetworkPassword){$LSVPassword = "-name LocalSpeedVaultPassword -value $NetworkPassword"}
            $parentPath = "\\" + [string]::join("\",$networkpath.Split("\")[2])
            if ($force) {
                start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled $LSVPath $LSVUser $LSVPassword" -WindowStyle Minimized
                Write-Host  "  FORCED PATH: LSV Path Set"
                get-psdrive | Where-object {($_.displayroot -like "*$parentPath*") -or ($_.displayroot -like "*\\$altpath*")}

            }else{
                try {
                    $altpath = Resolve-DnsName $parentPath.replace('\\','') -ErrorAction Stop | Select-Object IPAddress,NameHost
                    }
                    catch {
                        $Global:ERRfound = $true
                        Write-Warning -Message "Record not found for $($parentPath.replace('\\',''))"
                        Write-Host "  Error"
                        exit 1001
                        
                    }

                if ($altpath.IPAddress){$altpath = $altpath.IPAddress}elseif($altpath.namehost){$altpath = $altpath.namehost.split(".")[0]}
                
                if ($altpath) {
                    $AltNetworkPath = $networkpath -replace ($parentPath.replace('\\','')), $altpath
                    Write-output "  -ALT Path           = $AltNetworkPath"
                }
                
                if ((Test-Path $NetworkPath) -or (Test-Path $AltNetworkPath))  { 
                    Write-Host  "  PATH ERR: The LSV path or a similar network path is already mapped, this may not be a secure path."
                    get-psdrive | Where-object {$_.displayroot -like "*$parentPath*" -or $_.displayroot -like "*\\$altpath*"}
                    Write-Host "  Error"
                    exit 1001
                }else{
                    start-process -FilePath $clientool -ArgumentList "control.setting.modify $LSVEnabled $LSVPath $LSVUser $LSVPassword" -WindowStyle Minimized
                    Write-Host  "  PATH OK: LSV Path Set"
                    get-psdrive | Where-object {($_.displayroot -like "*$parentPath*") -or ($_.displayroot -like "*\\$altpath*")}
                    }
                }
            }
        'Help' {
            $Syntax = Get-Command $PSCommandPath -Syntax ; Write-Output "`n  Script Parameter Syntax:`n`n$Syntax" 
            Exit
            }
    }

    if ($Global:ERRfound) {
        Write-Host "  Error"
        exit 1001

    } else {
        Write-Host "  Success"
        exit 0

    } 

        