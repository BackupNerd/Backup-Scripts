# ----- About: ----
    # N-able Cove Data Protection Throttle Metered Connection  
    # Revision v02 - 2022-07-04
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
    # For use with N-able's Cove Data Protection backup solution
    # Must run as administrator or equivalent to use Backup Manager CLI/API 
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Install Module / Check for Metered Connection via https://www.powershellgallery.com/packages/NetMetered/1.0
    # Check for Running Backup Manager (BackupFP.exe)
    # Request Client Authentication Visa 
    # Enable/Disable Backwidth Throttling based on presence of a Metered Connection
    #
    # Use the -Up_Kb parameter to set Throttle upload value in Kilobytes  (32kbps is the minimal allowed setting)
    # Use the -Down_Kb parameter to set Throttle download value in Kilobytes  (32kbps is the minimal allowed setting)  
    # Use the -Throttle_Start parameter to set Throttle start time in 24:00 format
    # Use the -Throttle_Stop parameter to set Throttle stop time in 24:00 format
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/performance.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-guide/command-line.htm
    # https://support.microsoft.com/en-us/windows/metered-connections-in-windows-7b33928f-a144-b265-97b6-f2e95a87c408


# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)]$Up_Kb = "64",
        [Parameter(Mandatory=$False)]$Down_Kb = "64",
        [Parameter(Mandatory=$False)]$Throttle_Start = "0:01",
        [Parameter(Mandatory=$False)]$Throttle_Stop = "23:59"
    ) 

#region ----- Functions ----

Function Get-FPvisa {
    
    $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"
    $config = "C:\Program Files\Backup Manager\config.ini"

    if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { 
        Write-Warning "Backup Manager Not Running"
        Break 
    }else{ 
        try { $ErrorActionPreference = 'Stop'; $UIToken = & $clienttool in-agent-authentication-token.get -config-path $config | convertfrom-json 
        }catch{ 
            Write-Warning "Oops: $_" 
        }
    }

    $UIToken = $UIToken.inagentauthenticationToken

    $url = "http://localhost:5000/jsonrpcv1"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = 'jsonrpc'
    $data.method = 'InAgentAuthenticationTokenLogin'
    $data.params = @{}
    $data.params.token = $UIToken

    $webrequest = Invoke-WebRequest -Method POST `
    -ContentType "application/json; charset=utf-8" `
    -Body (ConvertTo-Json $data -depth 6) `
    -Uri $url `
    -SessionVariable Script:websession `
    -TimeoutSec 180 `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession

    $FPvisa = $webrequest | convertfrom-json
    $script:BackupFPvisa = $FPvisa.result.result
}

Function Set-Throttle {
    Write-Output "`n  Setting Bandwidth Throttle"
    if ($DOWN_KB -eq "Unlimited") {$DOWN_KB = "-1"}
    if ($UP_KB -eq "Unlimited") {$UP_KB = "-1"}

    $url = "http://localhost:5000/jsonrpcv1"
    $script:data = [ordered]@{}
    $data.jsonrpc = '2.0'
    $data.id = 'jsonrpc'
    $data.visa = $Script:BackupFPvisa
    $data.method = 'SaveBandwidthOptions'
    $data.params = @{}
    $data.params.limitBandWidth = $Throttle_Enabled
    $data.params.turnOnAt = $Throttle_Start
    $data.params.turnOffAt = $Throttle_Stop
    $data.params.maxUploadSpeed = [int]$Up_Kb
    $data.params.maxDownloadSpeed = [int]$Down_Kb
    $data.params.dataThroughputUnits = "KBits"
    $data.params.unlimitedDays = @($script:weekends)
    $data.params.pluginsToCancel = @()

    $webrequest = Invoke-WebRequest -Method POST `
    -ContentType "application/json; charset=utf-8" `
    -Body (ConvertTo-Json $data -depth 6) `
    -Uri $url `
    -SessionVariable Script:websession `
    -TimeoutSec 180 `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:SetThrottle = $webrequest | convertfrom-json
    
    [void]::$SetThrottle | convertto-json
    if ($SetThrottle.error) {
        $SetThrottle.error.message
    }else{
        $data.params.GetEnumerator() | Sort-Object -Property value -Descending
    }
}

Function Clear-Throttle {
    Write-Output "`n  Clearing Bandwidth Throttle"

    $url = "http://localhost:5000/jsonrpcv1"
    $script:data = @{}
    $data.jsonrpc = '2.0'
    $data.id = 'jsonrpc'
    $data.visa = $Script:BackupFPvisa
    $data.method = 'SaveBandwidthOptions'
    $data.params = @{}
    $data.params.limitBandWidth = $Throttle_Enabled
   
    $webrequest = Invoke-WebRequest -Method POST `
    -ContentType "application/json; charset=utf-8" `
    -Body (ConvertTo-Json $data -depth 6) `
    -Uri $url `
    -SessionVariable Script:websession `
    -TimeoutSec 180 `
    -UseBasicParsing
    $Script:cookies = $websession.Cookies.GetCookies($url)
    $Script:websession = $websession
    $Script:ClearThrottle = $webrequest | convertfrom-json
    
    [void]::$ClearThrottle | convertto-json
    if ($ClearThrottle.error) {
        $ClearThrottle.error.message
    }else{
        $data.params
    }
}

#endregion ----- Functions ----

if (Get-Module -ListAvailable -Name NetMetered) {
    Write-Host "  Module NetMetered Already Installed"
}else{
    try {
        Install-Module -Name NetMetered -Confirm:$False -Force      ## https://www.powershellgallery.com/packages/NetMetered/1.0
    }
    catch [Exception] {
        $_.message 
        exit
    }
}

If(Test-NetMetered) { 
    Write-Output "  Metered connection/s detected" 
    $Throttle_Enabled="true"
    If ($Unlimited_Weekends -eq "true") {$script:weekends = "Saturday","Sunday"}else{$script:weekends = ""}
    Get-FPvisa
    Set-Throttle
}else{
    Write-Output "  Metered connection/s not detected" 
    $Throttle_Enabled="false"
    Get-FPvisa
    Clear-Throttle
}
