<# ----- About: ----
    # N-able | Cove Data Protection - Windows Automatic Deployment with Bandwidth Throttle
    # Revision v05 - 2025-07-15
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
# -----------------------------------------------------------#>
<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able | Cove Data Protection
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Downloads and deploys a new Backup Manager as a Passphrase compatible device with an assigned Profile
    # Replace UID and PROFILEID variables at the begining of the script
    # Optionally specify a Retention Policy Name to override the Default Retention Policy
    # Run this Script from the TakeControl Shell or PowerShell 
    #
    # Name: CUID
    # Type: String Variable 
    # Value: 9696c2af4-678a-4727-9b6b-example
    # Note: Found @ Backup.Management | Customers
    #
    # Name: PROFILEID
    # Type: Integer Variable 
    # Value: ProfileID #
    # Note: Found @ Backup.Management | Profiles (use 0 for No Profile)
    #
    # Name: RETENTIONPOLICY
    # Type: Case Sensitive String Variable
    # Value: Retention Policy Name
    # Note: Found @ Backup.Management | Retention Policy (Not specifying a Retention Policy will use the default retention policy)
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/regular-install.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/reinstallation.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/uninstall-win-silent.htm
# -----------------------------------------------------------#>

# Begin Install Script

$CUID =             '01d10b9ee-2a24-4868-9ceb-example'
$PROFILEID =        '128555'
$RETENTIONPOLICY =  '30 Days'
$INSTALL =          "c:\windows\temp\bm#$CUID#$PROFILEID#.exe"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe","$($INSTALL)")

& $INSTALL -product-name `"$RETENTIONPOLICY`"

Start-Sleep -Seconds 180

# End Install Script

# Begin Set-Bandwidth
Function Get-FPvisa {
    $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"
    $config = "C:\Program Files\Backup Manager\config.ini"

    if (-not (Get-Process "BackupFP" -ErrorAction SilentlyContinue)) {
        Write-Warning "Backup Manager Not Running"
        return
    }

    try {
        $ErrorActionPreference = 'Stop'
        $UIToken = & $clienttool in-agent-authentication-token.get -config-path $config | ConvertFrom-Json
        $UIToken = $UIToken.inagentauthenticationToken
    } catch {
        Write-Warning "Error retrieving authentication token: $_"
        return
    }

    $url = "http://localhost:5000/jsonrpcv1"
    $data = @{
        jsonrpc = '2.0'
        id = 'jsonrpc'
        method = 'InAgentAuthenticationTokenLogin'
        params = @{
            token = $UIToken
        }
    }

    try {
        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType "application/json; charset=utf-8" `
            -Body (ConvertTo-Json $data -Depth 6) `
            -Uri $url `
            -SessionVariable Script:websession `
            -TimeoutSec 180 `
            -UseBasicParsing

        $Script:cookies = $websession.Cookies.GetCookies($url)
        $Script:websession = $websession

        $FPvisa = $webrequest | ConvertFrom-Json
        $script:BackupFPvisa = $FPvisa.result.result
    } catch {
        Write-Warning "Error during web request: $_"
    }
} ## Get Local Authentication Token from ClientTool.exe
Get-FPvisa

Function Set-Bandwidth {
    param(
        [String]$Active = "true",    # true = throttle enabled (case sensitive)
        [String]$Start = "08:00",     # Throttle Start Time (HH:mm format)
        [String]$Stop = "17:00",     # Throttle Stop Time (HH:mm format)
        [String]$download = "-1",    # Throttle Download Speed (-1 = Unlimited)
        [String]$upload = "5120"     # Throttle Upload Speed (in Kbits)
    )

    Start-Sleep 15
    Write-Output "`n  Setting Bandwidth Throttle"

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")

    $body = @{
        id = "jsonrpc"
        jsonrpc = "2.0"
        method = "SaveBandwidthOptions"
        visa = $script:BackupFPvisa
        params = @{
            limitBandWidth = $Active
            turnOnAt = $Start
            turnOffAt = $Stop
            maxUploadSpeed = [int]$upload
            maxDownloadSpeed = [int]$download
            dataThroughputUnits = "KBits"
            unlimitedDays = @("Saturday", "Sunday")
            pluginsToCancel = @()
        }
    } | ConvertTo-Json -Depth 10 -Compress

    $response = Invoke-RestMethod 'http://localhost:5000/jsonrpcv1' -Method 'POST' -Headers $headers -Body $body

    [void]::$response | convertto-json
    if ($response.error) {$response.error.message}
    else {
        $val = $body | convertfrom-json
        $val.method
        $val.params
    }
} ## Set Bandwidth Throttle

# Example 1: Standard throttle during work hours
# Set-Bandwidth -Active "true" -Start "08:00" -Stop "17:00" -download "-1" -upload "5120"
# Example 2: Disable throttle
# Set-Bandwidth -Active "false"

Set-Bandwidth

# End Set-Bandwidth