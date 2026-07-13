<# ----- About: ----
    # CDP Install Windows - Deploy with Post-Install Bandwidth Throttle
    # Revision v01 - 2026-07-02
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
    # For use with N-able | Cove Data Protection
    # Requires PowerShell 5.1 or higher with administrative privileges.
# -----------------------------------------------------------#>
<# ----- Behavior: ----
    # Downloads and installs Cove Backup Manager (full edition) with a custom
    # Profile and Retention Policy (e.g. 'All-In').
    # Skips install if config.ini already exists (no -Force overwrite).
    # After install, waits for the Backup Service Controller to start (up to 5 min),
    # then applies a single bandwidth throttle via the local JSON-RPC API.
    #
    # Variables to configure:
    #   $CUID                - Customer UID   @ https://backup.management | Customers
    #   $PROFILEID           - Profile ID     @ https://backup.management | Profiles
    #   $RETENTIONPOLICY     - Retention Policy name (case-sensitive, e.g. 'All-In')
    #                          @ https://backup.management | Retention Policies
    #   -ThrottleLimit       : $true = enforce throttle window | $false = always unlimited
    #   -ThrottleOnAt/OffAt  : 24-hour HH:mm window when throttle is enforced
    #   -ThrottleKbUp/KbDn   : max speed in KBits during throttle window (-1 = unlimited, min 128)
    #   -ThrottleExcludeDays : days exempt from throttle (unlimited all day)
    #
    # Use -Uninstall to remove an existing installation instead of deploying.
    #
    # References:
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/silent.htm
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/backup-manager/backup-manager-installation/uninstall-win-silent.htm
# -----------------------------------------------------------#>
<# ----- Exit Codes: ----
    #  0 - Success: installed and throttle applied, or already installed (skipped)
    #  1 - Uninstall: BackupIP.exe not found (Backup Manager not installed)
    #  2 - Download failed, or installer/uninstaller could not start or returned non-zero
    #  3 - Post-install: Backup Service Controller did not start within timeout
    #  4 - Post-install: API port 5000 not available within timeout
    #  5 - Post-install: account directory or auth token not found
    #  6 - Auth: InAgent token login failed or no visa returned
    #  7 - Throttle: SaveBandwidthOptions failed or shell did not initialize
# -----------------------------------------------------------#>

[CmdletBinding()]
Param (
    # ----- Deployment ----
    [string]$CUID            = 'dec8a8eb-9bf6-4a15-afe2-9bb6e5b94aaf',  # Replace with your Customer UID
    [string]$PROFILEID       = '128555',                                  # Replace with your Profile ID #
    [string]$RETENTIONPOLICY = 'All-In',                                 # Replace with your Retention Policy Name (not ID #)

    # ----- Bandwidth Throttle ----
    # KbUp / KbDn: -1 = unlimited, otherwise minimum 128 KBits
    [bool]$ThrottleLimit                                 = $true,
    [string]$ThrottleOnAt                                = '08:00',
    [string]$ThrottleOffAt                               = '18:00',

    [ValidateScript({ $_ -eq -1 -or $_ -ge 128 })]
    [int]$ThrottleKbUp                                   = 512,

    [ValidateScript({ $_ -eq -1 -or $_ -ge 128 })]
    [int]$ThrottleKbDn                                   = -1,

    # Days that are always unlimited regardless of ThrottleLimit/window settings.
    # Examples:
    #   @('Saturday', 'Sunday')                    - exempt weekends (default)
    #   @('Sunday')                                - exempt Sunday only
    #   @('Friday','Saturday','Sunday')            - exempt Fri-Sun
    #   @()                                        - no exempt days, throttle applies every day
    [string[]]$ThrottleExcludeDays                       = @('Saturday', 'Sunday'),

    # ----- Uninstall ----
    [switch]$Uninstall   # Remove existing Backup Manager installation
)

Set-Location -Path $PSScriptRoot

# ----- Helpers ----

function Wait-BmProcess ($proc, $status) {
    for ($i = 0; $i -le 100; $i = ($i + 1) % 100) {
        Write-Progress -Activity 'Cove Backup Manager' -Status $status -PercentComplete $i
        Start-Sleep -Milliseconds 100
        if ($proc.HasExited) { Write-Progress -Activity 'Cove Backup Manager' -Completed; break }
    }
}

$BMConfig = 'C:\Program Files\Backup Manager\config.ini'
$BackupIP = 'C:\Program Files\Backup Manager\BackupIP.exe'
$INSTALL  = 'c:\windows\temp\mxb-windows-x86_x64.exe'
$localUrl = 'http://localhost:5000/jsonrpcv1'

function Invoke-ClientApi ([string]$Method, [object]$Params = @{}, [string]$Visa = '') {
    $body = @{ jsonrpc='2.0'; id=1; method=$Method; visa=$Visa; params=$Params } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Uri $localUrl -Method POST -ContentType 'application/json' -Body $body
}

# ----- Uninstall ----

if ($Uninstall) {
    if (-not (Test-Path $BackupIP -PathType Leaf)) { Write-Warning "BackupIP.exe not found - Backup Manager does not appear to be installed."; exit 1 }
    Write-Host "Uninstalling Backup Manager..."
    Stop-Process -Name 'BackupFP' -Force -ErrorAction SilentlyContinue
    $proc = Start-Process -FilePath $BackupIP -ArgumentList "uninstall -interactive -path `"C:\Program Files\Backup Manager`" -sure" -PassThru
    Wait-BmProcess $proc 'Uninstalling'
    Write-Host "Uninstaller exited with code: $($proc.ExitCode)"; exit $proc.ExitCode
}

# ----- Install ----

if (Test-Path $BMConfig -PathType Leaf) { Write-Host "Existing Backup Manager config.ini found - skipping installation and throttle."; exit 0 }

Write-Host "Downloading Cove Backup Manager..."
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try   { (New-Object System.Net.WebClient).DownloadFile('https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe', $INSTALL) }
catch { Write-Warning "Download failed: $_"; exit 2 }
Stop-Process -Name 'BackupFP' -Force -ErrorAction SilentlyContinue
Write-Host "Installing Backup Manager..."
try   { $proc = Start-Process -FilePath $INSTALL -ArgumentList "-unattended-mode -silent -partner-uid $CUID -profile-id $PROFILEID -product-name `"$RETENTIONPOLICY`"" -PassThru }
catch { Write-Warning "Failed to start installer: $_"; exit 2 }
Wait-BmProcess $proc 'Installing'
Write-Host "Installer exited with code: $($proc.ExitCode)"
if ($proc.ExitCode -ne 0) { Write-Warning "Installer returned non-zero exit code: $($proc.ExitCode)."; exit 2 }

# ----- Wait: Service ----

Write-Host "Waiting for Backup Service Controller..."
$t = 0
while ((Get-Service 'Backup Service Controller' -ErrorAction SilentlyContinue).Status -ne 'Running' -and $t -lt 300) { Start-Sleep -Seconds 10; $t += 10 }
if ($t -ge 300) { Write-Warning "Service did not start within 300s. Throttle not applied."; exit 3 }
Write-Host "  Service: Running"

# ----- Wait: API port 5000 ----

Write-Host "Waiting for API port 5000..."
$t = 0; $apiReady = $false
do {
    try   { $tcp = New-Object System.Net.Sockets.TcpClient; $tcp.Connect('localhost', 5000); $tcp.Close(); $apiReady = $true }
    catch { Write-Host "  Not ready, retrying in 5s... ($t/120)"; Start-Sleep -Seconds 5; $t += 5 }
} until ($apiReady -or $t -ge 120)
if (-not $apiReady) { Write-Warning "API port 5000 not available within 120s. Throttle not applied."; exit 4 }
Write-Host "  Port 5000 ready."

# ----- Authenticate (InAgent token -> visa) ----

$storageRoot = 'C:\ProgramData\MXB\Backup Manager\storage'
$t = 0; $acctDir = $null
Write-Host "Locating account token file..."
do {
    $acctDir = (Get-ChildItem $storageRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne 'certificates' } |
        ForEach-Object {
            $tf = Join-Path $_.FullName 'in_agent_authentication_token'
            if (Test-Path $tf) { [PSCustomObject]@{ Dir = $_.FullName; T = (Get-Item $tf).LastWriteTime } }
        } | Sort-Object T -Descending | Select-Object -First 1).Dir
    if (-not $acctDir) { Write-Host "  Token not yet available, retrying in 10s... ($t/120)"; Start-Sleep -Seconds 10; $t += 10 }
} until ($acctDir -or $t -ge 120)
if (-not $acctDir) { Write-Warning "Could not locate account directory. Throttle not applied."; exit 5 }
Write-Host "  Account dir: $(Split-Path $acctDir -Leaf)"

$ct = 'C:\Program Files\Backup Manager\ClientTool.exe'
$ctOut = & $ct 'in-agent-authentication-token.get' '-config-path' $acctDir 2>&1
$token = if ($LASTEXITCODE -eq 0 -and $ctOut) { $ctOut.Trim() } else {
    $tf = Join-Path $acctDir 'in_agent_authentication_token'
    if (Test-Path $tf) { (Get-Content $tf).Trim() }
}
if (-not $token) { Write-Warning "Could not retrieve InAgent auth token. Throttle not applied."; exit 5 }

try   { $loginResp = Invoke-ClientApi -Method 'InAgentAuthenticationTokenLogin' -Params @{ token = $token } }
catch { Write-Warning "Login request failed: $_. Throttle not applied."; exit 6 }
if (-not $loginResp -or $loginResp.error) { Write-Warning "Login failed: $($loginResp.error.message). Throttle not applied."; exit 6 }
$visa = if ($loginResp.result.PSObject.Properties['result']) { $loginResp.result.result } else { $loginResp.result }
if (-not $visa) { Write-Warning "No visa returned. Throttle not applied."; exit 6 }
Write-Host "Authenticated via InAgent token."

# ----- Apply Bandwidth Throttle ----

Write-Host "Applying throttle: limit=$ThrottleLimit  On=$ThrottleOnAt  Off=$ThrottleOffAt  Up=${ThrottleKbUp}Kbps  Dn=${ThrottleKbDn}Kbps"
$throttleParams = @{
    limitBandWidth=$ThrottleLimit; turnOnAt=$ThrottleOnAt; turnOffAt=$ThrottleOffAt
    maxUploadSpeed=$ThrottleKbUp; maxDownloadSpeed=$ThrottleKbDn; dataThroughputUnits='KBits'
    unlimitedDays=$ThrottleExcludeDays; pluginsToCancel=@()
}
try {
    $t = 0; $done = $false
    do {
        $r = Invoke-ClientApi -Method 'SaveBandwidthOptions' -Params $throttleParams -Visa $visa
        if     ($r.error -and $r.error.message -match 'not initialized') { Write-Host "  Shell not ready, retrying in 10s... ($t/120)"; Start-Sleep -Seconds 10; $t += 10 }
        elseif ($r.error) { throw $r.error.message }
        else   { $done = $true }
    } until ($done -or $t -ge 120)
    if (-not $done) { Write-Warning "Shell did not initialize within 120s. Throttle not applied."; exit 7 }
    Write-Host "Throttle applied successfully."
} catch { Write-Warning "Failed to apply throttle: $_"; exit 7 }
exit 0
