<#
.SYNOPSIS
    Script to install the Cove | Backup Manager using Installation Package Name with optional self-managed encryption key and force uninstallation of prior versions.

.DESCRIPTION
    This script downloads and installs the Backup Manager. It validates the installation package name, optionally uses a self-managed encryption key, and can force the uninstallation of prior versions if specified. The script also checks the status of the Backup Service Controller and BackupFP process after installation.

.PARAMETER installationPackageName
    Mandatory parameter for the installation package name. The name must follow the structure 'cove#v1#<GUID>#.exe'.

.PARAMETER PrivateKey
    Optional parameter for the encryption key if the installation package was configured for it. The key must be at least 8 characters long and contain at least one digit, one uppercase letter, one lowercase letter, and one special character from the set [@#$%*?!;].

.PARAMETER downloadPath
    Optional parameter for the download directory with a default value of "C:\Windows\Temp".

.PARAMETER force
    Optional switch parameter to force uninstallation of prior version if it is already installed.

.EXAMPLE
    .\CDP-Install.v01.ps1 -installationPackageName "cove#v1#9696c2af4-678a-4727-9b6b-example#.exe"
    Installs the Backup Manager with the specified installation package name.

.EXAMPLE
    .\CDP-Install.v01.ps1 -installationPackageName "cove#v1#9696c2af4-678a-4727-9b6b-example#.exe" -force
    Installs the Backup Manager with the specified installation package name and forces uninstallation of any prior version.

.EXAMPLE
    .\CDP-Install.v01.ps1 -installationPackageName "cove#v1#9696c2af4-678a-4727-9b6b-example#.exe" -PrivateKey "P@ssw0rd!1234" -downloadPath "C:\Downloads" -force
    Installs the Backup Manager with the specified installation package name, uses the provided self managd encryption key, sets the download path to "C:\Downloads", and forces uninstallation of any prior version.

.NOTES
    - The script sets the security protocol to TLS 1.2 for downloading the installation package.
    - It checks if the Backup Manager is already installed and uninstalls it if the force switch is used.
    - After installation, it verifies if the Backup Service Controller is running and the BackupFP process is active.
    - If there is an issue, it retrieves the last error entry from the log file located in "C:\ProgramData\mxb\Backup Manager\logs\ClientTool".

.LEGAL
    Sample scripts are not supported under any N-able support program or service.
    The sample scripts are provided AS IS without warranty of any kind.
    N-able expressly disclaims all implied warranties including, warranties
    of merchantability or of fitness for a particular purpose. 
    In no event shall N-able or any other party be liable for damages arising
    out of the use of or inability to use the sample scripts.

COMPATIBILITY
    For use with the Standalone edition of N-able | Cove Data Protection
    Tested with Backup Manager 24.12
#>
param (
    [Parameter(Mandatory = $true)] 
    [string]$installationPackageName,  # Mandatory parameter for the installation package name
    
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        $_.Length -ge 8 -and 
        $_ -match '\d' -and 
        $_ -match '[A-Z]' -and 
        $_ -match '[a-z]' -and 
        $_ -match '[@#$%*?!;]'
    })]
    [string]$PrivateKey,  # Optional parameter for the encryption key if the installation package was configured for it
    
    [Parameter(Mandatory = $false)] 
    [string]$downloadPath = "C:\Windows\Temp",  # Optional parameter for the download directory with a default value
    
    [switch]$force  # Optional switch parameter to force uninstallation of prior version
)

$startTime = Get-Date  # Record the start time of the script

# Check if the force switch is not used and if the Backup Manager is already installed
if (-not $force -and (Test-Path "C:\Program Files\Backup Manager\BackupIP.exe")) { 
    Write-Warning "Prior Backup Manager installation present, use -force switch to uninstall prior version." 
    break  # Exit the script if prior installation is found and force is not used
}

# Validate the installation package name structure
if ($installationPackageName -match '^cove#v1#([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})#\.exe$') {
    Write-Output "Cove Installation package name structure is valid."
} else {
    Write-Warning "Cove Installation package name structure is invalid.`nCopy the full Installation package name from the Cove | Backup.Management | Add Device Wizard.`nIt should look like 'cove#v1#9696c2af4-678a-4727-9b6b-example#.exe'"
    break  # Exit the script if the package name is invalid
}

if ($PrivateKey) {
    Write-Output "Private Key is provided."
    $privatekeystring = "-encryption-key `"$PrivateKey`""  # Prepare the encryption key argument
}

# Create the download path directory if it does not exist
if (-not (Test-Path -Path $downloadPath)) {
    New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
    Write-Output "Download path created: $downloadPath"
}

$installpath = Join-Path -Path $downloadPath -ChildPath $installationPackageName  # Combine the download path with the installation package name

# Set the security protocol to TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Output "Downloading installation package..."

# Download the installation package
$client = New-Object System.Net.WebClient
$client.DownloadFile("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe", $installpath)
Write-Output "Installation package downloaded successfully."

# If force switch is used and Backup Manager is installed, uninstall it
if ($force -and (Test-Path "C:\Program Files\Backup Manager\BackupIP.exe")) {
    Write-Output "Force switch detected. Proceeding with uninstallation of prior Backup Manager..."
    Start-Process -FilePath "C:\Program Files\Backup Manager\BackupIP.exe" -ArgumentList "uninstall -interactive -path `"C:\Program Files\Backup Manager`" -sure" -Wait
    Write-Output "Prior Backup Manager uninstalled."
}

Write-Output "Installing Backup Manager..."
# Install the Backup Manager silently
Start-Process -FilePath $installpath -ArgumentList "-silent $privatekeystring" -Wait -NoNewWindow -PassThru | Out-Null
Start-Sleep -Seconds 5  # Wait for 5 seconds

# Check if the Backup Service Controller is running and BackupFP process is active
$service = Get-Service -Name "Backup Service Controller" -ErrorAction SilentlyContinue
$process = Get-Process -Name "BackupFP" -ErrorAction SilentlyContinue

if ($service.Status -eq 'Running' -and $process) {
    Write-Output "Backup Service Controller is running and BackupFP process is active."
    $ExitCode = 0  # Set exit code to 0 if everything is running fine
} else {
    Write-Warning "Backup Service Controller or BackupFP process is not running."
    $ExitCode = 1  # Set exit code to 1 if there is an issue

    # Retrieve the last error entry from the log file
    $logDirectory = "C:\ProgramData\mxb\Backup Manager\logs\ClientTool"
    $latestLogFile = Get-ChildItem -Path $logDirectory -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($latestLogFile) {
        $logContent = Get-Content -Path $latestLogFile.FullName | Where-Object { $_ -match '\[E\]' }
        $lastErrorEntry = $logContent | Select-Object -Last 1
        Write-Warning "Last Error line from the log file:"
        Write-Warning $lastErrorEntry
    } else {
        Write-Output "No log files found in $logDirectory."
    }
}

$endTime = Get-Date  # Record the end time of the script
Write-Output "Script run duration: $($endTime - $startTime)"  # Output the duration of the script run