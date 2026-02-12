<# ----- About: ----
    # Cove Data Protection Agent Deployment for Windows (NinjaOne)
    # Revision v26.2
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
    # Tested with release 25.10
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Downloads the latest version of the Cove Backup Manager
	# Save the download in C:\Users\Public\Downloads\
	# Silently installs the Backup Manager
	# Adds the new device to the appropriate Cove customer with optional profile and retention policy
	# Removes the downloaded installer after installation
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/external-cove-integrations/ninjaOne/NinjaOne-generate-deployment-script.htm

# -----------------------------------------------------------#>
# Begin Install Script Cove Data Protection deployment

$CoveInstallationID = Ninja-Property-Get 'CoveInstallationID'

# Validate the property
if ([string]::IsNullOrWhiteSpace($CoveInstallationID)) {
    Write-Output "InstallStatus=PropertyError"
    Write-Output "ErrorMessage=CoveInstallationID is not set."
    exit 1
}

# Regex check: enforce GUID/UID format
if ($CoveInstallationID -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
    Write-Output "InstallStatus=PropertyError"
    Write-Output "ErrorMessage=CoveInstallationID has invalid format: $CoveInstallationID"
    exit 1
}

# --- Stop script if Backup Manager is already installed (check for config.ini) ---
if (Test-Path "C:\Program Files\Backup Manager\config.ini") {
    Write-Output "InstallStatus=AlreadyInstalled"
    Write-Output "ErrorMessage=Found config.ini in C:\Program Files\Backup Manager"
    exit 1
}
elseif (Test-Path "C:\Program Files (x86)\Backup Manager\config.ini") {
    Write-Output "InstallStatus=AlreadyInstalled"
    Write-Output "ErrorMessage=Found config.ini in C:\Program Files (x86)\Backup Manager"
    exit 1
}

# Continue with installation
$INSTALL = "C:\Users\Public\Downloads\cove#v1#$CoveInstallationID.exe"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Try HTTPS first, fallback to HTTP
try {
    (New-Object System.Net.WebClient).DownloadFile(
	"https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe",
	$INSTALL
    )
    Write-Host "Download succeeded via HTTPS."
}
catch {
    Write-Warning "HTTPS download failed. Attempting HTTP..."
    try {
	(New-Object System.Net.WebClient).DownloadFile(
	    "http://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe",
	    $INSTALL
	)
	Write-Host "Download succeeded via HTTP."
    }
    catch {
	Write-Output "InstallStatus=DownloadFailed"
	Write-Output "ErrorMessage=Installer could not be downloaded via HTTPS/HTTP."
	exit 1
    }
}

Write-Output "Running installer..."
Start-Process -FilePath $INSTALL -ArgumentList "/S" -Wait
Remove-Item $INSTALL -Force

# End Install Script Cove Data Protection deployment
