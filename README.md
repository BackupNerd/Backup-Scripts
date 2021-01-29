
# READ ME

## TO DOWNLOAD FILES click the green ( CODE button ) to download the entire Script Repository as a single .ZIP file.

### Sample scripts are not supported under any SolarWinds support program or service.
### The sample scripts are provided AS IS without warranty of any kind.
### SolarWinds expressly disclaims all implied warranties including, warranties of merchantability or of fitness for a particular purpose. 
### In no event shall SolarWinds or any other party be liable for damages arising out of the use of or inability to use the sample scripts.

### Please review, test and understand any scripts you download before implementing them in production!

# Non Backup Specific ###

Get-Foldersizes.ps1

	Script to inventory a file tree and report sizes of each branch.

# SolarWinds Backup - Standalone Editon ###

BulkGenerateRedeployCommands.ps1

	Script to bulk generate a list of device credentials and redeployment commands for SolarWinds Backup devices.

BulkSetGUIPassword.ps1

	Script to bulk set/wipe a GUI Password that is useful forlimiting local and remote access to the Backup Manager client.

CleanupArchive.ps1

	Script to clean all archive sessions older than X months. 
	Optionally, check for existing Archive rules and/or create new Archive rules if no active archive rules are found.

CustomBackupThrottle.ps1
	
	Script to set backup throttling times and values for individual days of the week   

DeployBackupManager.ps1

	Universal deployment script for SolarWinds Backup
	Supports Automatic Deployment, Documents Deployment, Manual Deployment, Uninstall, Upgrade, Redeploy with Passphrase,
	Redeploy with Private Key, Store Credentials, Reinstall & Reuse Stored Credentials.

ExcludeUSB.ps1
	
	Script to exclude USB attached disks from backup.
	Identify disks attached via USB bus type and exclude those device volumes using backup filter or the FilesNotToBackup registry key.

GetDeviceErrors.ps1

	Script to pull recent Device Errors and output to the console or update a Custom Column in the Management Console
        
GetDeviceInstallations.ps1

	Script to Enumerate Device Installations.
	Incudes all Historic installation instances of a Device, including Backup, Restore Only, Bare-Metal Recovery, Recovery Console and Recovery Testing instances
        Useful for auditing the last activity date for a specific Installation Id.
	
GetDeviceStatistics.ps1

	Script to Enumerate Device Statistics.
	Predefined to Include Multiple Columns from the Management Console, can be expanded to include custom columns.
        Useful for automating an export of data for billing or reporting.

GetM365DeviceStatistics.ps1

	Script to Enumerate Microsoft 365 Device Statistics.
	Predefined to Include Mail, OnDrive and Sharepoint Statistics and Protected Users.
        Useful for automating an export of data for billing or reporting.
	
GetSessionFiles.ps1	

	Script to Enumerate Files from the most recent Backup Sessions.
	Incudes x largest transfered files, Hyper-V Files and Hyper-V Files incorectly included in File System Backups 
        Useful for Auding selections and ensuring that data is not double selected.

LSVSyncCheckFinal.ps1

	Script to check the current Status and Sync percentage of the LocalSpeedVault and Cloud Storage.

SetBackupLogging.ps1

	Script to set the logging level of the local Backup Manager client.
	
SetDeviceAlias.ps1
	
	Script to set a value for the Device Alias column in the Backup Management Console.
	This script is for Automatic Deployed, Passphrase enabled devices.
	Note, Using this script with a Private Key Encryption device will convert it to Passphrase Encryption.
	
SetDeviceProduct.ps1
	
	Script to bulk assign a Product to multiple devices.
	Useful for bulk assignment of a Product to devices without modifying their current parent partner location.
	
SetDeviceProfile.ps1
	
	Script to bulk assign a Profile to multiple devices.
	Useful for bulk removal, assignment or reassignment of a Profile to devices outside of the Management Console
	
SW MSP_MSPBenhancedMonitoringDeploymentConfiguration_01jul2020jr1.pdf

MSPB_CFG_CHECK_v1.amp

MSPB_CLOUD_CHECK_v1.amp

MSPB_LSV_CHECK_v1.amp

	3 AMP files and PDF documentation to enhance monitoring of SolarWinds Backup (Ncentral Integrated or Standalone Editions).
	
SolarWinds Backup - RMM auto deploy setup.pdf

SolarWinds Backup Deploy.amp

	1 AMP file and PDF documentation to enable RMM deployment of SolarWinds Backup (Standalone Edition)

SolarWinds Backup Redeploy.amp

	AMP file to redeploy SolarWinds Backup (Ncentral Integrated or Standalone Editions) 
	using device credentials and passphase 

SolarWinds Backup Store Credentials.amp

SolarWinds Backup Reuse Stored Credentials.amp

	2 AMP files for use with SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition, when deployed via Ncentral or RMM).
	Store Credentials will store your local Backup Manager credentials 
	Reuse Stored Credentials will redeploy SolarWinds Backup using stored credentials if the Backup Manager gets uninstalled. 

SolarWinds Backup Upgrade.amp

	AMP file to force upgrade of the Backup Manager to the latest downloadable version of SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition).
	TLS 1.2 Compatible
	
SolarWinds Backup Convert To PassPhrase.amp

	AMP file to convert Backup Manager devices with Private Key Encryption to PassPhrase Encryption (Standalone Edition Managed by Ncentral or RMM)

# SolarWinds Backup - Ncentral Integrated Editon ###

CleanupArchive.ps1

	Script to clean all archive sessions older than X months. 
	Optionally, check for existing Archive rules and/or create new Archive rules if no active archive rules are found.

SW MSP_MSPBenhancedMonitoringDeploymentConfiguration_01jul2020jr1.pdf

MSPB_CFG_CHECK_v1.amp

MSPB_CLOUD_CHECK_v1.amp

MSPB_LSV_CHECK_v1.amp

	3 AMP files and PDF documentation to enhance monitoring of SolarWinds Backup (Ncentral Integrated or Standalone Editions).

SolarWinds Backup Migration Prep.amp

SolarWinds Backup Migration Cleanup.amp

	2 AMP files for use when working with SolarWinds on an approved migration from Ncentral Integrated Backup to the Standalone Edition.
	Migration Prep will store your local Backup Manager credentials and block Ncentral from being able to uninstall the Backup Manager. 
	Migration Cleanup will revert changes made to block Ncentral from being able to uninstall the Backup Manager. 

SolarWinds Backup Store Credentials.amp

SolarWinds Backup Reuse Stored Credentials.amp

	2 AMP files for use with SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition, when deployed via Ncentral or RMM).
	Store Credentials will store your local Backup Manager credentials.
	Reuse Stored Credentials will redeploy SolarWinds Backup using stored credentials if the Backup Manager gets uninstalled. 

SolarWinds Backup Upgrade.amp

	AMP file to force upgrade of the Backup Manager to the latest download version of SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition).

# SolarWinds Backup - RMM Integrated Editon ###

CleanupArchive.ps1

	Script to clean all archive sessions older than X months. 
	Optionally, check for existing Archive rules and/or create new Archive rules if no active archive rules are found.

CustomBackupThrottle.ps1
	
	Script to set backup throttling times and values for individual days of the week   

ExcludeUSB.ps1
	
	Script to exclude USB attached disks from backup.
	Identify disks attached via USB bus type and exclude those device volumes using backup filter or the FilesNotToBackup registry key.

LSVSyncCheckFinal.ps1

	Script to check the current Status and Sync percentage of the LocalSpeedVault and Cloud Storage.

SetBackupLogging.ps1

	Script to set the logging level of the local Backup Manager client.






