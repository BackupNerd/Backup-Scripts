# Backup-Scripts
## READ ME 

### Sample scripts are not supported under any SolarWinds support program or service.
### The sample scripts are provided AS IS without warranty of any kind.
### SolarWinds expressly disclaims all implied warranties including, warranties of merchantability or of fitness for a particular purpose. 
### In no event shall SolarWinds or any other party be liable for damages arising out of the use of or inability to use the sample scripts.

### Please review, test and understand any scripts you download before implementing them in production!

## TO DOWNLOAD FILES click the green ( CODE button ) to download the entire Script Repository as a single .ZIP file.

### Non Backup Specific ###

Get-Foldersizes.ps1

	Script to inventory a file tree and report sizes of each branch.

### SolarWinds Backup - Standalone Editon ###

BackupManagerDeploy.ps1

	Universal deployment script for SolarWinds Backup
	Supports Automatic Deployment, Documents Deployment, Manual Deployment, Uninstall, Upgrade, Redeploy with Passphrase,
	Redeploy with Private Key, Store Credentials, Reinstall & Reuse Stored Crendentials.

BulkGenerateBackupManagerRedeployCommands.ps1

	Script to bulk generate a list of device credentials and redeployment commands for SolarWinds Backup devices.

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

SetDeviceAlias.ps1
	
	Script to set a value for the Device Alias column in the Backup Management Console.
	This script is for Automatic Deployed, Passphrase enabled devices.
	Note, Using this script with a Private Key Encrpytion device will convert it to Passphrase encryption.

SW MSP_MSPBenhancedMonitoringDeploymentConfiguration_01jul2020jr1.pdf

MSPB_CFG_CHECK_v1.amp

MSPB_CLOUD_CHECK_v1.amp

MSPB_LSV_CHECK_v1.amp

	3 AMP files and PDF documentation to enhance monitoring of SolarWinds Backup (Ncentral Integrated or Standalone Editions).
	
SolarWinds Backup - RMM auto deploy setup.pdf

SolarWinds Backup Deploy.amp

	1 AMP file and PDF documentation to enable RMM deployment of SolarWinds Backup - Standalone Edition

SolarWinds Backup Redeploy.amp

	AMP file to redeploy SolarWinds Backup (Ncentral Integrated or Standalone Editions) 
	using device credentials and passphase 

SolarWinds Backup Store Credentials.amp

SolarWinds Backup Reuse Stored Credentials.amp

	2 AMP files for use with SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition, when deployed via Ncentral or RMM).
	Store Credentials will store your local Backup Manager credentials 
	Reuse Stored Credentials will redeploy SolarWinds Backup using stored credentials if the Backup Manager gets uninstalled. 

SolarWinds Backup Upgrade.v02.amp

	AMP file to force upgrade of the Backup Manager to the latest download version of SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition).

### SolarWinds Backup - Ncentral Integrated Editon ###

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

SolarWinds Backup Upgrade.v02.amp

	AMP file to force upgrade of the Backup Manager to the latest download version of SolarWinds Backup (Ncentral Integrated Edition or Standalone Edition).

### SolarWinds Backup - RMM Integrated Editon ###

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






