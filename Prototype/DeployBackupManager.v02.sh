#!/bin/bash

# ----- About: ----
    # AutoDeploy Backup Manager for Linux
    # Revision v02 - 2021-03-20
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
# -----------------------------------------------------------#>  ## About

# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

# ----- Compatibility: ----
    # For use with the standalone edition of N-able Backup
    
# -----------------------------------------------------------#>  ## Compatibility

# ----- Behavior: ----
    # Determine 32/64-bit
    # Download Backup Manager .run installer
    # AutoDeploy to customer specifed by UID
    # Optionally set device Profile
    #
    # Use the -u parameter to set the Customer UID
    # Use the -p parameter to set the device Profile
    # Use the -h parameter for syntax help
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/auto-deployment.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-installation/non-ui-installer.htm
# -----------------------------------------------------------#  ## Behavior



COMMAND_SYNTAX='
N-able Backup
DeployBackupManager.sh 
Command line options:
  -u     Customer UID (36 chars) from https://backup.management | Customer management
  -p	 ProfileId from https://backup.management/#/profiles (0 for No Profile)

  example: # bash ./DeployBackupManager.v02.sh -u 6079722f-replace-me-b991-aa57f4773b21 -p 123456
'

while getopts :u:p:h flag
do
	case "${flag}" in
		u) CUSTOMERUID=${OPTARG};;
		p) DEVICEPROFILE=${OPTARG};;
		h) echo "$COMMAND_SYNTAX"
		   exit ;;		
	esac
done

#echo ${#CUSTOMERUID}
#echo $CUSTOMERUID
#echo $DEVICEPROFILE

if [ ${#CUSTOMERUID} -ne 36 ]; then echo "Invalid UID Length" ; exit
else echo "Correct UID Length"
fi

BIT=$(getconf LONG_BIT)
#echo $BIT
if [[ $BIT == *64* ]]; then echo "Downloading 64-Bit Installer"; runfile='mxb-linux-x86_64.run'
else echo "Downloading 32-Bit Installer"; runfile='mxb-linux-i686.run'
fi
#echo $runfile

filename="swibm#$CUSTOMERUID#$DEVICEPROFILE#.run"
#echo $filename
curl -o $filename https://cdn.cloudbackup.management/maxdownloads/$runfile

chmod +x $filename
./$filename


