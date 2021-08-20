# ----- About: ----
    # N-able Backup Get Client Errors
    # Revision v01 - 2021-08-20
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Reddit https://www.reddit.com/r/Nable/
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
    # For use with the Standalone and integrated editions of N-able Backup
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Sample scripts to use the Clienttool command line utility to pull errors and statistics
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-guide/command-line.htm

# -----------------------------------------------------------#>  ## Behavior
clear-host

$counter = 0

Do {

    $BackupService = get-service "Backup Service Controller" -ea SilentlyContinue
    if ($BackupService.status -ne "Running"){ 
        Write-output "ERROR: Backup Service Not Running"
        Write-output "`n-------- $Counter --------`n"
        $counter ++
        Start-Sleep -seconds 5 
    }
    else{   
        Write-output "`n-------- $Counter --------`n"
    
    # Get Backup Manager Initialization Errors
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $initerror = & "C:\Program Files\Backup Manager\ClientTool.exe" control.initialization-error.get  | convertfrom-json}catch{ Write-Warning "ERROR: $_`n" }}
        if ($initerror.code -gt 0) {write-output "ERROR: $($initerror.Message)`n"}else{ "Cloud Initialized`n"}

    # Get Backup Manager Application Status
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $AppStatus = & "C:\Program Files\Backup Manager\ClientTool.exe" control.application-status.get }catch{ Write-Warning "ERROR: $_" }}
        Write-output "Application Status:  $appstatus`n"

    # Get Backup Manager Job Status
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $JobStatus = & "C:\Program Files\Backup Manager\ClientTool.exe" control.status.get }catch{ Write-Warning "ERROR: $_" }}
        Write-output "Job Status:  $jobStatus`n"
        
    # Test VSS
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" vss.check }catch{ Write-Warning "Oops: $_" }}
        
    # Test Storage Node Connectivity
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" storage.test }catch{ Write-Warning "Oops: $_" }}
        Write-output ""

    # Get Settings
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" control.setting.list }catch{ Write-Warning "Oops: $_" }}
        Write-output ""
        Start-Sleep -seconds 2 

    # Get Selections
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" control.selection.list }catch{ Write-Warning "Oops: $_" }}
        Write-output ""
        Start-Sleep -seconds 2

    # Get Schedules
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" control.schedule.list }catch{ Write-Warning "Oops: $_" }}
        Write-output ""
        Start-Sleep -seconds 2

    # Get Archive Schedules
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" control.archiving.list }catch{ Write-Warning "Oops: $_" }}
        Write-output ""
        Start-Sleep -seconds 2

    # Get Selected Datasources via Selection
        if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $Datasources = & "C:\Program Files\Backup Manager\ClientTool.exe" control.selection.list  | ConvertFrom-String | select -skip 2 | ForEach {If ($_.P2 -eq "Inclusive") {echo $_.P1}} }catch{ }}
        
    # Get Errors for each selected datasource
        foreach ($datasource in  $Datasources | select-object -unique ) {
            if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ "`n$datasource"; try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\Backup Manager\ClientTool.exe" control.session.error.list -datasource $datasource -no-header}catch{ }}
            Start-Sleep -seconds 5
            }

        $counter ++
        Start-Sleep -seconds 5
 }

}until ($counter -ge 500)

