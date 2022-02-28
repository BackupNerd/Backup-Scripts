# ----- About: ----
    # N-able Backup Get-Failures
    # Revision v02 - 2022-02-08
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
    # For use with the Standalone and N-Central integrated editions of N-able Backup
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Read local sessions.xml file
    # Output failed jobs over last X Days or X Sessions
    # Uses Clienttool.exe to output detailed errors for each failed eession timestamp with an error count
    # 
    # Use the -Count parameter to specify the # of Days or Sessions to look back
    # Use the -Units parameter to specify Days or Sessions
    # Use the -MaxErrors parameter to specify the maximum number of errors to return per failed session
    #
    # ./Get-Failures
    # ./Get-Failures 14
    # ./Get-Failures 21 Days
    # ./Get-Failures 30 Sessions
    # ./Get-Failures 30 Sessions 25
    # ./Get-Failures -count 30 -units Sessions -maxerrors 25
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/backup-manager/backup-manager-guide/command-line.htm
    # https://documentation.n-able.com/backup/troubleshooting/Content/kb/MSP-How-can-I-generate-the-whole-list-of-backup-sessions-on-the-device.htm
    
# -----------------------------------------------------------#>  ## Behavior
[CmdletBinding()]
param (
[Parameter(Mandatory = $false)] [Int]$Count = 7, ## Unit count
[Parameter(Mandatory = $false)] [ValidateSet("Days","Sessions")] [string]$Units = "Days", ## Unit of measure
[Parameter(Mandatory = $false)] [Int]$MaxErrors = 100 ## Maximum errors per session
)

clear-host
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Push-Location $dir

$Success = @("InProcess","Completed")  ## Successful status times to ignore

[xml]$sessions = get-content -path "C:\ProgramData\MXB\Backup Manager\SessionReport.xml"  ## Session Report to parse

if ($units -eq "Days") {
    [datetime]$Start = (get-date).AddDays($Count/-1)
    $History = $sessions.SessionStatistics.session | sort-object -Descending Starttimeutc | Where-Object { ($_.Status -notin $Success) -and ($start -lt $_.starttimeutc)}
} ## Process output for Days

if ($units -eq "Sessions") {
    $History = $sessions.SessionStatistics.session[0..$count] | sort-object -Descending Starttimeutc | Where-Object {$_.Status -notin $Success}
} ## Process output for Sessions

if ($History) {$History | select-object Type,Plugin,starttimeutc,selectedSize,SelectedCount,errorscount,status | Format-Table }else{ Write-Output " No Backup Session Failures Found in Last $count $units"}

## Get errors for each unsuccessful Backup Session with an error count
$global:errors = @()

#$global:errors+= New-Object -TypeName PSObject -Property @{ StartTimeUTC = $null; Path = $null;Detail = $null}

ForEach ($Failed in $history | where-object {($_.errorscount -ne "0") -and ($_.type -eq "Backup")}) {
    # Get Selected Datasources via Selection
    if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $Datasources = & "C:\Program Files\Backup Manager\ClientTool.exe" control.selection.list | ConvertFrom-String | select -skip 2 | ForEach {If ($_.P2 -eq "Inclusive") {echo $_.P1}} }catch{ }}


    
    # Get Errors for each selected datasource
        foreach ($datasource in  $Datasources | select-object -unique ) {
            if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "Backup Manager Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $errormsg = & "C:\Program Files\Backup Manager\ClientTool.exe" control.session.error.list -datasource $datasource -time "$($failed.starttimeutc)" -limit $maxerrors | ConvertFrom-String }catch{ }}
            
            if ($errormsg -ne "No session errors found.") {$global:errors += $errormsg; write-host "Getting Error Details"}

            }
        }
        $errors | Where-Object {$_.p1 -notlike "*-------*"} | ft

