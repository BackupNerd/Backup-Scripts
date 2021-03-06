<# ----- About: ----
    # SolarWinds MSP | N-able Backup Set Active Hours Bandwidth Throttle
    # Revision v01 - 2021-03-06
    # Author: Eric Harless, Head Backup Nerd - SolarWinds MSP | N-able
    # Twitter @Backup_Nerd  Email:eric.harless@solarwinds.com
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any SolarWinds support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # SolarWinds expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall SolarWinds or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Set Backup Bandwidth Throttle Window using 
    # Active Hours settings from Reg Key HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings
    # Adjust Kb Upload/ Download limits 
    # https://docs.microsoft.com/en-us/windows/deployment/update/waas-restart#configure-active-hours
# -----------------------------------------------------------#>

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [string]$upload = "128",                          ## Throttle upload limit in Kb
        [Parameter(Mandatory=$False)] [string]$download = "4096"                        ## Throttle download limit in Kb
    )

    Clear-Host
    $Script:ActiveHours = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings  ## Get Active Hours Settings from Reg Key

    [String]$Par0 = ":00"                                                               ## Minutes
    [String]$Active = "true"                                                            ## true = throttle enabled (case sensitive)
    [String]$Start = [string]$($Script:ActiveHours.ActiveHoursStart) + [string]$Par0    ## Get Throttle Start Hour from Reg Key
    [String]$Stop = [string]$($Script:ActiveHours.ActiveHoursEnd) + [string]$Par0       ## Get Throttle Stop Hour from Reg Key

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $body = "{`"id`":`"jsonrpc`",`"jsonrpc`": `"2.0`",`"method`":`"SaveBandwidthOptions`",`"params`": {`"limitBandWidth`":$Active ,`"turnOnAt`":`"$Start`",`"turnOffAt`":`"$Stop`",`"maxUploadSpeed`":$upload,`"maxDownloadSpeed`":$download,`"dataThroughputUnits`":`"KBits`",`"unlimitedDays`":[],`"pluginsToCancel`":[]} }"
    $response = Invoke-RestMethod 'http://localhost:5000/jsonrpcv1' -Method 'POST' -Headers $headers -Body $body

    [void]::$response | convertto-json
    if ($response.error) {$response.error.message}
    else {
        $val = $body | convertfrom-json
        $val.method
        $val.params
    }
