clear-host 

    $RecoveryService = get-service "Recovery Console Service" -ea SilentlyContinue
    if ($RecoveryService.status -ne "Running"){ 
        Write-output "ERROR: Recovery Console Service Not Running"
        } ## Check if Recovery Console Service is running
        Else {
          if ((get-process "RecoveryConsole" -ea SilentlyContinue) -eq $Null) { "Recovery Console Not Running" }else{ try { $ErrorActionPreference = 'Stop'; $RCversion = & "C:\Program Files\RecoveryConsole\ClientTool.exe" -version }catch{ Write-Warning "ERROR: $_" }}
          $RCversion.split("`n").replace("Client Tool, v","Recovery Console V")[0]
          Write-Host ""
          }  ## Get RC\ClientTool Version

$ConfigPath = "$env:USERPROFILE\appdata\local\temp\"

    Write-Output "Scanning Restore Sessions"

do {

    $Script:Folders = $null 
    $Script:Folders = Get-Item -path $configpath\* -exclude mnt* | sort LastWriteTime -Descending
    $Script:Configs = $null 
    $Script:Configs = Get-ChildItem -path $Folders -filter con*.tmp -Recurse | sort LastWriteTime -Descending
    $configs
    #clear-host
    #if ($configs -eq $null) { Write-Output "No Running Restore Sessions"}
  
    Start-sleep -seconds 5
    } While ($configs -eq $null)



        foreach ($config in $configs) {

            $getPort = get-content (join-path $config.directory $config.name) | select-string -pattern "HttpServerPort=" 

            $port = $getport -replace("HttpServerPort=","")



            #$HTTP_Status = $null
        
            # First we create the request.
            $HTTP_Request = [System.Net.WebRequest]::Create("http://localhost:$port")

            Write-output "`nChecking http://localhost:$port"
            if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port control.initialization-error.get }catch{ Write-Warning "ERROR: $_" }}
            Write-output ""

        
            # We then get a response from the site.
            try { 
                $ErrorActionPreference = 'Stop'
                $HTTP_Response = $null
                $HTTP_Response = $HTTP_Request.GetResponse() 
                }  ##End Try
            catch { 
                Write-Warning "Restore Session Closed \ Completed `nRemoving Temp Config File"
                Write-Host ""
                remove-item (join-path $config.directory $config.name)
                }  ## End Catch
        
            If ($HTTP_Response.StatusCode -eq "OK") {
                Write-Host "Restore Session Found `nAttempting Connection"
                start-process $HTTP_Response.ResponseUri

                
                    $params = @{
                    Uri         = "$($HTTP_Response.ResponseUri)content/data/progress-status"
                    Method      = 'GET'
                    ContentType = 'application/json; charset=utf-8'
                    
                    }


                DO {

                $Details = Invoke-RestMethod @params
                Write-output "Service Initializing, Please Wait`n"
                Start-sleep -seconds 5
                }While ($details -eq "refresh")

                $Details = Invoke-RestMethod @params
                $Details1 = $Details.replace("[ESCAPE[]]","")
                
                $Details2 = $Details1 -split("<.*>")
                $Details3 = $Details2 -replace('\s(\w+\D):','"$1":')

                $A= $Details3[2] -replace ("\],\s*\n\}","] }") -replace ("\},\s\n\s*\]","}]") | ConvertFrom-Json 
                
                #write-output $a.InitializationError
                write-output $a.device
                write-output $a.LocalSpeedVault
                
                # write-output $a.activity | Select-Object * -ExcludeProperty SessionsinProgress
                $activity = $a.activity | Select-Object * -ExcludeProperty SessionsinProgress
                write-output $activity

                #write-output $activity -replace("&lt;","Less than")
                
                # $SessionsInProgress = $a.Activity.SessionsInProgress -replace("&lt;","Less than")
                # write-output $SessionsInProgress | format-list
                # $activity.RemainingTime.replace("&lt;","Less than")

                write-output $a.Activity.SessionsInProgress

                

                $B= $Details3[5] | convertfrom-json 
                $C= $Details3[8] | convertfrom-json 
                $D= $Details3[14] -replace(",\s\n.*\}","}") -replace("\},\n*\s*\]","}]") | convertfrom-json 
                $E= $Details3[17] | convertfrom-json 
                
                Write-Host ""
                if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port control.application-status.get }catch{ Write-Warning "ERROR: $_" }}
                if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port control.status.get }catch{ Write-Warning "ERROR: $_" }}
                if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port control.session.error.list -datasource VirtualDisasterRecovery}catch{ Write-Warning "ERROR: $_" }}

                #if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port storage.test }catch{ Write-Warning "ERROR: $_" }}
                #if ((get-process "BackupFP" -ea SilentlyContinue) -eq $Null) { "BackupFP Not Running" }else{ try { $ErrorActionPreference = 'Stop'; & "C:\Program Files\RecoveryConsole\ClientTool.exe" fp -host 127.0.0.1 -port $port control.session.list -datasource VirtualDisasterRecovery}catch{ Write-Warning "ERROR: $_" }}
       


                } ## End If
            
sleep -Seconds 5
            }  ## End Foreach

pause