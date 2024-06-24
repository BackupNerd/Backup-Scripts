#$Script:CUID = "27046ddb-e61e-4d32-8f1c-b5text8db"
$script:CUID = Ninja-Property-Get covecustomeruid

Function ConvertTo-Base64($InputString) {
    $BytesString = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $global:OutputBase64 = [System.Convert]::ToBase64String($BytesString)
    Write-Output $global:OutputBase64
} ## Covert string to Base64 encoding

Function ConvertFrom-Base64($InputBase64) {
    $BytesBase64 = [System.Convert]::FromBase64String($InputBase64)
    $global:OutputString = [System.Text.Encoding]::UTF8.GetString($BytesBase64)
    Write-Output $global:OutputString
} ## Covert Base64 encoding to string

Function Get-TimeStamp {
    return "[{0:yyyy/MM/dd} {0:HH:mm:ss}]" -f (Get-Date)
} ## Get formated Timestamp for use in logging

Function Set-Alias {

    $BMConfig = "C:\Program Files\Backup Manager\config.ini"
    $clienttool = "C:\Program Files\Backup Manager\clienttool.exe"  

    if ((Test-Path $BMConfig -PathType leaf) -eq $true) {
        
        $script:alias = "$(get-date) Enter Text here or Call Function to get machine Name and usage site and time stamp etc that you want for the alias"

        $AliasParam = "-device-alias `"$(ConvertTo-Base64 $alias)`""

        Write-Output "`nUpdating Backup Manager Device Alias Column"    
        Write-Output "`n$Script:Alias"
        #Write-Output "Alias Base64   : $AliasParam" 
        #Write-Output "Partner UID    : $UID"
        Write-Output ""

        $global:process = start-process -FilePath "$clienttool" -ArgumentList "takeover -config-path `"$bmconfig`" -partner-uid $Script:CUID $AliasParam" -PassThru

        for($i = 0; $i -le 100; $i = ($i + 1) % 100) {
            Write-Progress -Activity "N-able Backup Manager $DeployType" -PercentComplete $i -Status "Installing"
            Start-Sleep -Milliseconds 100
            if ($process.HasExited) {
                Write-Progress -Activity "Installer" -Completed
                Break
                }
            }
        }
        else{
            Write-Output "`nSet Alias aborted, existing Backup Manager deployment not found"
            Break
            }
} ## Set local Device Alias Name to ...


    Set-Alias
