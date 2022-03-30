clear-host

$Path = "C:\ProgramData\MXB\Backup Manager\logs\RecoveryConsole"

$LogHistory = 7
 
## Filtered Logs


Get-ChildItem $Path | 
select * -skiplast 1 | 
select * -last $LogHistory | 
get-content |
select-string -Pattern 'restore for device','Requesting restore for','tried for restore','skipped','user','abort','crashed','force' |
Select-Object -ExpandProperty Line | foreach-object { 
$_.Replace("VirtualDisasterRecoveryPlugin","VDR").Replace("Virtual disaster recovery","VDR").Replace("restored session Id is ","Session #").Replace("(session Id to restore: ","(Session #").Replace("session Id tried for restore was ","Session #").Replace(": requested restore of session #","Session #")        
}


Get-ChildItem $Path | 
select * -last 1 | 
get-content -wait |
select-string -Pattern 'restore for device','Requesting restore for','tried for restore','skipped','user','abort','crashed','force' |
Select-Object -ExpandProperty Line | foreach-object { 
$_.Replace("VirtualDisasterRecoveryPlugin","VDR").Replace("Virtual disaster recovery","VDR").Replace("restored session Id is ","Session #").Replace("(session Id to restore: ","(Session #").Replace("session Id tried for restore was ","Session #").Replace(": requested restore of session #","Session #")         
}



if ($unfiltered) {

## Unfiltered Logs

    Get-ChildItem $Path | 
    select * -skiplast 1 | 
    select * -last $LogHistory | 
    get-content |
    foreach-object { 
    $_.Replace("VirtualDisasterRecoveryPlugin","VDR").Replace("Virtual disaster recovery","VDR").Replace("restored session Id is ","Session #").Replace("(session Id to restore: ","(Session #").Replace("session Id tried for restore was ","Session #").Replace(": requested restore of session #","Session #")         

    }

    Get-ChildItem $Path | 
    select * -last 1 | 
    get-content -wait |
    foreach-object { 
    $_.Replace("VirtualDisasterRecoveryPlugin","VDR").Replace("Virtual disaster recovery","VDR").Replace("restored session Id is ","Session #").Replace("(session Id to restore: ","(Session #").Replace("session Id tried for restore was ","Session #").Replace(": requested restore of session #","Session #")         

    }
}

#>
