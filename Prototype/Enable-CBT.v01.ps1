Clear-Host
$ESXServer = Read-Host "Enter ESX Server address (FQDN or IP address):"
$ESXUser = Read-Host "Enter ESX user name:"
$ESXUserPassword = Read-Host "Enter ESX password:" -AsSecureString:$true

# Collect Username and Password as Credential
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $ESXUser,$ESXUserPassword

# Connect to ESX
Connect-VIServer -Server $ESXServer -Credential $Credentials

# List VMs with CBT disabled
$CBTdisabled = Get-VM | Where-Object{$_.ExtensionData.Config.ChangeTrackingEnabled -eq $false} 

Write-output "`nVMs with CBT currently DISABLED"
$CBTdisabled | Format-Table

$CBTdisabled = Get-VM | Where-Object{$_.ExtensionData.Config.ChangeTrackingEnabled -eq $false} | Out-GridView -title "Select one or more VMs to enable CBT (Change Block Tracking)" -OutputMode Multiple

foreach ($selectedVM in $CBTdisabled) {
 
    $vmConfig = get-vm $selectedVM.name | get-view
    $vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $vmConfigSpec.changeTrackingEnabled = $true
    $vmConfig.reconfigVM($vmConfigSpec)
    Write-output "Enabling CBT for $($selectedVM.name)"
}

# List VMs with CBT enabled
$CBTenabled = Get-VM | Where-Object{$_.ExtensionData.Config.ChangeTrackingEnabled -eq $true} 
Write-output "`nVMs with CBT NOW enabled"
$CBTenabled | Format-Table


<#
$CBTenabled = Get-VM | Where-Object{$_.ExtensionData.Config.ChangeTrackingEnabled -eq $true} | Out-GridView -title "Select one or more VMs to disable CBT (Change Block Tracking)" -OutputMode Multiple

foreach ($selectedVM in $CBTenabled) {
 
    $vmConfig = get-vm $selectedVM.name | get-view
    $vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $vmConfigSpec.changeTrackingEnabled = $false
    $vmConfig.reconfigVM($vmConfigSpec)
    Write-output "Disabling CBT for $($selectedVM.name)"
}
#>
