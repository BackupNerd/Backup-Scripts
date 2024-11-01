<# ----- About: ----
    # N-able | Analyze billing efile
    # Revision v24.11.01 - 2024-11-01
    # Author: Eric Harless, Head Backup Nerd - N-able 
    # Twitter @Backup_Nerd  Email:eric.harless@N-able.com    
    # Reddit https://www.reddit.com/r/Nable/
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
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
    # For use with the N-able monthly billing efile
    # Sample scripts are subject to change without notification
    # Some script elements may be developed, tested or documented using AI
  # -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # This script analyzes N-able billing efile data.
    # The script then processes the selected files to group usage data by unique customer/site and product fields.
    # It calculates subtotals for each group and exports both detailed and summary reports to CSV files.
    # The script requires the ImportExcel module to handle XLSX files.
    # Requires ImportExcel module
    # 
    # To download your N-able billing efile, follow these steps:
    # 
    # Log in to N-able Me: Go to the N-able Me login page.
    # Navigate to Billing: On the left menu bar, click on My Account and then select Billing.
    # Select the Invoice: In the billing portal, choose either the Payment History or Outstanding Invoices tab.
    # Download the efile: Find the invoice you need, click the ellipses (three dots) next to it, and select Data Usage to download the efile.
    # If you encounter any issues, you can create a case with N-able’s Customer Care team for further assistance.
    #
    # https://me.n-able.com/s/article/How-to-see-billing-and-invoice-information-for-MSP-products

    # -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)][string]$ExportPath = "$PSScriptRoot",                     ## Export Path
        [Parameter(Mandatory=$False)][switch]$IncludePerGB                                      ## Set to $false to exclude Per GB products from reports
    )


#region ----- Environment, Variables, Names and Paths ----
Clear-Host
#Requires -Version 5.1
$scriptStartTime = Get-Date  # Set the start time at the beginning of the script
$ConsoleTitle = "N-able | Analyze Billing efile"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

Write-output "  $ConsoleTitle`n`n$ScriptPath"
$Syntax = Get-Command $PSCommandPath -Syntax 
Write-Output "  Script Parameter Syntax:`n`n  $Syntax"

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Push-Location $dir
$CurrentDate = Get-Date -format "yyyy-MM-dd_HH-mm-ss"
$ShortDate = Get-Date -format "yyyy-MM-dd"

#end region ----- Environment, Variables, Names and Paths ----

if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "Module ImportExcel Already Installed"
} else {
    try {
        ## https://powershell.one/tricks/parsing/excel
        ## https://github.com/dfinke/ImportExcel https://github.com/dfinke/ImportExcel
        Install-Module -Name ImportExcel -Confirm:$False -Force
    }
    catch [Exception] {
        $_.message
        exit
    }
}

function Show-ElapsedTime {
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $scriptStartTime
    Write-Output "Elapsed time: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s"
}  ## Function to display the elapsed time


Function Open-FileName {
    param (
        [string]$initialDirectory,
        [switch]$multiselect
    )
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.Filter = "N-able Billing efile(Efile_*.csv;efile_*.xlsx)|efile_*.csv;efile_*.xlsx"
    $OpenFileDialog.Title = "Select one or more N-able Billing Efiles to Analyze (*.csv, *.xlsx)"
    $OpenFileDialog.Multiselect = $multiselect
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileNames
    } else {
        return @()
    }
}

Function Process-EFile  ($usageeData) {
    # Group by unique fields and calculate subtotal
    $subtotalDataa = $usageData | Group-Object -Property "Usage Customer", "Customer Site", Product | ForEach-Object {
        [PSCustomObject]@{
            UsageCustomer = $_.Group[-1]."Usage Customer"
            CustomerSite = if ($_.Group[-1]."Customer Site" -eq $_.Group[-1]."Usage Customer") { "" } else { $_.Group[-1]."Customer Site" }
            Product = $_.Group[-1].Product
            From = if (($null -ne $_.Group[-1]."Period From") -and ($_.Group[-1]."Period From" -ne "")) {get-date -Date $_.Group[-1]."Period From" -Format "dd-MMM" -ErrorAction SilentlyContinue } else {}
            To = if (($null -ne $_.Group[-1]."Period To") -and ($_.Group[-1]."Period To" -ne "")) {get-date -Date $_.Group[-1]."Period To" -Format "dd-MMM-yyyy" -ErrorAction SilentlyContinue } else {}
            RatingMethod = $_.Group[-1]."Rating Method"
            UOM = $_.Group[-1].UOM
            Quantity = ($_.Group | Measure-Object -Property Quantity -Sum).Sum
            
            ListPrice = $_.Group[-1]."List Price"
            #Rate = $_.Group[-1].Rate
            Rate = ($_.Group | Measure-Object -Property Rate -Maximum).Maximum
            #Cost = $_.Group[-1].Cost
            Cost = ($_.Group | Measure-Object -Property Cost -Sum).Sum
        }
    } -ErrorAction SilentlyContinue
    
    # Display the subtotal data
    if ($IncludePerGB) {
        $subtotalDataa | Sort-object Usagecustomer,Product | Format-Table -AutoSize
    } else {
        $subtotalDataa | where-object {$_.product -notlike "*Per GB*"} | Sort-object Usagecustomer,Product | Format-Table -AutoSize
    }

    # Export to CSV
     
    $outfilename = "$exportpath\$(@($usagedata)[0]."account name") - EFile_$(@($usagedata)[0]."Invoice Number") - Detail - $(get-date @($usagedata)[0]."Invoice Date" -Format "yyyy-MM-dd").csv"

    if ($IncludePerGB) {
        $subtotalDataa | Sort-object Usagecustomer,Product | Export-Csv -Path $outfilename -NoTypeInformation
    } else {
        $subtotalDataa | where-object {$_.product -notlike "*Per GB*"} | Sort-object Usagecustomer,Product | Export-Csv -Path $outfilename -NoTypeInformation
    }
    
    Write-output "File saved to $outfilename"

    # Group by unique fields and calculate subtotal
    $subtotalDatab = $usageData | Group-Object -Property Product | ForEach-Object {
        [PSCustomObject]@{
    
            Product = $_.Group[0].Product
            From = if (($null -ne $_.Group[-1]."Period From") -and ($_.Group[-1]."Period From" -ne "")) {get-date -Date $_.Group[-1]."Period From" -Format "dd-MMM" -ErrorAction SilentlyContinue } else {}
            To = if (($null -ne $_.Group[-1]."Period To") -and ($_.Group[-1]."Period To" -ne "")) {get-date -Date $_.Group[-1]."Period To" -Format "dd-MMM-yyyy" -ErrorAction SilentlyContinue } else {}
            RatingMethod = $_.Group[0]."Rating Method"
            UOM = $_.Group[0].UOM
            Quantity = ($_.Group | Measure-Object -Property Quantity -Sum).Sum
    
            ListPrice = $_.Group[0]."List Price"
            #Rate = $_.Group[-1].Rate
            Rate = ($_.Group | Measure-Object -Property Rate -Maximum).Maximum
            #Cost = $_.Group[-1].Cost
            Cost = ($_.Group | Measure-Object -Property Cost -Sum).Sum
        }
    } -ErrorAction SilentlyContinue
    
    if ($IncludePerGB) {
        $subtotalDatab | Sort-object Product | Format-Table -AutoSize
    } else {
        $subtotalDatab | where-object {$_.product -notlike "*Per GB*"} | Sort-object Product | Format-Table -AutoSize
    }

    # Export to CSV
    
    $outfilename = "$exportpath\$(@($usagedata)[0]."account name") - EFile_$(@($usagedata)[0]."Invoice Number") - Summary - $(get-date @($usagedata)[0]."Invoice Date" -Format "yyyy-MM-dd").csv"

    if ($IncludePerGB) {
        $subtotalDatab | Sort-object Usagecustomer,Product | Export-Csv -Path $outfilename -NoTypeInformation
    } else {
        $subtotalDatab | where-object {$_.product -notlike "*Per GB*"} | Sort-object Usagecustomer,Product | Export-Csv -Path $outfilename -NoTypeInformation
    }

    Write-output "File saved to $outfilename"

}

Open-FileName -initialDirectory $PSScriptRoot -multiselect | ForEach-Object {
    if ($_.EndsWith(".csv")) {
        $usageData = Import-Csv -Path $_
    } elseif ($_.EndsWith(".xlsx")) {
        $usageData = Import-Excel -Path $_
    }

    $requiredHeaders = @("Account Name", "Account Number", "Invoice Number", "Invoice Date", "Contract Number", "Tenant", "Product", "UOM", "Quantity", "Rate", "Cost", "List Price", "Usage Customer", "Customer Site", "Device Name", "Device Type", "Device ID", "Rating Method", "Period From", "Period To")
    $missingHeaders = $requiredHeaders | Where-Object { $_ -notin $usageData[0].PSObject.Properties.Name }

    if ($missingHeaders.Count -gt 0) {
        Write-Warning "File | $_ `nis missing the following required headers: $($missingHeaders -join ', ').`nPlease try again with an unmodified file."
        break
    } else {
        Process-EFile -usageData $usageData
    }
    Show-ElapsedTime 
}
