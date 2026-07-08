<# ----- About: ----
    # Sync Cove Customers to ConnectWise Companies
    # Compares Cove partner hierarchy to CWM companies and creates missing ones
    # Revision v10 - 2026-01-08
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
    # For use with N-able | Cove Data Protection
    # Requires ConnectWise Manage API access
    # Credentials are stored using Windows DPAPI encryption and can only be 
    # decrypted by the same user account on the same machine where created.
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Authenticate to Cove Data Protection API
    # Authenticate to ConnectWise Manage API
    # Enumerate Cove partners with backup issues (or use existing CSV)
    # Resolve partners to End-Customer level (skips Sites)
    # Compare against existing ConnectWise companies using multiple matching strategies:
    #   - Exact name match
    #   - Identifier match (company ID embedded in Cove partner reference)
    #   - Normalized name match (case-insensitive, special characters removed)
    # Display comparison results in GridView for review
    # Create missing companies in ConnectWise Manage with:
    #   - Auto-generated intelligent identifier (company abbreviation)
    #   - Configurable status (default: Active)
    #   - Configurable type (default: Customer)
    # Export comparison results to CSV for record-keeping
    #
    # Use the -PartnerName parameter to specify Cove partner name to analyze
    # Use the -CSVPath parameter to load existing monitoring script CSV (skips API enumeration)
    # Use the -CreateCompany parameter to create specific company (skip GridView selection)
    # Use the -CompanyStatus parameter to set company status (default="Active")
    # Use the -CompanyType parameter to set company type (default="Customer")
    # Use the -WhatIf parameter to preview company creation without making changes (default=$false)
    # Use the -NonInteractive parameter to skip user prompts when called from another script (default=$false)
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
    # https://github.com/christaylorcodes/ConnectWiseManageAPI
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)][string]$PartnerName = "", 
    
    [Parameter(Mandatory=$false)][string]$CSVPath = "",             ## Path to CSV from Cove monitoring script (skips API enumeration)
    
    [Parameter(Mandatory=$false)][string]$CreateCompany = "",       ## Specific company name to create (skip GridView)
    
    [Parameter(Mandatory=$false)][string]$CompanyStatus = "Active", ## ConnectWise Company Status to assign
    
    [Parameter(Mandatory=$false)][string]$CompanyType = "Customer", ## ConnectWise Company Type to assign
    
    [Parameter(Mandatory=$false)][bool]$WhatIf = $false,            ## Simulate company creation without making changes
    
    [Parameter(Mandatory=$false)][bool]$NonInteractive = $false     ## Skip user prompts when called from another script
)

#Requires -Version 7.0

# PowerShell 7 version check with helpful error message
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "Download PowerShell 7: https://aka.ms/powershell" -ForegroundColor Cyan
    exit 1
}

Clear-Host
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "`n=== Sync Cove Customers to ConnectWise Companies ===" -ForegroundColor Cyan

#region ----- Authenticate to Cove API -----

Write-Host "`nAuthenticating to Cove Backup.Management API..." -ForegroundColor Cyan

$APIcredfile = "C:\ProgramData\MXB\${env:computername}_${env:username}_API_Credentials.Secure.xml"
if (-not (Test-Path $APIcredfile)) {
    Write-Host "ERROR: Cove API credentials not found at: $APIcredfile" -ForegroundColor Red
    Write-Host "TIP: Run Cove2CWM-SyncTickets.v10.ps1 script first to create credentials" -ForegroundColor Yellow
    exit 1
}

# Load XML credential file
$APIcredentials = Import-Clixml -Path $APIcredfile
$cred0 = $APIcredentials.PartnerName
$cred1 = $APIcredentials.Username
$cred2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR(($APIcredentials.Password | ConvertTo-SecureString)))

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$url = "https://api.backup.management/jsonapi"

$data = @{
    jsonrpc = '2.0'
    id = '2'
    method = 'Login'
    params = @{
        username = $cred1
        password = $cred2
    }
}

$response = Invoke-RestMethod -Method POST -ContentType 'application/json' -Body (ConvertTo-Json $data) -Uri $url
$Script:visa = $response.visa

Write-Host "✓ Authenticated to Cove API: $cred0" -ForegroundColor Green

#endregion

#region ----- Cove API Helper Functions -----

Function Send-EnumerateAncestorPartners ($PartnerId) {
    $url = "https://api.backup.management/jsonapi"
    $data = @{
        jsonrpc = '2.0'
        id = '2'
        visa = $Script:visa
        method = 'EnumerateAncestorPartners'
        params = @{ partnerId = [int]$PartnerId }
    }

    $params = @{
        Uri         = $url
        Method      = 'POST'
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $data -depth 6)))
        ContentType = 'application/json; charset=utf-8'
    }
    
    return Invoke-RestMethod @params
}

Function Get-PartnerHierarchyInfo ($PartnerID, $PartnerName) {
    # Walk up partner tree to find End Customer level
    
    # Check cache first
    if ($Script:PartnerHierarchyCache.ContainsKey($PartnerID)) {
        return $Script:PartnerHierarchyCache[$PartnerID]
    }
    
    $result = @{
        EndCustomer = $PartnerName
        Site = $null
        PartnerId = $PartnerID
        Level = 'Unknown'
    }
    
    try {
        # First get this partner's level
        $partnerInfo = Send-GetPartnerInfo -PartnerName $PartnerName
        if ($partnerInfo.result.result) {
            $result.Level = $partnerInfo.result.result.Level
        }
        
        $ancestors = Send-EnumerateAncestorPartners -PartnerId $PartnerID
        
        if ($ancestors.result.result) {
            foreach ($ancestor in $ancestors.result.result) {
                if ($ancestor.Level -eq 'EndCustomer') {
                    $result.EndCustomer = $ancestor.Name
                    $result.PartnerId = $ancestor.Id
                    if ($PartnerName -ne $ancestor.Name) {
                        $result.Site = $PartnerName
                    }
                    break
                }
            }
        }
    }
    catch {
        Write-Verbose "Could not retrieve ancestor partners for $PartnerID"
    }
    
    $Script:PartnerHierarchyCache[$PartnerID] = $result
    return $result
}

Function Send-GetPartnerInfo ($PartnerName) {
    $url = "https://api.backup.management/jsonapi"
    $data = @{
        jsonrpc = '2.0'
        id = '2'
        visa = $Script:visa
        method = 'GetPartnerInfo'
        params = @{ name = $PartnerName }
    }

    $params = @{
        Uri         = $url
        Method      = 'POST'
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $data -depth 6)))
        ContentType = 'application/json; charset=utf-8'
    }
    
    return Invoke-RestMethod @params
}

Function Get-IntelligentIdentifier {
    param(
        [string]$CompanyName,
        [array]$ExistingCompanies
    )
    
    # Clean the name (remove special chars, keep alphanumeric and spaces)
    $cleanName = $CompanyName -replace '[^a-zA-Z0-9\s]', ''
    
    # Strategy 1: Use clean full name (truncate to 25 chars)
    $fullClean = $cleanName -replace '\s+', ''
    if ($fullClean.Length -le 25) {
        $exists = $ExistingCompanies | Where-Object { $_.identifier -eq $fullClean }
        if (-not $exists) {
            return $fullClean
        }
    } else {
        # Truncate to 25 chars
        $truncated = $fullClean.Substring(0, 25)
        $exists = $ExistingCompanies | Where-Object { $_.identifier -eq $truncated }
        if (-not $exists) {
            return $truncated
        }
    }
    
    # Strategy 2: First word only (if unique and descriptive)
    $firstWord = ($cleanName -split '\s+')[0]
    if ($firstWord.Length -ge 5 -and $firstWord.Length -le 25) {
        $exists = $ExistingCompanies | Where-Object { $_.identifier -eq $firstWord }
        if (-not $exists) {
            return $firstWord
        }
    }
    
    # Strategy 3: Acronym (fallback for very long names)
    $words = $cleanName -split '\s+' | Where-Object { $_ -match '^[A-Za-z]' }
    $acronym = ($words | ForEach-Object { $_[0] }) -join ''
    $acronym = $acronym.ToUpper()
    if ($acronym.Length -ge 3 -and $acronym.Length -le 25) {
        $exists = $ExistingCompanies | Where-Object { $_.identifier -eq $acronym }
        if (-not $exists) {
            return $acronym
        }
    }
    
    # If all fail, append numbers to truncated full name
    $baseIdentifier = if ($fullClean.Length -gt 22) { $fullClean.Substring(0, 22) } else { $fullClean }
    $counter = 1
    do {
        $testIdentifier = "$baseIdentifier$counter"
        $exists = $ExistingCompanies | Where-Object { $_.identifier -eq $testIdentifier }
        $counter++
    } while ($exists -and $counter -lt 100)
    
    return $testIdentifier
}

#endregion

# Initialize hierarchy cache
$Script:PartnerHierarchyCache = @{}

#region ----- Fast Path: Direct Company Creation -----

# If CreateCompany is specified in NonInteractive mode, skip all enumeration and create directly
if ($CreateCompany -and $NonInteractive) {
    Write-Host "`nDirect company creation mode (non-interactive)" -ForegroundColor Cyan
    
    # Authenticate to ConnectWise only
    Write-Host "Authenticating to ConnectWise Manage..." -ForegroundColor Cyan
    
    if (-not (Get-Module -ListAvailable -Name "ConnectWiseManageAPI")) {
        Install-Module -Name ConnectWiseManageAPI -Force -AllowClobber
    }
    Import-Module ConnectWiseManageAPI -ErrorAction SilentlyContinue
    
    $CWMAPICredsFile = "C:\ProgramData\MXB\${env:computername}_${env:username}_CWM_Ticketing_Credentials.Secure.xml"
    
    if (-not (Test-Path $CWMAPICredsFile)) {
        Write-Host "ERROR: ConnectWise credentials not found" -ForegroundColor Red
        exit 1
    }
    
    $CWMAPICreds = Import-Clixml -Path $CWMAPICredsFile
    
    $securePrivateKey = $CWMAPICreds.privateKey | ConvertTo-SecureString -Force
    $CWMAPICreds.privateKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePrivateKey))
    
    $securePubKey = $CWMAPICreds.pubKey | ConvertTo-SecureString -Force
    $CWMAPICreds.pubKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePubKey))
    
    $secureClientId = $CWMAPICreds.clientId | ConvertTo-SecureString -Force
    $CWMAPICreds.clientId = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureClientId))
    
    Connect-CWM @CWMAPICreds
    Write-Host "✓ Connected to ConnectWise: $($CWMAPICreds.Server)" -ForegroundColor Green
    
    # Get existing companies for identifier uniqueness check
    $cwmCompanies = Get-CWMCompany -all
    
    # Generate intelligent identifier
    $suggestedIdentifier = Get-IntelligentIdentifier -CompanyName $CreateCompany -ExistingCompanies $cwmCompanies
    
    Write-Host "Creating company: $CreateCompany" -ForegroundColor Cyan
    Write-Host "  Identifier: $suggestedIdentifier" -ForegroundColor Gray
    
    try {
        # Get company statuses and types
        $companyStatuses = Get-CWMCompanyStatus -all
        $companyTypes = Get-CWMCompanyType -all
        
        $statusMatch = $companyStatuses | Where-Object { $_.name -eq $CompanyStatus }
        $typeMatch = $companyTypes | Where-Object { $_.name -eq $CompanyType }
        
        $statusId = if ($statusMatch) { $statusMatch.id } else { 1 }
        $typeId = if ($typeMatch) { $typeMatch.id } else { 1 }
        
        $companySite = @{ name = "Main" }
        
        # Truncate company name to 50 characters (ConnectWise limit)
        # But store full name for matching later
        $fullCompanyName = $CreateCompany
        $companyName = if ($CreateCompany.Length -gt 50) {
            $CreateCompany.Substring(0, 50)
        } else {
            $CreateCompany
        }
        
        $result = New-CWMCompany -identifier $suggestedIdentifier -name $companyName -status @{id=$statusId} -type @{id=$typeId} -site $companySite
        
        # Return both full name and CWM company object for cache matching
        Write-Host "✓ Created company [ID: $($result.id)] $($result.name)" -ForegroundColor Green
        if ($fullCompanyName -ne $companyName) {
            Write-Host "  Note: Full name '$fullCompanyName' truncated to '$companyName'" -ForegroundColor Yellow
        }
        exit 0
    }
    catch {
        Write-Host "✗ Failed to create company: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

#endregion

#region ----- Get Cove Customers -----

if ($CSVPath) {
    # Load from CSV (monitoring script output)
    Write-Host "`nLoading customers from CSV: $CSVPath" -ForegroundColor Cyan
    $devices = Import-Csv $CSVPath
    
    $partnersWithIssues = $devices | Where-Object { $_.IssueSeverity -ne 'Success' } | 
        Select-Object PartnerName, PartnerId -Unique
    
    Write-Host "Found $($partnersWithIssues.Count) unique partners with issues" -ForegroundColor Cyan
} else {
    # Get from API
    if (-not $PartnerName) {
        $PartnerName = Read-Host "`nEnter EXACT partner name to analyze"
    }
    
    Write-Host "`nGetting partner info for: $PartnerName" -ForegroundColor Cyan
    
    $partnerResponse = Send-GetPartnerInfo -PartnerName $PartnerName
    
    if ($partnerResponse.error) {
        Write-Host "Error: $($partnerResponse.error.message)" -ForegroundColor Red
        exit 1
    }
    
    $partnerId = $partnerResponse.result.result.Id
    $partnerLevel = $partnerResponse.result.result.Level
    
    Write-Host "✓ Partner: $PartnerName (ID: $partnerId, Level: $partnerLevel)" -ForegroundColor Green
    
    # Get devices for this partner
    Write-Host "Retrieving devices..." -ForegroundColor Cyan
    
    $data = @{
        jsonrpc = '2.0'
        id = '2'
        visa = $Script:visa
        method = 'EnumerateAccountStatistics'
        params = @{
            query = @{
                PartnerId = [int]$partnerId
                Columns = @("AU","AR","AN","PF","PN")
                OrderBy = "AR ASC"
                StartRecordNumber = 0
                RecordsCount = 5000
            }
        }
    }
    
    $jsondata = ConvertTo-Json $data -depth 6
    $params = @{
        Uri         = $url
        Method      = 'POST'
        Headers     = @{ 'Authorization' = "Bearer $Script:visa" }
        Body        = ([System.Text.Encoding]::UTF8.GetBytes($jsondata))
        ContentType = 'application/json; charset=utf-8'
    }
    
    $deviceResponse = Invoke-RestMethod @params
    
    # Group by partner/customer
    $partnersWithIssues = $deviceResponse.result.result | Group-Object { $_.Settings.AR -join '' } | ForEach-Object {
        $firstDevice = $_.Group[0]
        [PSCustomObject]@{
            PartnerName = $firstDevice.Settings.AR -join ''
            PartnerId = ''  # Will be resolved via API lookup
            Reference = $firstDevice.Settings.PF -join ''
            DeviceCount = $_.Count
        }
    }
    
    Write-Host "✓ Found $($partnersWithIssues.Count) unique customers with $($deviceResponse.result.result.Count) total devices" -ForegroundColor Green
}

#endregion

#region ----- Resolve to End-Customer Level -----

Write-Host "`nResolving partner hierarchy to End-Customer level..." -ForegroundColor Cyan

$endCustomers = @{}

foreach ($partner in $partnersWithIssues) {
    $partnerId = $partner.PartnerId
    $partnerName = $partner.PartnerName
    
    if (-not $partnerId -or $partnerId -eq "") {
        try {
            $partnerInfo = Send-GetPartnerInfo -PartnerName $partnerName
            $partnerId = $partnerInfo.result.result.Id
        }
        catch {
            Write-Warning "Could not get partner ID for: $partnerName"
            continue
        }
    }
    
    $hierarchy = Get-PartnerHierarchyInfo -PartnerID $partnerId -PartnerName $partnerName
    $endCustomerName = $hierarchy.EndCustomer
    
    if (-not $endCustomers.ContainsKey($endCustomerName)) {
        $endCustomers[$endCustomerName] = @{
            Name = $endCustomerName
            PartnerId = $hierarchy.PartnerId
            Level = $hierarchy.Level
            Reference = if ($partner.Reference) { $partner.Reference } else { '' }
            DeviceCount = if ($partner.DeviceCount) { $partner.DeviceCount } else { 1 }
            Sites = @()
        }
        
        if ($hierarchy.Site) {
            $endCustomers[$endCustomerName].Sites += $hierarchy.Site
        }
    } else {
        $endCustomers[$endCustomerName].DeviceCount += if ($partner.DeviceCount) { $partner.DeviceCount } else { 1 }
        
        if ($hierarchy.Site -and $hierarchy.Site -notin $endCustomers[$endCustomerName].Sites) {
            $endCustomers[$endCustomerName].Sites += $hierarchy.Site
        }
    }
}

Write-Host "✓ Resolved to $($endCustomers.Count) unique End-Customer partners" -ForegroundColor Green

#endregion

#region ----- Authenticate to ConnectWise -----

Write-Host "`nAuthenticating to ConnectWise Manage..." -ForegroundColor Cyan

# Install/Import CWM module
if (-not (Get-Module -ListAvailable -Name "ConnectWiseManageAPI")) {
    Write-Host "Installing ConnectWise Manage PowerShell Module..." -ForegroundColor Yellow
    Install-Module -Name ConnectWiseManageAPI -Force -AllowClobber
}
Import-Module ConnectWiseManageAPI -ErrorAction SilentlyContinue

$CWMAPICredsFile = "C:\ProgramData\MXB\${env:computername}_${env:username}_CWM_Ticketing_Credentials.Secure.xml"

if (-not (Test-Path $CWMAPICredsFile)) {
    Write-Host "ERROR: ConnectWise credentials not found" -ForegroundColor Red
    exit 1
}

$CWMAPICreds = Import-Clixml -Path $CWMAPICredsFile

$securePrivateKey = $CWMAPICreds.privateKey | ConvertTo-SecureString -Force
$CWMAPICreds.privateKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePrivateKey))

$securePubKey = $CWMAPICreds.pubKey | ConvertTo-SecureString -Force
$CWMAPICreds.pubKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePubKey))

$secureClientId = $CWMAPICreds.clientId | ConvertTo-SecureString -Force
$CWMAPICreds.clientId = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureClientId))

Connect-CWM @CWMAPICreds
Write-Host "✓ Connected to ConnectWise: $($CWMAPICreds.Server)" -ForegroundColor Green

# Get existing companies
Write-Host "Retrieving existing ConnectWise companies..." -ForegroundColor Cyan
$cwmCompanies = Get-CWMCompany -all
Write-Host "✓ Found $($cwmCompanies.Count) ConnectWise companies" -ForegroundColor Green

#endregion

#region ----- Compare and Match -----

Write-Host "`nComparing Cove customers to ConnectWise companies..." -ForegroundColor Cyan

$comparisonResults = @()

foreach ($customer in $endCustomers.Values) {
    # Filter: Only include EndCustomer level (exclude Reseller, Root, Sub-root, ServiceOrg)
    if ($customer.Level -ne 'EndCustomer') {
        Write-Verbose "Skipping non-EndCustomer: $($customer.Name) (Level: $($customer.Level))"
        continue
    }
    
    $matchMethod = "No Match"
    $cwmCompanyName = ""
    $cwmCompanyId = ""
    $suggestedIdentifier = Get-IntelligentIdentifier -CompanyName $customer.Name -ExistingCompanies $cwmCompanies
    
    # Strategy 1: Extract CWM Company ID from partner name (e.g., "Company | ID ~ 12345")
    if ($customer.Name -match '\|\s*ID\s*~\s*(\d+)') {
        $extractedId = $matches[1]
        $match = $cwmCompanies | Where-Object { $_.id -eq $extractedId }
        if ($match) {
            $matchMethod = "CWM ID from Name"
            $cwmCompanyName = ($match | Select-Object -First 1).name
            $cwmCompanyId = [int](($match | Select-Object -First 1).id)
        }
    }
    
    # Strategy 2: Exact name match
    if ($matchMethod -eq "No Match") {
        $match = $cwmCompanies | Where-Object { $_.name -eq $customer.Name }
        if ($match) {
            $matchMethod = "Exact Name Match"
            $cwmCompanyName = ($match | Select-Object -First 1).name
            $cwmCompanyId = [int](($match | Select-Object -First 1).id)
        }
    }
    
    # Strategy 3: Clean name match (strip email/extra info)
    if ($matchMethod -eq "No Match") {
        $cleanCustomerName = ($customer.Name -split '\(')[0].Trim()
        $match = $cwmCompanies | Where-Object { 
            $cleanCWMName = ($_.name -split '\(')[0].Trim()
            $cleanCWMName -eq $cleanCustomerName
        }
        if ($match) {
            $matchMethod = "Clean Name Match"
            $cwmCompanyName = ($match | Select-Object -First 1).name
            $cwmCompanyId = [int](($match | Select-Object -First 1).id)
        }
    }
    
    # Strategy 4: Partial name match
    if ($matchMethod -eq "No Match") {
        $cleanCustomerName = ($customer.Name -split '\(')[0].Trim()
        $match = $cwmCompanies | Where-Object { $_.name -like "*$cleanCustomerName*" }
        if ($match) {
            $matchMethod = "Partial Name Match"
            $cwmCompanyName = $match[0].name
            # Ensure ID is a single integer, not an array
            $cwmCompanyId = [int]($match[0].id | Select-Object -First 1)
        }
    }
    
    $comparisonResults += [PSCustomObject]@{
        CoveCustomer = $customer.Name
        CoveReference = $customer.Reference
        CoveLevel = $customer.Level
        DeviceCount = $customer.DeviceCount
        Sites = if ($customer.Sites.Count -gt 0) { $customer.Sites -join '; ' } else { '' }
        MatchStatus = if ($matchMethod -eq "No Match") { "Missing" } else { "Matched" }
        MatchMethod = $matchMethod
        CWMCompany = $cwmCompanyName
        CWMCompanyID = $cwmCompanyId
        SuggestedIdentifier = $suggestedIdentifier
        CovePartnerId = $customer.PartnerId
    }
}

# Summary
$matched = ($comparisonResults | Where-Object { $_.MatchStatus -eq "Matched" }).Count
$missing = ($comparisonResults | Where-Object { $_.MatchStatus -eq "Missing" }).Count

Write-Host "`n=== Comparison Summary ===" -ForegroundColor Cyan
Write-Host "Total Cove End-Customers: $($comparisonResults.Count)" -ForegroundColor White
Write-Host "Matched in CWM: $matched ($([math]::Round($matched/$comparisonResults.Count*100,1))%)" -ForegroundColor Green
Write-Host "Missing in CWM: $missing ($([math]::Round($missing/$comparisonResults.Count*100,1))%)" -ForegroundColor $(if($missing -gt 0){'Yellow'}else{'Green'})

#endregion

#region ----- Handle Company Creation -----

if ($CreateCompany) {
    # Create specific company (bypass GridView)
    $companyToCreate = $comparisonResults | Where-Object { $_.CoveCustomer -eq $CreateCompany -and $_.MatchStatus -eq "Missing" }
    
    if (-not $companyToCreate) {
        Write-Host "`nERROR: Company '$CreateCompany' not found in missing companies list" -ForegroundColor Red
        Write-Host "Available missing companies:" -ForegroundColor Yellow
        $comparisonResults | Where-Object { $_.MatchStatus -eq "Missing" } | Select-Object CoveCustomer | Format-Table
        exit 1
    }
    
    $selectedCompanies = @($companyToCreate)
    Write-Host "`nCreating specific company: $CreateCompany" -ForegroundColor Cyan
} else {
    # Show GridView for selection
    Write-Host "`nOpening comparison results in GridView..." -ForegroundColor Cyan
    Write-Host "Select companies to create (missing companies only), or Cancel to exit" -ForegroundColor Yellow
    
    $selectedCompanies = $comparisonResults | 
        Sort-Object MatchStatus, CoveCustomer |
        Out-GridView -Title "Cove to ConnectWise Comparison - Select MISSING Companies to Create" -OutputMode Multiple
    
    if (-not $selectedCompanies -or $selectedCompanies.Count -eq 0) {
        Write-Host "`nNo companies selected. Exiting." -ForegroundColor Yellow
        
        # Export results
        $companiesFolder = "$PSScriptRoot\Companies"
        if (-not (Test-Path $companiesFolder)) {
            New-Item -Path $companiesFolder -ItemType Directory -Force | Out-Null
        }
        $exportPath = "$companiesFolder\CoveToConnectWise_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $comparisonResults | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "Results exported to: $exportPath" -ForegroundColor Green
        exit 0
    }
    
    # Filter to only missing companies
    $selectedCompanies = $selectedCompanies | Where-Object { $_.MatchStatus -eq "Missing" }
    
    if ($selectedCompanies.Count -eq 0) {
        Write-Host "`nNo missing companies selected. All selected companies already exist in ConnectWise." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "`nSelected $($selectedCompanies.Count) companies to create:" -ForegroundColor Green
$selectedCompanies | ForEach-Object {
    $siteInfo = if ($_.Sites) { " (Sites: $($_.Sites))" } else { "" }
    Write-Host "  - $($_.CoveCustomer) - $($_.DeviceCount) devices$siteInfo" -ForegroundColor White
    Write-Host "    Suggested Identifier: $($_.SuggestedIdentifier)" -ForegroundColor Gray
}

if ($WhatIf) {
    Write-Host "`n[WHATIF] Would create $($selectedCompanies.Count) companies" -ForegroundColor Magenta
    exit 0
}

# Confirm creation (skip in non-interactive mode)
if (-not $NonInteractive) {
    Write-Host "`nPress ENTER to create these companies, or Ctrl+C to cancel..." -ForegroundColor Yellow
    Read-Host
} else {
    Write-Host "`n[NonInteractive Mode] Creating companies automatically..." -ForegroundColor Cyan
}

#endregion

#region ----- Create Companies -----

# Get company statuses and types
$companyStatuses = Get-CWMCompanyStatus -all
$companyTypes = Get-CWMCompanyType -all

$statusMatch = $companyStatuses | Where-Object { $_.name -eq $CompanyStatus }
$typeMatch = $companyTypes | Where-Object { $_.name -eq $CompanyType }

if ($statusMatch) {
    $statusId = $statusMatch.id
} else {
    $statusId = 1  # Default Active
}

if ($typeMatch) {
    $typeId = $typeMatch.id
} else {
    $typeId = 1  # Default Customer
}

Write-Host "`nCreating companies..." -ForegroundColor Cyan
Write-Host "Status: $CompanyStatus (ID: $statusId)" -ForegroundColor Gray
Write-Host "Type: $CompanyType (ID: $typeId)" -ForegroundColor Gray

$created = 0
$failed = 0
$results = @()

foreach ($company in $selectedCompanies) {
    try {
        Write-Host "`n  Creating: $($company.CoveCustomer)..." -ForegroundColor White -NoNewline
        Write-Host " [ID: $($company.SuggestedIdentifier)]" -ForegroundColor Gray -NoNewline
        
        $companySite = @{ name = "Main" }
        
        # Truncate company name to 50 characters (ConnectWise limit)
        # But store full name for matching later
        $fullCompanyName = $company.CoveCustomer
        $companyName = if ($company.CoveCustomer.Length -gt 50) {
            $company.CoveCustomer.Substring(0, 50)
        } else {
            $company.CoveCustomer
        }
        
        $result = New-CWMCompany -identifier $company.SuggestedIdentifier -name $companyName -status @{id=$statusId} -type @{id=$typeId} -site $companySite
        
        if ($fullCompanyName -ne $companyName) {
            Write-Host "  Note: Full name '$fullCompanyName' truncated to '$companyName'" -ForegroundColor Yellow
        }
        
        Write-Host " ✓ Created (CWM ID: $($result.id))" -ForegroundColor Green
        
        $results += [PSCustomObject]@{
            CoveCustomer = $company.CoveCustomer
            Identifier = $company.SuggestedIdentifier
            CWMCompanyID = $result.id
            Status = "Created"
            DeviceCount = $company.DeviceCount
            Sites = $company.Sites
        }
        
        $created++
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Host " ✗ FAILED: $errorMsg" -ForegroundColor Red
        
        $results += [PSCustomObject]@{
            CoveCustomer = $company.CoveCustomer
            Identifier = $company.SuggestedIdentifier
            CWMCompanyID = ""
            Status = "Failed: $errorMsg"
            DeviceCount = $company.DeviceCount
            Sites = $company.Sites
        }
        
        $failed++
    }
}

#endregion

#region ----- Summary and Export -----

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export creation results
if ($results.Count -gt 0) {
    $creationExportPath = "$PSScriptRoot\tickets\CompanyCreation_$timestamp.csv"
    $results | Export-Csv -Path $creationExportPath -NoTypeInformation
    Write-Host "`nCreation results exported to: $creationExportPath" -ForegroundColor Cyan
}

# Export comparison results
$companiesFolder = "$PSScriptRoot\Companies"
if (-not (Test-Path $companiesFolder)) {
    New-Item -Path $companiesFolder -ItemType Directory -Force | Out-Null
}
$comparisonExportPath = "$companiesFolder\CoveToConnectWise_Comparison_$timestamp.csv"
$comparisonResults | Export-Csv -Path $comparisonExportPath -NoTypeInformation
Write-Host "Comparison results exported to: $comparisonExportPath" -ForegroundColor Cyan

# Summary
Write-Host "`n=== Creation Summary ===" -ForegroundColor Cyan
Write-Host "Successfully Created: $created" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if($failed -gt 0){'Red'}else{'Green'})

if ($created -gt 0) {
    Write-Host "`n✓ Companies created successfully!" -ForegroundColor Green
    Write-Host "  You can now run the monitoring script to create tickets" -ForegroundColor Gray
}

#endregion
