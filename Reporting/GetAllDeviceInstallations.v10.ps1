<# ----- About: ----
    # Get All N-able Backup Device Installations
    # Revision v10 - 2026-09-09
    #   - API CHANGE ADDRESSED: GetPartnerInfo(name=) is deprecated by N-able (silently resolves
    #     duplicate partner names to "oldest active match"; see
    #     https://developer.n-able.com/n-able-cove/docs/migrating-from-getpartnerinfo-to-getpartnerinfobyid)
    #   - Replaced with Get-CovePartnerByName (GetPartnerTree + partnerFilter.NamePattern, case-insensitive
    #     substring match, ALL matches returned) + Get-CovePartnerInfoById (GetPartnerInfoById for full
    #     Uid/Level details once an Id is chosen) - see Cove-MCP-Server CLAUDE.md "Known Migrations" section
    #   - Authenticated partner's own Level is checked first (GetPartnerInfoById on AuthPartnerId from
    #     Login); Root/Sub-root/Distributor levels are rejected up front and prompt for a real name,
    #     instead of wasting a GetPartnerTree search that can never match the root against itself
    #   - Out-GridView added for disambiguation when multiple partners match
    #   - Switched Invoke-WebRequest -> Invoke-RestMethod for all JSON API calls (avoids byte-array
    #     Content parsing gotcha; matches project-wide PowerShell convention)
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
    # For use with the Standalone edition of N-able Backup
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Check/ Get/ Store secure credentials 
    # Authenticate to https://backup.management console
    # Get / Prompt for exact partner name 
    # Enumerate device installations
    # Optionally display output via Powershell Out-Gridview
    # Export to XLS/CSV
    # Optionally launch XLS/CSV file
    #
    #   Results include all Historic installation instances of current devices (Backup, Restore Only, Bare-Metal Recovery, Recovery Console and Recovery Testing)
    #   Useful for auditing the last active date for a specific installation Id 
    #
    # Use the -GridView switch parameter to display output via Powershell Out-Gridview
    # Use the -DeviceCount ## (default=5000) parameter to define the maximum number of devices returned
    # Use the -Launch switch parameter to launch the XLS/CSV file after completion
    # Use the -ExportPath (?:\Folder) parameter to specify alternate XLS/CSV file path
    # Use the -Delimiter (default=',') parameter to set the delimiter for XLS/CSV output (i.e. use ';' for The Netherland)
    # Use the -ClearCredentials parameter to remove stored API credentials at start of script
    #
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/home.htm
    # https://documentation.n-able.com/backup/userguide/documentation/Content/service-management/json-api/API-column-codes.htm
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$False)] [int]$DeviceCount = 6000,         ## Change Maximum Number of devices results to return
        [Parameter(Mandatory=$False)] [switch]$GridView,                ## Display Output via Powershell Out-Gridview
        [Parameter(Mandatory=$False)] [switch]$Launch,                  ## Launch XLS or CSV file
        [Parameter(Mandatory=$False)] [string]$Delimiter = ',',         ## specify ',' or ';' Delimiter for XLS & CSV file
        [Parameter(Mandatory=$False)] $ExportPath = "$PSScriptRoot",    ## Export Path
        [Parameter(Mandatory=$False)] [switch]$ClearCredentials         ## Remove Stored API Credentials at start of script
    )
   

    #region ----- Environment, Variables, Names and Paths ----
    Clear-Host
    $ConsoleTitle = "Get All Device Installation"
    $host.UI.RawUI.WindowTitle = $ConsoleTitle
    $scriptpath = $MyInvocation.MyCommand.Path
    Write-output "  $ConsoleTitle`n`n$ScriptPath"
    $Syntax = Get-Command $PSCommandPath -Syntax
    Write-Output "  Script Parameter Syntax:`n`n  $Syntax"
    $dir = Split-Path $scriptpath
    Push-Location $dir
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Script:strLineSeparator = "  ---------"
    $CurrentDate = Get-Date -format "yyy-MM-dd_hh-mm-ss"
   
    Write-output "  Current Parameters:"
    Write-output "  -GridView      = $GridView"
    Write-output "  -Launch        = $Launch"
    Write-output "  -ExportPath    = $ExportPath"
    Write-output "  -Delimiter     = $Delimiter"

    









#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----
    Function Get-CoveXmlCredentials {
        ## DPAPI-encrypted credential store (Export-Clixml/Import-Clixml) - reuses an existing file
        ## in place (e.g. one already created by Cove-MCP-Server) instead of overwriting it.
        $Script:CredFile = "C:\ProgramData\MXB\mcpcred-iaso.xml"

        if ($ClearCredentials -and (Test-Path $Script:CredFile)) {
            Remove-Item -Path $Script:CredFile -Force
            $Script:ClearCredentials = $false
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Cleared"
        }

        if (Test-Path $Script:CredFile) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Using existing credential file: $Script:CredFile"
            $Script:BackupCred = Import-Clixml -Path $Script:CredFile
        }else{
            Write-Output $Script:strLineSeparator
            Write-Output "  No credential file found at $Script:CredFile"
            $Script:BackupCred = Get-Credential -Message 'Enter Login Email and Password for Backup.Management API'
            if (-not $Script:BackupCred) { throw "No credentials provided." }
            $credDir = Split-Path $Script:CredFile -Parent
            if (-not (Test-Path $credDir)) { New-Item -ItemType Directory -Path $credDir -Force | Out-Null }
            $Script:BackupCred | Export-Clixml -Path $Script:CredFile -Force
            Write-Output "  Credentials saved to $Script:CredFile"
        }

        $Script:cred1 = $Script:BackupCred.UserName
        $Script:cred2 = $Script:BackupCred.GetNetworkCredential().Password

        Write-Output $Script:strLineSeparator
        Write-output "  Stored Backup API User     = $Script:cred1"
        Write-output "  Stored Backup API Password = Encrypted"

    }  ## Get (or create) DPAPI-encrypted XML credentials
           
    Function Send-APICredentialsCookie {

    Get-CoveXmlCredentials  ## Read (or create) XML Credential File before Authentication

    $url = "https://api.backup.management/jsonapi"
    $data = @{}
    $data.jsonrpc = '2.0'
    $data.id = '2'
    $data.method = 'Login'
    $data.params = @{}
    $data.params.username = $Script:cred1
    $data.params.password = $Script:cred2

    $Script:Authenticate = Invoke-RestMethod -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $url
    
    #Debug Write-output "$($Script:cookies[0].name) = $($cookies[0].value)"
    
if ($authenticate.PSObject.Properties['visa'] -and $authenticate.visa) { 
    
        $Script:visa = $authenticate.visa
        $Script:AuthPartnerId = [int]$authenticate.result.result.PartnerId  ## Authenticated user's own partner Id, used as default GetPartnerTree search scope
        Write-Output $Script:strLineSeparator
        Write-output "  Authenticated PartnerId (default GetPartnerTree scope) = $Script:AuthPartnerId"
        }else{
            Write-Output    $Script:strLineSeparator 
            Write-output "  Authentication Failed: Please confirm your Backup.Management Credentials"
            Write-output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output    $Script:strLineSeparator 
            
            Remove-Item -Path $Script:CredFile -Force -ErrorAction SilentlyContinue  ## Stored creds are stale/invalid - re-prompt
            Send-APICredentialsCookie
        }

    }  ## Use Backup.Management credentials to Authenticate

    Function Get-VisaTime {
        if ($Script:visa) {
            $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
            If ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)){
                Send-APICredentialsCookie
            }
            
        }
    
    }

#endregion ----- Authentication ----

#region ----- Data Conversion ----
    Function Convert-UnixTimeToDateTime($inputUnixTime){
        if ($inputUnixTime -gt 0 ) {
        $epoch = Get-Date -Date "1970-01-01 00:00:00Z"
        $epoch = $epoch.ToUniversalTime()
        $epoch = $epoch.AddSeconds($inputUnixTime)
        return $epoch
        }else{ return ""}
    }  ## Convert epoch time to date time 

    Function Save-CSVasExcel {
        param (
            [string]$CSVFile = $(Throw 'No file provided.')
        )
        
        BEGIN {
            function Resolve-FullPath ([string]$Path) {    
                if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) {
                    # $Path = Join-Path (Get-Location) $Path
                    $Path = "$PWD\$Path"
                }
                [IO.Path]::GetFullPath($Path)
            }
    
            function Release-Ref ($ref) {
                ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            $CSVFile = Resolve-FullPath $CSVFile
            $xl = New-Object -ComObject Excel.Application
        }
    
        PROCESS {
            $wb = $xl.workbooks.open($CSVFile)
            $xlOut = $CSVFile -replace '\.csv$', '.xlsx'
            
            # can comment out this part if you don't care to have the columns autosized
            $ws = $wb.Worksheets.Item(1)
            $range = $ws.UsedRange 
            [void]$range.AutoFilter()
            [void]$range.EntireColumn.Autofit()
    
            $num = 1
            $dir = Split-Path $xlOut
            $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
            $nextname = $xlOut
            while (Test-Path $nextname) {
                $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
                $num++
            }
    
            $wb.SaveAs($nextname, 51)
        }
    
        END {
            $xl.Quit()
        
            $null = $ws, $wb, $xl | ForEach-Object {Release-Ref $_}
    
            # del $CSVFile
        }
    } ## Save as output XLS Routine

#endregion ----- Data Conversion ----

#region ----- Backup.Management JSON Calls ----
    Function Get-CovePartnerByName {
        ## Drop-in replacement for the deprecated GetPartnerInfo(name=...) method.
        ## https://developer.n-able.com/n-able-cove/docs/migrating-from-getpartnerinfo-to-getpartnerinfobyid
        ## NOTE: unlike GetPartnerInfo (EXACT, case-sensitive name match), this is a CASE-INSENSITIVE
        ## SUBSTRING match, and match counts below are "at least N" - the API appears to cap total
        ## filtered matches (~50 seen) regardless of childrenLimit, so a high count is not guaranteed complete.
        param (
            [Parameter(Mandatory=$True)]  [string]$PartnerName,
            [Parameter(Mandatory=$False)] [int]$PartnerId = $Script:AuthPartnerId   ## Defaults to the authenticated partner's own Id
        )

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'GetPartnerTree'
        $data.params = @{}
        $data.params.partnerId = [int]$PartnerId
        $data.params.fields = @(0,1,3,4)
        $data.params.childrenLimit = 5000
        $data.params.partnerFilter = @{ NamePattern = [String]$PartnerName }

        $Script:PartnerTree = Invoke-RestMethod -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data -depth 6) `
            -Uri $url

        ## Recursively flatten the (already name-filtered) tree, keeping only nodes whose own Name matches
        $Script:PartnerMatches = @()
        Function Get-PartnerTreeMatches ($Node, $Needle) {
            if ($Node.Info -and $Node.Info.Name -and ($Node.Info.Name -like "*$Needle*")) {
                $Script:PartnerMatches += New-Object -TypeName PSObject -Property @{
                    Id       = $Node.Info.Id
                    Name     = $Node.Info.Name
                    Level    = $Node.Info.Level
                    ParentId = $Node.Info.ParentId
                }
            }
            ForEach ($Child in $Node.Children) { Get-PartnerTreeMatches $Child $Needle }
        }

        if ($PartnerTree.PSObject.Properties['error']) {
            write-output "  $($PartnerTree.error.message)"
            DO{ $PartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'" }
            WHILE ($PartnerName.length -eq 0)
            Get-CovePartnerByName -PartnerName $PartnerName
            return
        }

        if ($PartnerTree.result.result) { Get-PartnerTreeMatches $PartnerTree.result.result $PartnerName }  ## GetPartnerTree double-wraps its result (result.result)

        ## Root/Sub-root/Distributor levels are not valid lookup targets (matches old GetPartnerInfo behavior)
        $Script:PartnerMatches = @($Script:PartnerMatches | Where-Object { $_.Level -notin @('Root', 'Sub-root', 'Distributor') })

        if ($Script:PartnerMatches.Count -eq 0) {
            Write-Output $Script:strLineSeparator
            Write-output "  No partner found matching '$PartnerName' (Root/Sub-root/Distributor level matches are not allowed)"
            DO{ $PartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'" }
            WHILE ($PartnerName.length -eq 0)
            Get-CovePartnerByName -PartnerName $PartnerName
            return
        }elseif ($Script:PartnerMatches.Count -gt 1) {
            Write-Output $Script:strLineSeparator
            Write-output "  At least $($Script:PartnerMatches.Count) partners match '$PartnerName' (API may cap results - refine your search if the one you want isn't listed). GetPartnerInfo would have silently picked one for you. Select one:"
            Write-Output $Script:strLineSeparator
            $Selected = $Script:PartnerMatches | Out-GridView -Title "Select Partner matching '$PartnerName'" -OutputMode Single
            if (-not $Selected) {
                Write-Output "  No partner selected."
                DO{ $PartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'" }
                WHILE ($PartnerName.length -eq 0)
                Get-CovePartnerByName -PartnerName $PartnerName
                return
            }
        }else{
            $Selected = $Script:PartnerMatches[0]
        }

        Get-CovePartnerInfoById -PartnerId $Selected.Id

    } ## get Partner Id/Level candidates via GetPartnerTree, disambiguate via GridView if needed

    Function Get-CovePartnerInfoById ($PartnerId) {
        ## Fetch full partner details (Uid, Level, etc.) once the Id is known - replaces GetPartnerInfo's result payload
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'GetPartnerInfoById'
        $data.params = @{}
        $data.params.partnerId = [int]$PartnerId

        $Script:Partner = Invoke-RestMethod -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data -depth 5) `
            -Uri $url

        if ($Partner.PSObject.Properties['error']) {
            Write-Output $Script:strLineSeparator
            Write-output "  $($Partner.error.message)"
            Write-Output $Script:strLineSeparator
            return
        }

        $Script:Uid          = [String]$Partner.result.result.Uid
        $Script:PartnerId    = [String]$Partner.result.result.Id
        $Script:PartnerName  = [String]$Partner.result.result.Name
        $Script:PartnerLevel = [String]$Partner.result.result.Level

        Write-Output $Script:strLineSeparator
        Write-output "  $Script:PartnerName - $Script:PartnerId - $Script:Uid"
        Write-Output $Script:strLineSeparator

    } ## get full Partner details by Id (GetPartnerInfoById)

    Function Get-Devices {

        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.visa = $Script:visa
        $data.method = 'EnumerateAccountStatistics'
        $data.params = @{}
        $data.params.query = @{}
        $data.params.query.PartnerId = [int]$PartnerId
        $data.params.query.SelectionMode = "PerInstallation"
        $data.params.query.Filter = $Filter1
        $data.params.query.Columns = @("AU","AR","AN","PF","LN","OP","OI","OS","OT","PD","AP","PN","AA843","MN","TS","EI","IP","MO","MF","CD","VN","II","IM","RTG","RP")
        $data.params.query.OrderBy = "T7 ASC"
        $data.params.query.StartRecordNumber = 0
        $data.params.query.RecordsCount = $devicecount
        $data.params.query.Totals = @("COUNT(AT==1)","SUM(T3)","SUM(US)")
    
        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data -depth 6) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            $Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Devices = $webrequest | convertfrom-json

        Write-Output "  Requesting details for $($Devices.result.result.count) devices."
        Write-Output "  Please be patient, this could take some time."
        Write-Output $Script:strLineSeparator

        $Script:DeviceDetail = @()

        ForEach ( $DeviceResult in $Devices.result.result ) {

            Get-VisaTime

            $Script:DeviceDetail += New-Object -TypeName PSObject -Property @{ AccountID      = [String]$DeviceResult.AccountId;
                                                                        PartnerID      = [string]$DeviceResult.PartnerId;
                                                                        DeviceName     = $DeviceResult.Settings.AN -join '' ;                                                                    
                                                                        PartnerName    = $DeviceResult.Settings.AR -join '' ;
                                                                        Reference      = $DeviceResult.Settings.PF -join '' ;
                                                                        DataSources    = $DeviceResult.Settings.AP -join '' ;                                                                
                                                                        Location       = $DeviceResult.Settings.LN -join '' ;
                                                                        Notes          = $DeviceResult.Settings.AA843 -join '' ;
                                                                        Product        = $DeviceResult.Settings.PN -join '' ;
                                                                        ProductID      = $DeviceResult.Settings.PD -join '' ;
                                                                        Profile        = $DeviceResult.Settings.OP -join '' ;
                                                                        OS             = $DeviceResult.Settings.OS -join '' ;
                                                                        MachineName    = $DeviceResult.Settings.MN -join '' ;  
                                                                        MFG_Name       = $DeviceResult.Settings.MF -join '' ;
                                                                        MFG_Model      = $DeviceResult.Settings.MO -join '' ;
                                                                        IP             = $DeviceResult.Settings.IP -join '' ;  
                                                                        Ext_IP         = $DeviceResult.Settings.EI -join '' ;
                                                                        Creation       = Convert-UnixTimeToDateTime ($DeviceResult.Settings.CD -join '') ;
                                                                        TimeStamp      = Convert-UnixTimeToDateTime ($DeviceResult.Settings.TS -join '') ;  
                                                                        ClientVersion  = $DeviceResult.Settings.VN -join '' ;
                                                                        InstallID      = $DeviceResult.Settings.II -join '' ;  
                                                                        InstallMode    = $DeviceResult.Settings.IM -join '' ;
                                                                        LastRestore    = Convert-UnixTimeToDateTime ($DeviceResult.Settings.RTG -join '') ;  
                                                                        ProfileID      = $DeviceResult.Settings.OI -join '' 
                                                                    }
        }     


    }  ## Enumerate devices under specified Sub-distributor or lower partner

#endregion ----- Backup.Management JSON Calls ----

#endregion ----- Functions ----

    
    Send-APICredentialsCookie

    Write-Output $Script:strLineSeparator
    Write-Output "" 

    Get-CovePartnerInfoById -PartnerId $Script:AuthPartnerId   ## learn the authenticated partner's own Level before deciding whether a name prompt is required

    if ($Script:PartnerLevel -in @('Root','Sub-root','Distributor')) {
        Write-Output $Script:strLineSeparator
        Write-output "  Authenticated partner '$Script:PartnerName' is Level '$Script:PartnerLevel' - restricted, not a valid lookup target for this script."
        DO{ $PartnerName = Read-Host "  Enter Customer/ Partner display name (or partial name) to lookup i.e. 'Acme'" }
        WHILE ($PartnerName.length -eq 0)
    }else{
        $PartnerName = $Script:PartnerName   ## authenticated partner is a valid lookup target itself
    }

    Get-CovePartnerByName -PartnerName $PartnerName

    $filter1 = "AT == 1"  ## Exclude M365 devices
   
    Get-Devices

    $DeviceDetail = $DeviceDetail | select-object PartnerID,PartnerName,Reference,AccountId,DeviceName,MachineName,Creation,LastRestore,TimeStamp,DataSources,ClientVersion,MFG_Name,MFG_Model,OS,InstallMode,InstallID,Ext_IP,IP,location,productID,Profile,Notes | Sort-Object Partnername,devicename,timestamp 

    ## Display GridView (Required -GridView Parameter)

    if ($GridView) { $DeviceDetail | out-gridview -Title "Device Installation Audit" }

    ## Export CSV

    $Script:csvoutputfile = "$ExportPath\$($CurrentDate)_Backup_Install_Audit_$($Partnername -replace(`" \(.*\)`",`"`") -replace(`"[^a-zA-Z_0-9]`",`"`"))_$($PartnerId).csv"
    $DeviceDetail | Export-Csv -Path $csvoutputfile -delimiter "$Delimiter" -NoTypeInformation -Encoding UTF8 -append

    ## Generate XLS from CSV

    $xlsoutputfile = $csvoutputfile.Replace("csv","xlsx")
    Save-CSVasExcel $csvoutputfile
        
    Write-output $Script:strLineSeparator

    ## Launch CSV or XLS if Excel is installed  (Required -Launch Parameter)
        
    if ($Launch) {
        If (test-path HKLM:SOFTWARE\Classes\Excel.Application) { 
            Start-Process "$xlsoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening XLS file"
            }else{
            Start-Process "$csvoutputfile"
            Write-output $Script:strLineSeparator
            Write-Output "  Opening CSV file"
            Write-output $Script:strLineSeparator            
            }
        }
    Write-output $Script:strLineSeparator
    Write-Output "  CSV Path = $csvoutputfile"
    Write-Output "  XLS Path = $xlsoutputfile"
    Write-Output ""

    Start-Sleep -seconds 10

    