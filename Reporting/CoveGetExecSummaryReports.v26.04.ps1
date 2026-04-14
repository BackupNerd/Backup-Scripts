# ----- About: ----
    # Cove Data Protection | Get Exec Summary Reports
    # Revision v01.0 - 2026-04-14
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
    # For use with the Standalone edition of N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    #
    # Authenticate to https://backup.management console
    # If login level is not Reseller or ServiceOrganization, prompt to look up a qualifying partner
    # Enumerate child partners via EnumerateChildPartners JSON-RPC
    #   - ServiceOrganization level partners are automatically selected
    # For each ServiceOrganization, enumerate its child partners
    # Query the Report Manager REST API for existing sub-partner reports since -Since date
    # For child partners with no existing report, trigger generation via the generate endpoint
    #   - Generation requests are capped by -MaxGenerate (default 1, 0 = unlimited)
    # Poll until all InProgress reports reach a terminal state (Ready or NoData)
    # Obtain a download token per Ready report via the Downloads Proxy REST API
    # Download all Ready-state PDF reports to -ExportPath
    #
    # Use the -Since parameter to set the report period (month) to query
    #   Default: first day of the prior calendar month
    # Use the -ExportPath parameter to override the PDF download root folder
    #   Default: .\Exec_Summary_Reports_<yyyy-MM>\ beside the script
    # Use the -MaxGenerate parameter to cap generation requests per run
    #   Default: 6  |  Set to 0 for unlimited
    # Use the -ClearCredentials switch to wipe and re-enter stored API credentials
    #
    # REST endpoints used:
    #   JSON-RPC POST https://api.backup.management/jsonapi
    #              (Login, GetPartnerInfo, EnumerateChildPartners)
    #   GET  https://api.backup.management/report-manager/partners/{nri}/subpartners/reports/{iso8601}
    #   POST https://api.backup.management/report-manager/partners/{nri}/reports/{iso8601}/generate
    #   POST https://api.backup.management/notifications-publisher-api/negotiate
    #   POST https://api.backup.management/downloads-proxy-api/files/{fileId}/download-tokens
    #   GET  https://downloads-proxy-api.cloudbackup.management/downloads?token={token}
    #
# -----------------------------------------------------------#>  ## Behavior

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)][datetime]$Since,                               ## Earliest report date (default = first day of prior month)
    [Parameter(Mandatory=$False)][string]$ExportPath = "",                        ## Root download folder (default: script directory)
    [Parameter(Mandatory=$False)][int]$MaxGenerate = 6,                          ## Max report generation requests per run (0 = unlimited)
    [Parameter(Mandatory=$False)][switch]$ClearCredentials                       ## Wipe stored credentials
)

#region ----- Environment, Variables, Names and Paths ----

Clear-Host
#Requires -Version 5.1

$ConsoleTitle = "Cove Data Protection | Get Exec Summary Reports"
$host.UI.RawUI.WindowTitle = $ConsoleTitle

Set-Location $PSScriptRoot
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$CurrentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ShortDate   = Get-Date -Format "yyyy-MM-dd"

if (-not $Since) {
    $d = (Get-Date -Day 1).AddMonths(-1)
    $Since = Get-Date -Year $d.Year -Month $d.Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
}

$Script:strLineSeparator = "  ---------"

$urlJSON     = 'https://api.backup.management/jsonapi'
$urlReports  = 'https://api.backup.management/report-manager'
$urlFiles    = 'https://api.backup.management/downloads-proxy-api'
$urlDownload = 'https://downloads-proxy-api.cloudbackup.management/downloads'

if (-not $ExportPath) { $ExportPath = $PSScriptRoot }            ## PS5: $PSScriptRoot is unavailable in param defaults
$ExportPath  = Join-Path -Path $ExportPath -ChildPath "Exec_Summary_Reports_$($Since.ToString('yyyy-MM'))"
New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null

Write-Output "`n  $ConsoleTitle"
Write-Output $Script:strLineSeparator
Write-Output "  Current Parameters:"
Write-Output "  -Since       = $($Since.ToString('yyyy-MM-dd'))"
Write-Output "  -ExportPath  = $ExportPath"
Write-Output "  -MaxGenerate = $(if ($MaxGenerate -eq 0) { 'Unlimited' } else { $MaxGenerate })"
Write-Output $Script:strLineSeparator

#endregion ----- Environment, Variables, Names and Paths ----

#region ----- Functions ----

#region ----- Authentication ----

Function Set-APICredentials {
    Write-Output $Script:strLineSeparator
    Write-Output "  Setting Backup API Credentials"
    if (Test-Path $APIcredpath) {
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential Path Present"
    } else {
        New-Item -ItemType Directory -Path $APIcredpath -Force | Out-Null
    }
    Write-Output "  Enter Exact, Case Sensitive Partner Name for N-able | Cove | Backup.Management API"
    Write-Output "  Example: 'Acme, Inc (bob@acme.net)'"
    DO { $Script:PartnerName = Read-Host "  Enter Login Partner Name" }
    WHILE ($PartnerName.Length -eq 0)
    $PartnerName | Out-File $APIcredfile

    $BackupCred = Get-Credential -Message 'Enter Login Email and Password for N-able | Cove | Backup.Management API'
    $BackupCred | Add-Member -MemberType NoteProperty -Name PartnerName -Value "$PartnerName"

    $BackupCred.UserName | Out-File -Append $APIcredfile
    $BackupCred.Password | ConvertFrom-SecureString | Out-File -Append $APIcredfile

    Start-Sleep -Milliseconds 300
    Send-APICredentialsCookie
}  ## Set API credentials if not present

Function Get-APICredentials {
    $Script:True_path   = "C:\ProgramData\MXB\"
    if (-not (Test-Path -Path $Script:True_path)) {
        New-Item -ItemType Directory -Path $Script:True_path -Force | Out-Null
    }
    $Script:APIcredfile = Join-Path -Path $True_Path -ChildPath "$env:computername API_Credentials.Secure.txt"
    $Script:APIcredpath = Split-Path -Path $APIcredfile

    if (($ClearCredentials) -and (Test-Path $APIcredfile)) {
        Remove-Item -Path $Script:APIcredfile
        $ClearCredentials = $Null
        Write-Output $Script:strLineSeparator
        Write-Output "  Backup API Credential File Cleared"
        Send-APICredentialsCookie
    } else {
        Write-Output $Script:strLineSeparator
        Write-Output "  Getting Backup API Credentials"
        if (Test-Path $APIcredfile) {
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Present"
            $APIcredentials   = Get-Content $APIcredfile
            $Script:cred0     = [string]$APIcredentials[0]
            $Script:cred1     = [string]$APIcredentials[1]
            $Script:cred2     = $APIcredentials[2] | ConvertTo-SecureString
            $Script:cred2     = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:cred2))
            Write-Output $Script:strLineSeparator
            Write-Output "  Stored Backup API Partner  = $Script:cred0"
            Write-Output "  Stored Backup API User     = $Script:cred1"
            Write-Output "  Stored Backup API Password = Encrypted`n"
        } else {
            Write-Output $Script:strLineSeparator
            Write-Output "  Backup API Credential File Not Present"
            Set-APICredentials
        }
    }
}  ## Get API credentials if present

Function Send-APICredentialsCookie {
    Get-APICredentials

    $data = @{
        jsonrpc = '2.0'
        id      = '2'
        method  = 'Login'
        params  = @{
            partner  = $Script:cred0
            username = $Script:cred1
            password = $Script:cred2
        }
    }

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data) `
        -Uri $urlJSON `
        -TimeoutSec 30 `
        -SessionVariable Script:websession `
        -UseBasicParsing

    $Script:websession   = $websession
    $Script:Authenticate = $webrequest | ConvertFrom-Json

    if ($Script:Authenticate.visa) {
        $Script:visa   = $Script:Authenticate.visa
        $Script:UserId = $Script:Authenticate.result.result.id
        $Script:cred2  = $null  # Zero plaintext password from memory — re-read from encrypted file on next re-auth
    } else {
        Write-Output $Script:strLineSeparator
        Write-Warning "`n  $($Script:Authenticate.error.message)"
        Write-Output "  Authentication Failed: Please confirm your Backup.Management Partner Name and Credentials"
        Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lock out your user account"
        Write-Output $Script:strLineSeparator
        Set-APICredentials
    }
}  ## Use Backup.Management credentials to Authenticate

Function Visa-Check {
    if ($Script:visa) {
        $VisaTime = (Convert-UnixTimeToDateTime ([int]$Script:visa.split("-")[3]))
        if ($VisaTime -lt (Get-Date).ToUniversalTime().AddMinutes(-10)) {
            Send-APICredentialsCookie
        }
    }
}  ## Recheck remaining Visa time and re-authenticate if needed

#endregion ----- Authentication ----

#region ----- Data Conversion ----

Function Convert-UnixTimeToDateTime ([long]$inputUnixTime) {
    if ($inputUnixTime -gt 0) {
        $epoch = (Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()
        return $epoch.AddSeconds($inputUnixTime)
    } else { return "" }
}  ## Convert epoch time to DateTime

#endregion ----- Data Conversion ----

#region ----- JSON-RPC API Calls ----

Function Send-GetPartnerInfo ([string]$PartnerName) {
    $data = @{
        jsonrpc = '2.0'
        id      = '2'
        visa    = $Script:visa
        method  = 'GetPartnerInfo'
        params  = @{ name = [string]$PartnerName }
    }

    $webrequest        = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -Depth 5) `
        -Uri $urlJSON `
        -SessionVariable Script:websession `
        -UseBasicParsing
    $Script:websession = $websession
    $Script:Partner    = $webrequest | ConvertFrom-Json
    $Script:visa       = $Script:Partner.visa

    if ($Script:Partner.error) {
        Write-Warning "  GetPartnerInfo Error: $($Script:Partner.error.message)"
        return
    }

    [int]$Script:PartnerId      = $Script:Partner.result.result.Id
    [string]$Script:Level       = $Script:Partner.result.result.Level
    [string]$Script:PartnerName = $Script:Partner.result.result.Name
    [string]$Script:Uid         = $Script:Partner.result.result.Uid
}  ## Get PartnerID, Level, Name and UID

Function Send-EnumerateChildPartners ([int]$PartnerId) {
    Write-Output $Script:strLineSeparator
    Write-Output "  Enumerating Child Partners for: $Script:PartnerName (ID $PartnerId)"
    Visa-Check

    $data = @{
        jsonrpc = '2.0'
        id      = 'jsonrpc'
        visa    = $Script:visa
        method  = 'EnumerateChildPartners'
        params  = @{
            partnerId     = $PartnerId
            fields        = @(0, 1, 3, 4, 5, 8, 11, 12, 18, 21)
            range         = @{ Offset = 0; Size = 10000 }
            partnerFilter = @{ SortOrder = "ByLevelAndName" }
        }
    }

    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -Depth 6) `
        -Uri $urlJSON `
        -SessionVariable Script:websession `
        -UseBasicParsing
    $Script:websession = $websession
    $response          = $webrequest | ConvertFrom-Json
    $Script:visa       = $response.visa

    if ($response.error) {
        Write-Output $Script:strLineSeparator
        Write-Warning "  EnumerateChildPartners Error: $($response.error.message)"
        Exit
    }

    # Flatten Children[].Info into a simple list with ActualChildCount
    $Script:ChildPartnersList = $response.result.result.Children | ForEach-Object {
        $info = $_.Info
        [PSCustomObject]@{
            Id               = $info.Id
            Name             = $info.Name
            Level            = $info.Level
            State            = $info.State
            ServiceType      = $info.ServiceType
            ActualChildCount = $_.ActualChildCount
            Reference        = $info.ExternalCode
            LocationId       = $info.LocationId
            ParentId         = $info.ParentId
            TrialExpires     = if ($info.TrialExpirationTime -gt 0) {
                                   Convert-UnixTimeToDateTime $info.TrialExpirationTime
                               } else { "" }
            Uid              = $info.Uid
        }
    }

    $soCount = ($Script:ChildPartnersList | Where-Object { $_.Level -eq 'ServiceOrganization' }).Count
    Write-Output "  Found $($Script:ChildPartnersList.Count) child partner(s)  |  ServiceOrganization: $soCount"
    Write-Output $Script:strLineSeparator
}  ## EnumerateChildPartners JSON-RPC call

#endregion ----- JSON-RPC API Calls ----

#region ----- Report Manager REST Calls ----

Function Get-SubPartnerReports ([int]$ForPartnerId, [datetime]$SinceDate) {
    Visa-Check

    # API requires strictly first-of-month with all time components zeroed
    $sinceISO = "$($SinceDate.Year)-$('{0:D2}' -f $SinceDate.Month)-01T00:00:00.000Z"
    $nri      = "nable:cove::partner:$ForPartnerId"
    # NRI colons must NOT be percent-encoded in the URL path
    $url      = "$urlReports/partners/$nri/subpartners/reports/$sinceISO"

    Write-Host "  Querying Report Manager: $nri  since $sinceISO"

    $headers = @{
        'Authorization' = "Bearer $Script:visa"
        'Accept'        = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Method GET `
            -Uri $url `
            -Headers $headers
        return $response.reports
    } catch {
        # Extract readable message from the exception (PS7 embeds response body)
        $errMsg = $_.Exception.Message
        try { $errBody = ($_ | Get-Error).Exception.Response | ForEach-Object { $null } } catch {}
        Write-Warning "  Report Manager query failed for partner $ForPartnerId : $errMsg"
        return $null
    }
}  ## GET available sub-partner reports from Report Manager API

Function Get-DownloadToken ([string]$FileId) {
    Visa-Check

    $url     = "$urlFiles/files/$FileId/download-tokens"
    $headers = @{
        'Authorization' = "Bearer $Script:visa"
        'Accept'        = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Method POST `
            -Uri $url `
            -Headers $headers `
            -ContentType 'application/json'
        return $response
    } catch {
        $sc   = try { $_.Exception.Response.StatusCode.value__ } catch { '?' }
        $body = $_.ErrorDetails.Message
        if (-not $body) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $stream.Position = 0
                $body = [System.IO.StreamReader]::new($stream).ReadToEnd()
            } catch {}
        }
        Write-Warning "  download-tokens POST failed for file $FileId : HTTP $sc"
        if ($body) { Write-Warning "  Response body: $body" }
        return $null
    }
}  ## POST to downloads-proxy-api to obtain a per-file download token

Function Save-ReportFile ([string]$FileId, [string]$ReportPartnerNri, [string]$SaveAsName) {
    Write-Output "  Obtaining download token for file: $FileId"
    $tokenResponse = Get-DownloadToken -FileId $FileId

    if (-not $tokenResponse) {
        Write-Warning "  No token response for $FileId — skipping"
        return
    }

    # The download token (field: "token") is used as Bearer auth on GET /files/{fileId}/download
    $downloadToken = $null
    if     ($tokenResponse.token)         { $downloadToken = $tokenResponse.token }
    elseif ($tokenResponse.downloadToken) { $downloadToken = $tokenResponse.downloadToken }

    # Always dump token response so we can verify the field/URL structure
    #Write-Output "  Token response (raw):"
    #$tokenResponse | ConvertTo-Json -Depth 5 | Write-Host -ForegroundColor Cyan

    if (-not $downloadToken) {
        Write-Warning "  Cannot find token field in download-tokens response — skipping"
        return
    }
    # Build a clean output filename: ChildPartnerName_reportPartnerId_MonthName-Year.pdf  (matches Cove GUI)
    $reportPartnerId = ($ReportPartnerNri -split ':')[-1] -replace '[^\d]', ''  # Digits only — prevent path traversal from API response
    $reportMonth   = $Since.ToString('MMMM-yyyy')
    $cleanSaveName = $SaveAsName -replace '[^a-zA-Z0-9_\-]', '_' -replace '_+', '_' -replace '^_|_$', ''
    $outFile       = Join-Path -Path $ExportPath -ChildPath "${cleanSaveName}_${reportPartnerId}_${reportMonth}.pdf"

    # Token is passed as ?token= query param on the cloudbackup.management download domain
    # Pattern confirmed from browser DevTools: GET https://downloads-proxy-api.cloudbackup.management/downloads?token={urlEncodedToken}
    $tokEnc      = [uri]::EscapeDataString($downloadToken)
    $downloadUrl = "${urlDownload}?token=$tokEnc"

    Write-Output "  Downloading → $(Split-Path $outFile -Leaf)"

    try {
        Invoke-WebRequest -Method GET `
            -Uri $downloadUrl `
            -OutFile $outFile `
            -UseBasicParsing `
            -ErrorAction Stop
        Write-Output "  ✅  Saved: $outFile"
        return $true
    } catch {
        Write-Warning "  Download failed for $FileId : $($_.Exception.Message)"
        return $false
    }
}  ## Download a single report PDF and save to ExportPath

Function Get-NotificationsSession {
    ## POST notifications-publisher-api/negotiate → { sessionId, token }
    ## The sessionId is required in the report generate body
    Visa-Check
    $url     = 'https://api.backup.management/notifications-publisher-api/negotiate'
    $headers = @{ Authorization = "Bearer $Script:visa" }
    try {
        $resp = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -ContentType 'application/json' -EA Stop
        Write-Output "  Notifications session: $($resp.sessionId)"
        return [string]$resp.sessionId
    } catch {
        $sc   = try { $_.Exception.Response.StatusCode.value__ } catch { '?' }
        $body = $_.ErrorDetails.Message
        Write-Verbose "  negotiate failed HTTP $sc — generate will proceed without sessionId"
        return $null
    }
}  ## Negotiate a notifications session required by the generate endpoint

Function Get-SOChildPartners ([int]$PartnerId) {
    ## Returns the direct children of a partner via EnumerateChildPartners
    ## without overwriting $Script:ChildPartnersList
    Visa-Check
    $data = @{
        jsonrpc = '2.0'
        id      = 'jsonrpc'
        visa    = $Script:visa
        method  = 'EnumerateChildPartners'
        params  = @{
            partnerId     = $PartnerId
            fields        = @(0, 1, 3, 4, 5, 8, 11, 12, 18, 21)
            range         = @{ Offset = 0; Size = 10000 }
            partnerFilter = @{ SortOrder = "ByLevelAndName" }
        }
    }
    $webrequest = Invoke-WebRequest -Method POST `
        -ContentType 'application/json; charset=utf-8' `
        -Body (ConvertTo-Json $data -Depth 6) `
        -Uri $urlJSON `
        -SessionVariable Script:websession `
        -UseBasicParsing
    $Script:websession = $websession
    $response = $webrequest | ConvertFrom-Json
    $Script:visa = $response.visa
    if ($response.error) {
        Write-Warning "  EnumerateChildPartners Error for $PartnerId : $($response.error.message)"
        return @()
    }
    return $response.result.result.Children | ForEach-Object {
        $info = $_.Info
        [PSCustomObject]@{
            Id   = $info.Id
            Name = $info.Name
            Nri  = "nable:cove::partner:$($info.Id)"
        }
    }
}  ## EnumerateChildPartners for an SO — returns list without overwriting global state

Function Request-ReportGeneration ([string]$PartnerNri, [datetime]$SinceDate) {
    ## POST report-manager/partners/{nri_encoded}/reports/{date_encoded}/generate → 202 Accepted
    Visa-Check
    # Both the NRI and the date ISO string must be fully URL-encoded in the path
    # (confirmed from browser DevTools: 2026-02-01T00%3A00%3A00Z)
    $sinceISO  = "$($SinceDate.Year)-$('{0:D2}' -f $SinceDate.Month)-01T00:00:00Z"
    $nriEnc    = [uri]::EscapeDataString($PartnerNri)
    $dateEnc   = [uri]::EscapeDataString($sinceISO)
    $url       = "$urlReports/partners/$nriEnc/reports/$dateEnc/generate"
    $headers   = @{ Authorization = "Bearer $Script:visa" }
    Write-Output "  POST: $url"
    try {
        # Flat body as confirmed from browser DevTools:
        # {"isEmailDeliveryRequired":false,"sessionId":"..."}
        $genBody = [ordered]@{ isEmailDeliveryRequired = $false }
        if ($Script:NotificationsSessionId) { $genBody['sessionId'] = [string]$Script:NotificationsSessionId }
        $body = ConvertTo-Json $genBody -Compress -Depth 2
        $resp = Invoke-WebRequest -Method POST `
            -Uri $url `
            -Headers $headers `
            -Body $body `
            -ContentType 'application/json' `
            -UseBasicParsing `
            -ErrorAction Stop
        Write-Output "  ⚡  Generation triggered: $PartnerNri  (HTTP $($resp.StatusCode))"
        return $true
    } catch {
        $sc   = try { $_.Exception.Response.StatusCode.value__ } catch { '?' }
        # Try all available sources for the response body
        $body = $_.ErrorDetails.Message
        if (-not $body) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $stream.Position = 0
                $body = [System.IO.StreamReader]::new($stream).ReadToEnd()
            } catch {}
        }
        Write-Warning "  Generate failed for $PartnerNri : HTTP $sc"
        return $false
    }
}  ## Trigger report generation for a specific partner NRI

Function Wait-ForReports ([int]$SoPartnerId, [datetime]$SinceDate, [string[]]$PendingNris, [int]$TimeoutSeconds = 300) {
    ## Polls subpartners/reports until all $PendingNris leave the InProgress state, then returns the final list
    Write-Output "  ⏳  Waiting for $($PendingNris.Count) report(s) to finish generating (timeout: ${TimeoutSeconds}s)..."
    $elapsed  = 0
    $interval = 20
    do {
        Start-Sleep -Seconds $interval
        $elapsed += $interval
        Visa-Check
        $current = Get-SubPartnerReports -ForPartnerId $SoPartnerId -SinceDate $SinceDate
        $still   = @($PendingNris | Where-Object {
            $nri = $_
            $match = $current | Where-Object { $_.partnerNri -eq $nri }
            (-not $match) -or ($match.state -eq 'InProgress')
        })
        $doneCount = $PendingNris.Count - $still.Count
        Write-Output "  ⏳  ${elapsed}s elapsed — $doneCount/$($PendingNris.Count) done$(if ($still) { ';  still pending: ' + ($still -join ', ') })"
        if ($still.Count -eq 0) { break }
    } while ($elapsed -lt $TimeoutSeconds)
    if ($elapsed -ge $TimeoutSeconds) {
        Write-Warning "  Timeout reached — proceeding with reports that are currently Ready"
    }
    return Get-SubPartnerReports -ForPartnerId $SoPartnerId -SinceDate $SinceDate
}  ## Poll until all triggered reports leave InProgress state

#endregion ----- Report Manager REST Calls ----

#endregion ----- Functions ----

#region ----- Main Execution ----

# ── Step 1: Authenticate ──────────────────────────────────────
Send-APICredentialsCookie

# ── Step 2: Determine the login partner's account level ───────
Send-GetPartnerInfo $Script:cred0

Write-Output $Script:strLineSeparator
Write-Output "  Login Partner : $Script:PartnerName"
Write-Output "  Partner Level : $Script:Level"
Write-Output "  Partner ID    : $Script:PartnerId"
Write-Output $Script:strLineSeparator

# ── Step 3: If not already at an actionable level, ask for a partner ──
$ActionableLevels = @("Reseller", "ServiceOrganization", "SubDistributor", "Distributor")

if ($Script:Level -notin $ActionableLevels) {
    do {
        if ($Script:Partner.error) { Write-Warning "  $($Script:Partner.error.message)" }
        if ($Script:Level -notin $ActionableLevels) {
            Write-Output "  Login level '$Script:Level' cannot enumerate child partners directly."
            Write-Output "  Please enter a Reseller or ServiceOrganization partner name."
        }
        Write-Output $Script:strLineSeparator
        $lookupName = Read-Host "  Enter EXACT Case Sensitive Partner Name (e.g. 'Acme MSP (admin@acme.com)')"
        Send-GetPartnerInfo $lookupName
        Write-Output "  Found: $Script:PartnerName  |  Level: $Script:Level  |  ID: $Script:PartnerId"
    } while ($Script:Level -notin $ActionableLevels -or $Script:Partner.error)
}

# ── Step 4: Enumerate children, auto-select ServiceOrganization partners ──
Send-EnumerateChildPartners -PartnerId $Script:PartnerId

if (-not $Script:ChildPartnersList -or $Script:ChildPartnersList.Count -eq 0) {
    Write-Warning "  No child partners found under $Script:PartnerName. Exiting."
    Exit
}

# Auto-select ServiceOrganization level partners — no GridView needed
$Script:SelectedPartners = @($Script:ChildPartnersList | Where-Object { $_.Level -eq 'ServiceOrganization' })

if ($Script:SelectedPartners.Count -eq 0) {
    Write-Output $Script:strLineSeparator
    Write-Output "  No ServiceOrganization partners found under $Script:PartnerName."
    Write-Output "  Falling back: processing end customers directly under the Reseller."
    # Treat the reseller itself as the single "SO" — its direct children are the end customers
    $Script:SelectedPartners = @([PSCustomObject]@{ Name = $Script:PartnerName; Id = $Script:PartnerId })
}

Write-Output "  Auto-selected $($Script:SelectedPartners.Count) ServiceOrganization partner(s):"
$Script:SelectedPartners | ForEach-Object { Write-Output "    • $($_.Name)  (ID $($_.Id))" }

# ── Step 5: Negotiate a notifications session, then for each SO generate + download ──
$Script:NotificationsSessionId = Get-NotificationsSession

$totalReports    = 0
$downloadedCount = 0
$skippedCount    = 0
$generatedCount  = 0

foreach ($soPartner in $Script:SelectedPartners) {
    Write-Output $Script:strLineSeparator
    Write-Output "  ► Processing SO: $($soPartner.Name)  (ID $($soPartner.Id))"

    # Get the SO's direct children to identify who needs a report
    $soChildren = Get-SOChildPartners -PartnerId $soPartner.Id
    Write-Output "  Found $($soChildren.Count) child partner(s) under this SO"

    # Get current report states for all sub-partners of this SO
    $reports = Get-SubPartnerReports -ForPartnerId $soPartner.Id -SinceDate $Since

    # Build NRI → report lookup
    $reportLookup = @{}
    if ($reports) { foreach ($r in $reports) { if ($r.partnerNri) { $reportLookup[$r.partnerNri] = $r } } }

    # Identify children that need generation triggered
    $pendingNris = [System.Collections.Generic.List[string]]::new()
    foreach ($child in $soChildren) {
        $nri = $child.Nri
        if (-not $reportLookup.ContainsKey($nri)) {
            # No report exists at all — trigger generation (subject to -MaxGenerate cap)
            if ($MaxGenerate -gt 0 -and $generatedCount -ge $MaxGenerate) {
                Write-Output "  No report for $($child.Name) — skipping (MaxGenerate limit of $MaxGenerate reached)"
            } else {
                Write-Output "  No report for $($child.Name) — requesting generation"
                if (Request-ReportGeneration -PartnerNri $nri -SinceDate $Since) {
                    $pendingNris.Add($nri)
                    $generatedCount++
                } else {
                    Write-Output "  Skipping wait for $($child.Name) (generate failed)"
                }
            }
        } elseif ($reportLookup[$nri].state -eq 'InProgress') {
            # Already generating — just wait for it
            Write-Output "  InProgress for $($child.Name) — will wait for completion"
            $pendingNris.Add($nri)
        } elseif ($reportLookup[$nri].state -eq 'NoData') {
            Write-Output "  NoData for $($child.Name) — no data available for this period, skipping"
        }
        # Ready — handled in download loop below
    }

    # Wait for all pending (newly triggered + already InProgress) to reach terminal state
    if ($pendingNris.Count -gt 0) {
        $reports = Wait-ForReports -SoPartnerId $soPartner.Id -SinceDate $Since -PendingNris $pendingNris.ToArray()
    }

    if (-not $reports -or $reports.Count -eq 0) {
        Write-Output "  No reports available — skipping"
        continue
    }

    Write-Output "  Processing $($reports.Count) report entry/entries"
    $totalReports += $reports.Count

    # Build NRI → child name lookup for filename construction
    $childNameByNri = @{}
    foreach ($c in $soChildren) { $childNameByNri[$c.Nri] = $c.Name }

    foreach ($report in $reports) {
        # Skip empty stubs returned while reports are still generating
        if (-not $report.partnerNri) { $skippedCount++; continue }

        Write-Output ""
        Write-Output "  Partner NRI : $($report.partnerNri)"
        Write-Output "  State       : $($report.state)"
        Write-Output "  File ID     : $($report.fileId)"
        Write-Output "  Created UTC : $($report.createdAtUtc)"

        if ($report.state -ne 'Ready') {
            Write-Output "  ⚠️  State is '$($report.state)' — skipping download"
            $skippedCount++
            continue
        }

        $childName = if ($childNameByNri.ContainsKey($report.partnerNri)) { $childNameByNri[$report.partnerNri] } else { $soPartner.Name }
        $saved = Save-ReportFile `
            -FileId           $report.fileId `
            -ReportPartnerNri $report.partnerNri `
            -SaveAsName       $childName

        if ($saved) { $downloadedCount++ } else { $skippedCount++ }
    }
}

# ── Summary ───────────────────────────────────────────────────
$iw = 50   # inner width between the two ║ borders

function Write-SRow {
    param([string]$Label, [string]$Value = "", [ConsoleColor]$LC = [ConsoleColor]::Gray, [ConsoleColor]$VC = [ConsoleColor]::White)
    $lp  = "  $Label"
    $vp  = if ($Value) { " : $Value" } else { "" }
    $pad = " " * ([Math]::Max(0, $iw - $lp.Length - $vp.Length))
    Write-Host -NoNewline "  ║" -ForegroundColor Cyan
    Write-Host -NoNewline $lp  -ForegroundColor $LC
    if ($vp) { Write-Host -NoNewline $vp -ForegroundColor $VC }
    Write-Host "$pad║" -ForegroundColor Cyan
}

$top = "  ╔" + ("═" * $iw) + "╗"
$div = "  ╠" + ("═" * $iw) + "╣"
$bot = "  ╚" + ("═" * $iw) + "╝"

# Centered title
$t    = "  Cove | Executive Summary Report Download"
$tpad = " " * ([Math]::Max(0, $iw - $t.Length))

Write-Host ""
Write-Host $Script:strLineSeparator -ForegroundColor DarkGray
Write-Host $top -ForegroundColor Cyan
Write-Host -NoNewline "  ║" -ForegroundColor Cyan
Write-Host -NoNewline $t   -ForegroundColor White
Write-Host "$tpad║"         -ForegroundColor Cyan
Write-Host $div -ForegroundColor Cyan
Write-SRow "Partner          " $Script:PartnerName                         -VC Cyan
Write-Host $div -ForegroundColor Cyan
Write-SRow "SOs Processed    " $Script:SelectedPartners.Count              -VC Cyan
Write-SRow "Reports Found    " $totalReports                               -VC White
Write-SRow "Generated        " "$generatedCount$(if ($MaxGenerate -gt 0) { " (limit: $MaxGenerate)" })"  -VC $(if ($generatedCount -gt 0) {"Yellow"} else {"DarkGray"})
Write-SRow "Downloaded       " $downloadedCount -VC $(if ($downloadedCount -gt 0) {"Green" } else {"DarkGray"})
Write-SRow "Skipped          " $skippedCount    -VC $(if ($skippedCount    -gt 0) {"Yellow"} else {"DarkGray"})
Write-Host $div -ForegroundColor Cyan
Write-SRow "Period           " $Since.ToString("MMMM yyyy")                -VC Cyan
Write-SRow "Output Folder    " (Split-Path $ExportPath -Leaf)              -VC Yellow
Write-Host $bot -ForegroundColor Cyan
Write-Host $Script:strLineSeparator -ForegroundColor DarkGray

Start-Sleep -Seconds 3

#endregion ----- Main Execution ----
