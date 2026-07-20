<# ----- About: ----
    # Cove Data Protection - Get Local Backup Selections (v2)
    # Revision v2 - 2026-07-06
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Uses InAgent token auth + local JSON-RPC API instead of ClientTool CLI
# -----------------------------------------------------------#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)]
    [string[]]$DataSources = @("VssExchange","FileSystem","MySql","NetworkShares","Oracle","VssSystemState","SystemState","LinuxSystemState","VssHyperV","VssMsSql","VssSharePoint","VMWare"),
    [Parameter(Mandatory=$False)]
    [string]$ClientToolPath = "C:\Program Files\Backup Manager\ClientTool.exe",
    [Parameter(Mandatory=$False)]
    [string]$LocalUrl       = "http://localhost:5000/jsonrpcv1",
    [Parameter(Mandatory=$False)]
    [switch]$ShowJson = $true                ## Dump raw JSON response per datasource
)

#region ----- Auth ----

function Get-InAgentVisa {
    param ([string]$CtPath, [string]$LocalUrl)

    ## Find the account directory (contains in_agent_authentication_token file)
    $storageRoot = 'C:\ProgramData\MXB\Backup Manager\storage'
    $acctDir = (Get-ChildItem $storageRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne 'certificates' } |
        ForEach-Object {
            $tf = Join-Path $_.FullName 'in_agent_authentication_token'
            if (Test-Path $tf) { [PSCustomObject]@{ Dir = $_.FullName; T = (Get-Item $tf).LastWriteTime } }
        } | Sort-Object T -Descending | Select-Object -First 1).Dir

    if (-not $acctDir) { throw "Cannot locate account directory under $storageRoot" }
    Write-Verbose "Account dir: $(Split-Path $acctDir -Leaf)"

    ## Get InAgent auth token — try ClientTool first, fall back to raw file
    $token = $null
    if (Test-Path $CtPath) {
        $ctOut = & $CtPath 'in-agent-authentication-token.get' '-config-path' $acctDir 2>&1
        if ($LASTEXITCODE -eq 0 -and $ctOut) { $token = $ctOut.Trim() }
    }
    if (-not $token) {
        $tf = Join-Path $acctDir 'in_agent_authentication_token'
        if (Test-Path $tf) { $token = (Get-Content $tf -Raw).Trim() }
    }
    if (-not $token) { throw "Could not retrieve InAgent auth token" }

    ## Login to local agent
    $loginBody = @{ jsonrpc='2.0'; id=1; method='InAgentAuthenticationTokenLogin'; params=@{ token=$token } } | ConvertTo-Json
    $loginResp = Invoke-RestMethod -Uri $LocalUrl -Method POST -ContentType 'application/json' -Body $loginBody

    if ($loginResp.error) { throw "Login failed: $($loginResp.error.message)" }

    ## Visa may be in result.result (double-nested) or result directly
    $visa = if ($loginResp.result -and $loginResp.result.PSObject.Properties['result']) { $loginResp.result.result } else { $loginResp.result }
    if (-not $visa) { throw "No visa returned from InAgentAuthenticationTokenLogin" }

    return $visa
}

function Invoke-LocalApi ([string]$Method, [hashtable]$Params, [string]$Visa, [string]$Url) {
    $body = @{ jsonrpc='2.0'; id=1; method=$Method; visa=$Visa; params=$Params } | ConvertTo-Json -Depth 10
    return Invoke-RestMethod -Uri $Url -Method POST -ContentType 'application/json' -Body $body
}

#endregion

#region ----- Main ----

## Authenticate
$visa = $null
try {
    $visa = Get-InAgentVisa -CtPath $ClientToolPath -LocalUrl $LocalUrl
    Write-Host "  Authenticated via InAgent token." -ForegroundColor Green
} catch {
    Write-Error "Authentication failed: $_"
    exit 1
}

## Query selections per datasource
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($ds in $DataSources) {
    $apiParams = @{ plugin = $ds }
    $reqBody   = @{ jsonrpc='2.0'; id=1; method='EnumerateBackupSelections'; visa=$visa; params=$apiParams } | ConvertTo-Json -Compress

    $resp = Invoke-LocalApi -Method 'EnumerateBackupSelections' -Params $apiParams -Visa $visa -Url $LocalUrl

    if ($ShowJson) {
        Write-Host "`n  [$ds] request:" -ForegroundColor DarkYellow
        Write-Host $reqBody -ForegroundColor DarkYellow
        Write-Host "  [$ds] response:" -ForegroundColor DarkCyan
        Write-Host ($resp | ConvertTo-Json -Depth 8 -Compress) -ForegroundColor DarkGray
    }

    if ($resp.error) {
        Write-Verbose "  $ds : API error $($resp.error.message)"
        continue
    }

    ## Result may be double-nested (result.result) or flat (result)
    $rows = if ($resp.result -and $resp.result.PSObject.Properties['result']) { $resp.result.result } else { $resp.result }

    if (-not $rows) {
        ## null = plugin not configured on this device — skip (no synthetic entries)
        continue
    }

    foreach ($row in @($rows)) {
        $results.Add([PSCustomObject]@{
            DataSource = $row.PluginId
            Type       = $row.Type
            Priority   = $row.Priority
            Path       = $row.Path
            Flags      = ($row.Flags -join ',')
        })
    }
}

if ($results.Count -eq 0) {
    Write-Warning "No selections found."
    exit 0
}

#endregion

#region ----- Output ----

$results | Format-Table DataSource, Type, Priority, Path, Flags -AutoSize

Write-Host "`n  Selection Summary" -ForegroundColor Cyan
Write-Host "  -----------------" -ForegroundColor Cyan

foreach ($ds in ($results | Select-Object -ExpandProperty DataSource -Unique)) {
    $rows     = $results | Where-Object { $_.DataSource -eq $ds }
    $includes = @($rows | Where-Object { $_.Type -eq 'Inclusive' })
    $excludes = @($rows | Where-Object { $_.Type -eq 'Exclusive' })
    $excPaths = @($excludes | Where-Object { $_.Path } | Select-Object -ExpandProperty Path)

    Write-Host "`n  $ds" -ForegroundColor Yellow

    foreach ($inc in $includes) {
        $p    = if ($inc.Path) { $inc.Path } else { "(entire $ds datasource)" }
        $flag = if ($inc.Flags -like '*CreatedByAccountProfile*') { ' [Profile]' } else { '' }
        $isReInclude = $excPaths | Where-Object { $inc.Path -and $inc.Path -like "$_*" -and $inc.Path -ne $_ }
        $label = if ($isReInclude) { "Re-includes:" } else { "Backs up:   " }
        Write-Host "    $label $p$flag" -ForegroundColor Green
    }
    foreach ($exc in $excludes) {
        $flag = if ($exc.Flags -like '*CreatedByAccountProfile*') { ' [Profile]' } else { '' }
        Write-Host "    Excludes:    $($exc.Path)$flag" -ForegroundColor Red
    }

    $incList = ($includes | ForEach-Object { if ($_.Path) { $_.Path } else { "(entire $ds)" } }) -join ', '
    $excList = $excPaths -join ', '
    $summary = "    Summary: Backing up [$incList]"
    if ($excList) { $summary += ", excluding [$excList]" }
    Write-Host $summary -ForegroundColor DarkCyan
}

#endregion
