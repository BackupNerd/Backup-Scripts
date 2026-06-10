<# ----- About: ----
    # N-able | Cove Data Protection — Set NinjaOne CoveInstallationID
    # Revision v2.0 - 2026-06-10
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email: eric.harless@n-able.com
    # Reddit  https://www.reddit.com/r/Nable/
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
# -----------------------------------------------------------#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>

<# ----- Compatibility: ----
    # For use with the Standalone edition of N-able | Cove Data Protection
    # Requires a NinjaOne M2M API application (Client Credentials flow)
    # Tested with PowerShell 5.1 and PowerShell 7+
    # Sample scripts are subject to change without notification
    # Some script elements may be developed, tested or documented using AI
# -----------------------------------------------------------#>

<# ----- Behavior: ----
    # Reads a Cove customers-with-packages CSV (produced by Add-CoveCustomers.ps1
    # or similar), matches each row to a NinjaOne organization by name, and
    # PATCHes the CoveInstallationID org custom field with the Cove Installation
    # Package GUID.
    #
    # This automates Step 3 of the official Cove/NinjaOne integration guide:
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/
    #   Content/external-cove-integrations/ninjaOne/NinjaOne.htm
# -----------------------------------------------------------#>
#
# ── WHAT THIS SCRIPT DOES (end-to-end flow) ─────────────────────────────────
#
#   1.  You supply a CSV that contains one row per Cove customer, with the
#       Cove "Installation Package" GUID pre-populated (the GUID from the
#       Add Devices wizard, e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).
#
#   2.  The script authenticates to the NinjaOne REST API using OAuth2
#       Client Credentials (machine-to-machine — no browser pop-up needed).
#
#   3.  It fetches the complete list of NinjaOne organizations from the API.
#
#   4.  For each CSV row it tries to find the matching NinjaOne org by name:
#         a) Exact (normalised) match on the Cove "CoveName" column
#         b) Fallback: exact (normalised) match on "LegalCompanyName"
#       Normalisation strips punctuation, spaces, and case so that
#       "Animal Care, LLC" matches "Animal Care LLC" etc.
#
#   5.  If a match is found, the script reads the org's existing custom fields
#       to check whether CoveInstallationID is already populated.  If it is,
#       the row is skipped (unless you pass -Force to overwrite).
#
#   6.  If the field is empty (or -Force is set), it PATCHes the field with
#       the GUID from the CSV.
#
#   7.  Every row — updated, skipped, not-matched, or errored — is written to
#       an output CSV so you have a complete audit trail.
#
# ── PREREQUISITES ────────────────────────────────────────────────────────────
#
#   A. NinjaOne custom field "CoveInstallationID" MUST already exist.
#      Create it manually ONCE before running this script:
#        NinjaOne Dashboard → Administration → Organizations
#                           → Organization Custom Fields → Add Custom Field
#        Label   : Cove Installation ID
#        Name    : CoveInstallationID          ← this is the API field name
#        Type    : Text
#        Permissions → Automations: Read Only | API: None | Technician: Editable
#      Reference: NinjaOne Integration Step 1:
#      https://documentation.n-able.com/covedataprotection/USERGUIDE/
#        documentation/Content/external-cove-integrations/ninjaOne/
#        NinjaOne-create-custom-fields.htm
#
#   B. A NinjaOne Machine-to-Machine (M2M) API application.
#      Create it once, then use its Client ID + Client Secret with this script:
#        NinjaOne Dashboard → Administration → Apps → API
#                           → Add → Machine to Machine
#        Grant Types : Client Credentials
#        Scopes      : monitoring  management
#      IMPORTANT: Copy the Client Secret immediately — NinjaOne only shows it
#      once at creation time.
#
# ── USAGE ────────────────────────────────────────────────────────────────────
#
#   # ALWAYS run with -WhatIf first to see what would change (no writes occur):
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 `
#       -CsvPath ".\output\mph-customers-with-packages.csv" -WhatIf
#
#   # Apply changes (you will be prompted for Client ID and Client Secret):
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 `
#       -CsvPath ".\output\mph-customers-with-packages.csv"
#
#   # Apply with credentials on the command line (useful for automation):
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 `
#       -CsvPath ".\output\cove-customers-with-packages.csv" `
#       -ClientId  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
#       -ClientSecret "your-secret-here"
#
#   # EU-hosted NinjaOne instance:
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 -CsvPath "..." -NinjaRegion "eu"
#
#   # Force-overwrite orgs that already have a value set:
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 -CsvPath "..." -Force
#
#   # Show verbose skip messages (AlreadySet rows):
#   .\Set-NinjaOneCoveInstallationID-v2.ps1 -CsvPath "..." -Verbose
#
# ── HOW THE TWO CSVs RELATE ─────────────────────────────────────────────────
#
#   You have two source files:
#
#   1) Cove packages CSV  (cove-customers-with-packages.csv)
#      Produced by Add-CoveCustomers.ps1 (or similar).
#      Key columns this script uses:
#        ExternalCode        — NinjaOne org name (e.g. "ST001 - Example Animal Clinic")
#                              THIS IS THE PRIMARY MATCH KEY.
#                              Cove stores this as its own "external code" field, but
#                              the value was set to match the NinjaOne org name exactly.
#        CoveName            — short internal code (e.g. "EXAMPLE")  — fallback match only
#        LegalCompanyName    — full legal name (e.g. "Example Animal Clinic") — fallback
#        InstallationPackage — Cove GUID to write into NinjaOne (e.g. xxxxxxxx-xxxx-...)
#
#   MATCH STRATEGY (3-pass, first hit wins):
#     Pass 1: ExternalCode (Cove) == NinjaOne org name       ← almost always succeeds
#     Pass 2: LegalCompanyName   == NinjaOne org name        ← fallback
#     Pass 3: CoveName           == NinjaOne org name        ← last resort
#   All comparisons are normalised (strip punctuation/spaces/case).
#
# ── OUTPUT CSV COLUMNS ───────────────────────────────────────────────────────
#
#   ExternalCode   — Cove ExternalCode column (= NinjaOne org name)
#   CoveName       — short internal code from the Cove CSV
#   LegalName      — full legal company name from the Cove CSV
#   PackageGuid    — the Cove Installation Package GUID
#   NinjaOrgId     — NinjaOne numeric org ID (blank if no match)
#   NinjaOrgName   — NinjaOne org name exactly as returned by the API
#   Result         — one of: Updated | AlreadySet | NoMatch | NoPackage | BadGuid | Error | WhatIf
#   Detail         — human-readable explanation of the result
#
# ── RESULT CODE MEANINGS ─────────────────────────────────────────────────────
#
#   Updated    — field was empty; GUID written successfully
#   AlreadySet — field already had a value; row skipped (use -Force to overwrite)
#   NoMatch    — ExternalCode (and fallbacks) did not match any NinjaOne org name
#                  → most likely cause: ExternalCode in Cove does not yet match
#                    the NinjaOne org name; compare output CSV to NinjaOne dashboard
#   NoPackage  — the InstallationPackage column was empty for this CSV row
#                  → re-generate the package in Cove and re-run
#   BadGuid    — the InstallationPackage value is present but not a valid GUID format
#                  → verify the value in Cove matches the pattern xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
#   Error      — the API PATCH call returned an error (see Detail for HTTP info)
#                  → check: CoveInstallationID field exists in NinjaOne and
#                    your M2M app has the "management" scope
#   WhatIf     — -WhatIf was active; shows what WOULD have happened, no change made
#
# ── OFFLINE TESTING WITHOUT A NINJAONE API KEY ───────────────────────────────
#
#   Use -WhatIf without credentials to preview CSV rows offline without contacting
#   NinjaOne at all.  Useful for validating the CSV before you have an API key.
#
#   In WhatIf (no credentials) mode:
#     • No OAuth2 token is requested (ClientId / ClientSecret not needed)
#     • Get-AllNinjaOrganizations is replaced by reading the mock CSV
#       Fake numeric org IDs are assigned sequentially (1, 2, 3, ...)
#     • Get-NinjaOrgCustomFields always returns an empty object
#       (simulates all fields being unset — every matched row will show
#        Result = "WhatIf" since -WhatIf is automatically forced in mock mode)
#     • Set-NinjaOrgCustomField is never called — no writes to any system
#
#   Usage:
#     .\.Set-NinjaOneCoveInstallationID-v2.ps1 `
#         -CsvPath ".\output\cove-customers-with-packages.csv" -WhatIf
#
#   The output CSV will contain WhatIf rows for every row in the input CSV.
#   This lets you validate the CSV contents before connecting to NinjaOne.
#
# ── NINJAONE API REFERENCES ──────────────────────────────────────────────────
#
#   Auth overview (OAuth2 intro, grant type guide):
#   https://app.ninjarmm.com/apidocs-beta/authorization/overview
#
#   Register a Machine-to-Machine (M2M) application:
#   https://app.ninjarmm.com/apidocs-beta/authorization/create-applications/machine-to-machine-apps
#
#   Client Credentials Flow (token request, request/response shape):
#   https://app.ninjarmm.com/apidocs-beta/authorization/flows/client-credentials-flow
#
#   POST /ws/oauth/token  — obtain Bearer token:
#   (See Client Credentials Flow link above — no separate page for the endpoint)
#
#   GET  /v2/organizations — list all orgs (cursor-paginated):
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/getOrganizations
#
#   GET  /v2/organization/{id}/custom-fields — read org custom fields:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/getNodeCustomFields_1
#
#   PATCH /v2/organization/{id}/custom-fields — write org custom fields:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/updateNodeAttributeValues_1
#
#   Custom field values — field types, name vs label, body format:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/articles/customFields/customFieldsValues/custom-fields-values
#
#   Cove/NinjaOne integration guide (manual steps this script automates):
#   https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/external-cove-integrations/ninjaOne/NinjaOne.htm
#
################################################################################

[CmdletBinding(SupportsShouldProcess)]
param (

    # ── Required ─────────────────────────────────────────────────────────────
    # Full or relative path to the Cove customers-with-packages CSV.
    # This is the file produced by Add-CoveCustomers.ps1 (or similar).
    #
    # REQUIRED columns (script validates these at startup and stops if missing):
    #   ExternalCode        — NinjaOne org name, e.g. "ST001 - Example Animal Clinic"
    #                         This is the PRIMARY match key.  In Cove it is stored as the
    #                         "External Code" field on the customer/partner record, and was
    #                         intentionally set to exactly match the NinjaOne org name.
    #   InstallationPackage — Cove package GUID, e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    #                         This is what gets written to the NinjaOne custom field.
    #
    # OPTIONAL columns (used as name-match fallbacks if ExternalCode doesn't hit):
    #   CoveName            — short internal code, e.g. "EXAMPLE"
    #   LegalCompanyName    — full legal name, e.g. "Example Animal Clinic"
    #
    # All other columns in the standard CSV (Row, PartnerId, ParentId, Level,
    # State, Status, etc.) are present but ignored.
    [Parameter(Mandatory, HelpMessage = 'Path to the Cove customers-with-packages CSV (e.g. .\output\cove-customers-with-packages.csv)')]
    [string]$CsvPath,

    # ── NinjaOne API credentials ──────────────────────────────────────────────
    # Client ID from your NinjaOne M2M API application.
    # If omitted, the script prompts interactively.
    [string]$ClientId,

    # Client Secret from your NinjaOne M2M API application.
    # If omitted, the script prompts interactively (input is masked with ****).
    # NEVER hard-code this in a shared or version-controlled script file.
    # Use the interactive prompt, a secrets vault, or a secure environment var.
    [string]$ClientSecret,

    # ── NinjaOne region ───────────────────────────────────────────────────────
    # Choose the data-center region that matches your NinjaOne account login URL:
    #   us  → app.ninjarmm.com      (North America — default)
    #   eu  → eu.ninjarmm.com       (Europe)
    #   oc  → oc.ninjarmm.com       (Oceania / ANZ)
    #   ca  → ca.ninjarmm.com       (Canada)
    # Using the wrong region causes 401 errors or connection timeouts.
    [ValidateSet("us","eu","oc","ca")]
    [string]$NinjaRegion = "us",

    # ── Custom field name ─────────────────────────────────────────────────────
    # The API/internal "Name" of the NinjaOne org custom field to write to.
    # This is the "Name" field set when you created it in the dashboard
    # (NOT the "Label" — Labels are for display only):
    #   Label  : Cove Installation ID   (shown in the NinjaOne UI)
    #   Name   : CoveInstallationID     ← what the API uses — must match exactly
    # Change this parameter only if you used a different name when creating
    # the custom field.
    [string]$CustomFieldName = "CoveInstallationID",

    # ── Force overwrite ───────────────────────────────────────────────────────
    # By default, orgs that already have a value in CoveInstallationID are
    # SKIPPED to prevent accidentally overwriting values that were manually
    # corrected in the NinjaOne dashboard.
    # Add -Force to overwrite existing values — for example, if you
    # regenerated all Cove installation packages and need to push new GUIDs.
    [switch]$Force,

    # ── Output path ───────────────────────────────────────────────────────────
    # Where to write the per-row result CSV for auditing.
    # Defaults to <workspace>\output\ninja-cove-sync-YYYYMMDD-HHmmss.csv
    # relative to the input CSV's parent folder.
    [string]$OutputCsv,

    # ── Interactive row selection ─────────────────────────────────────────────
    # Open a GUI grid view after loading the CSV so you can hand-pick which
    # customers to process.  Hold Ctrl or Shift to multi-select rows, then
    # click OK.  Only the selected rows will be processed; all others are
    # skipped (they will not appear in the output CSV).
    #
    # Requires Windows PowerShell or PowerShell 7+ on Windows (Out-GridView is
    # Windows-only).  Not compatible with -NonInteractive sessions.
    #
    # Pair with -WhatIf to preview which rows you selected before committing:
    #   .\Set-NinjaOneCoveInstallationID-v2.ps1 -CsvPath '...' -SelectFromGridView -WhatIf
    [switch]$SelectFromGridView = $true
)

# Change to the script's own directory so that relative paths (e.g. .\output\...)
# resolve correctly regardless of which directory the caller invoked from.
Set-Location -LiteralPath $PSScriptRoot

# Stop immediately on any unhandled error rather than silently continuing.
# This prevents partial runs that appear successful but silently dropped rows.
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

################################################################################
# SECTION 1 — Region / base URL setup
################################################################################
# NinjaOne hosts separate API instances per geographic region.
# The $BaseUrl is prefixed to EVERY API call in this script, so choosing the
# wrong region means all calls go to the wrong server.
#
# All NinjaOne v2 REST paths follow:   {BaseUrl}/v2/{resource}
# The OAuth2 token endpoint is at:    {BaseUrl}/ws/oauth/token
$RegionBaseUrl = @{
    us = "https://app.ninjarmm.com"   # North America (default)
    eu = "https://eu.ninjarmm.com"    # Europe
    oc = "https://oc.ninjarmm.com"    # Oceania / ANZ
    ca = "https://ca.ninjarmm.com"    # Canada
}
$BaseUrl = $RegionBaseUrl[$NinjaRegion]

################################################################################
# SECTION 2 — Output file path
################################################################################
# Build a timestamped output path unless the caller supplied one explicitly.
# We look for a sibling "output" folder one level up from the CSV's directory;
# if that doesn't exist, we fall back to writing next to the CSV itself.
if (-not $OutputCsv) {
    $Stamp     = (Get-Date -Format "yyyyMMdd-HHmmss")
    $OutputDir = Join-Path (Split-Path $CsvPath -Parent) "..\output"
    $OutputDir = (Resolve-Path $OutputDir -ErrorAction SilentlyContinue)?.Path
    if (-not $OutputDir) {
        # Fallback: same folder as the input CSV
        $OutputDir = Split-Path $CsvPath -Parent
    }
    $OutputCsv = Join-Path $OutputDir "ninja-cove-sync-$Stamp.csv"
}

################################################################################
# SECTION 3 — Credential acquisition  (skipped in mock mode)
################################################################################
# We need two values from the NinjaOne M2M application:
#   ClientId     — the application identifier (looks like a UUID/GUID)
#   ClientSecret — the secret shown ONCE at app creation time in NinjaOne
#
# Credentials can be:
#   • Passed as -ClientId / -ClientSecret parameters (automation / pipeline)
#   • Typed interactively when prompted (human-run sessions)
#
# In -WhatIf mode with no credentials, all NinjaOne API calls are skipped.
# The script previews what each CSV row would write (using ExternalCode as the
# org name) without contacting NinjaOne at all — useful for a quick sanity check
# of the CSV before you have an API key.
#
# Security guidance:
#   • Don't pass -ClientSecret on the command line in shared environments —
#     it may appear in shell history or process listings.
#   • Prefer the interactive prompt, a secrets vault, or an environment variable
#     that is set immediately before running the script and cleared after.

# $SkipApi = true when WhatIf is active and no credentials were supplied.
# In this state the script shows a preview without touching NinjaOne.
$SkipApi = $WhatIfPreference -and (-not $ClientId) -and (-not $ClientSecret)

if (-not $SkipApi) {
    if (-not $ClientId -or -not $ClientSecret) {
        Write-Host ""
        Write-Host "NinjaOne API credentials required." -ForegroundColor Cyan
        Write-Host "Create an M2M app at: $BaseUrl/app/settings/integrations/api" -ForegroundColor DarkCyan
        Write-Host "Required scopes: monitoring, management" -ForegroundColor DarkCyan
        Write-Host ""

        if (-not $ClientId) {
            $ClientId = (Read-Host "Enter NinjaOne Client ID").Trim()
        }
        if (-not $ClientSecret) {
            # -AsSecureString masks the typed characters (shows ****)
            $SecureSecret = Read-Host "Enter NinjaOne Client Secret" -AsSecureString
            # Convert SecureString → plain text for the OAuth2 form body.
            # This is necessary; Invoke-RestMethod cannot use a SecureString directly.
            $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureSecret)
            )
        }
    }
}

################################################################################
# FUNCTION: Get-NinjaToken
################################################################################
# PURPOSE:
#   Exchange Client ID + Client Secret for a short-lived Bearer access token
#   using the OAuth2 "Client Credentials" grant type.  No user interaction or
#   browser required — this is the M2M (machine-to-machine) flow.
#
# API CALL:
#   POST {BaseUrl}/ws/oauth/token
#   Content-Type: application/x-www-form-urlencoded
#   Body:
#     grant_type    = client_credentials
#     client_id     = <your app's client ID>
#     client_secret = <your app's client secret>
#     scope         = monitoring management
#
# SUCCESSFUL RESPONSE (HTTP 200):
#   {
#     "access_token": "eyJhbGciOiJSUzI1...",   ← use this in Authorization header
#     "expires_in":   3600,                     ← valid for 1 hour
#     "token_type":   "Bearer"
#   }
#
# HOW THE TOKEN IS USED:
#   Every subsequent GET / PATCH call includes:
#     Authorization: Bearer <access_token>
#   as an HTTP request header.
#
# SCOPE REQUIREMENTS:
#   monitoring  — needed to GET /v2/organizations and GET custom-fields (read)
#   management  — needed to PATCH /v2/organization/{id}/custom-fields (write)
#   If either scope is missing from the M2M app, calls requiring it return 403.
#
# COMMON FAILURE REASONS:
#   HTTP 400 — wrong grant_type value, or scope string is malformed
#   HTTP 401 — wrong client_id or client_secret
#   HTTP 403 — correct credentials but the required scope was not granted to the app
#   Connection error — wrong NinjaRegion (pointing at the wrong data center)
#
# REFERENCE:
#   https://app.ninjarmm.com/apidocs-beta/authorization/flows/client-credentials-flow
function Get-NinjaToken {

    $TokenUrl = "$BaseUrl/ws/oauth/token"

    # The body MUST be form-encoded (application/x-www-form-urlencoded), NOT JSON.
    # Invoke-RestMethod handles URL-encoding automatically when given a hashtable
    # plus -ContentType "application/x-www-form-urlencoded".
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "monitoring management"
        #              ↑ space-separated list of scopes — both are required:
        #                monitoring → read orgs, read custom fields
        #                management → write (PATCH) custom fields
    }

    try {
        $Response = Invoke-RestMethod -Uri $TokenUrl -Method POST -Body $Body `
            -ContentType "application/x-www-form-urlencoded"

        # $Response.access_token is the Bearer string we attach to all subsequent calls.
        # It is valid for ~1 hour — sufficient for even a large batch run.
        return $Response.access_token
    }
    catch {
        # Surface the HTTP status code to help diagnose the problem.
        $StatusCode = $_.Exception.Response?.StatusCode?.value__
        throw "NinjaOne OAuth token request failed (HTTP $StatusCode). " +
              "Check Client ID/Secret and that both scopes are granted. Error: $_"
    }
}

################################################################################
# FUNCTION: Get-AllNinjaOrganizations
################################################################################
# PURPOSE:
#   Retrieve every NinjaOne organization visible to the authenticated app,
#   handling pagination transparently so the caller gets a flat list.
#
# API CALL (per page):
#   GET {BaseUrl}/v2/organizations?pageSize=200[&after={lastId}]
#   Authorization: Bearer <token>
#
# PAGINATION MODEL:
#   NinjaOne v2 uses cursor-based pagination (not page-number based).
#   • Each request returns up to `pageSize` items as a JSON array.
#   • If the page is full (count == pageSize), there MAY be more records.
#   • To get the next page, pass the numeric `id` of the last item as `after=`.
#   • Stop when a page returns fewer items than pageSize (no more data).
#
# RESPONSE SHAPE (array of org objects):
#   [
#     { "id": 123, "name": "Acme Corp", "description": "...", "userData": {} },
#     { "id": 124, "name": "Beta LLC",  ...                                  },
#     ...
#   ]
#   Key fields we use:
#     id    (integer) — NinjaOne org identifier; used in PATCH URL path
#     name  (string)  — display name; matched against CSV columns
#
# COMMON FAILURE REASONS:
#   HTTP 401 — token is expired, or missing "monitoring" scope
#   HTTP 403 — M2M app does not have permission to list organizations
#
# REFERENCE:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/getOrganizations
function Get-AllNinjaOrganizations {
    param(
        # Bearer token returned by Get-NinjaToken
        [string]$Token
    )

    $Headers  = @{ Authorization = "Bearer $Token"; Accept = "application/json" }
    $All      = [System.Collections.Generic.List[object]]::new()
    $After    = $null   # cursor — null/empty on the very first request
    $PageSize = 200     # 200 is the max page size the NinjaOne v2 API supports

    do {
        # Build the URL, appending the cursor only after the first page.
        $Url = "$BaseUrl/v2/organizations?pageSize=$PageSize"
        if ($After) { $Url += "&after=$After" }

        # Invoke-RestMethod automatically deserialises the JSON array into
        # a PowerShell array of PSCustomObjects.
        $Page = Invoke-RestMethod -Uri $Url -Headers $Headers -Method GET

        if ($Page -and $Page.Count -gt 0) {
            # Accumulate this page's orgs into the master list
            $All.AddRange([object[]]$Page)

            # Advance the cursor to the id of the last org on this page.
            # The next request will start immediately after this id.
            $After = $Page[-1].id
        }
        else {
            # Empty page → no more data
            break
        }

        # Loop while the page was completely full.
        # A partial page (Count < PageSize) signals the last page of data.
    } while ($Page.Count -eq $PageSize)

    return $All   # Complete, flat list of all org objects
}

################################################################################
# FUNCTION: Get-NinjaOrgCustomFields
################################################################################
# PURPOSE:
#   Read the current custom field values for a single NinjaOne organization.
#   Used to check whether CoveInstallationID is already populated before
#   deciding whether to PATCH it.
#
# API CALL:
#   GET {BaseUrl}/v2/organization/{orgId}/custom-fields
#   Authorization: Bearer <token>
#
# RESPONSE SHAPE (object with one property per custom field):
#   {
#     "CoveInstallationID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
#     "AnotherField":       "some value",
#     ...
#   }
#
#   If a custom field has never been set, it may be absent from the object
#   or present with a null/empty value.  Both cases are treated as "not set".
#
# WHEN IS THIS CALLED:
#   Once per matched org (unless -Force is active, in which case this call
#   is skipped entirely and we always overwrite).
#
# COMMON FAILURE REASONS:
#   HTTP 401 — token expired
#   HTTP 404 — orgId does not exist (e.g. org was deleted between the list
#              call and this call — handled by returning $null)
#
# REFERENCE:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/getNodeCustomFields_1
function Get-NinjaOrgCustomFields {
    param(
        [string]$Token,
        [int]$OrgId       # numeric NinjaOne org identifier from the list endpoint
    )

    $Headers = @{ Authorization = "Bearer $Token"; Accept = "application/json" }
    $Url     = "$BaseUrl/v2/organization/$OrgId/custom-fields"

    try {
        # Returns a PSCustomObject.  Access a specific field like:
        #   $Result.CoveInstallationID
        return Invoke-RestMethod -Uri $Url -Headers $Headers -Method GET
    }
    catch {
        # Return $null on any error so the caller treats this org as
        # "field not set" rather than aborting the entire batch run.
        return $null
    }
}

################################################################################
# FUNCTION: Set-NinjaOrgCustomField
################################################################################
# PURPOSE:
#   Write (create or update) a single custom field value on a NinjaOne org.
#
# API CALL:
#   PATCH {BaseUrl}/v2/organization/{orgId}/custom-fields
#   Authorization: Bearer <token>
#   Content-Type:  application/json
#   Body:          { "CoveInstallationID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" }
#
# PATCH SEMANTICS:
#   PATCH is a partial update — only the fields listed in the JSON body are
#   changed.  All other custom fields on the org are left exactly as they are.
#   This is safe to call even if the org has many other custom fields configured.
#
# BODY FORMAT:
#   A JSON object where each key is the custom field "Name" (not "Label"):
#     { "CoveInstallationID": "<guid>" }
#   For TEXT type fields the value is a plain string (max 200 characters).
#   A Cove Installation Package GUID is 36 characters — well within the limit.
#
# CUSTOM FIELD NAME vs LABEL:
#   The JSON key MUST match the "Name" field, NOT the "Label":
#     Label (UI display):  "Cove Installation ID"
#     Name  (API key):     "CoveInstallationID"     ← use this in the body
#   These were set when the custom field was created in:
#     Administration → Organizations → Organization Custom Fields
#
# SUCCESSFUL RESPONSE:
#   HTTP 204 No Content — empty body.  That's expected and correct.
#   We pipe the result to Out-Null because there is nothing useful to read.
#
# COMMON FAILURE REASONS:
#   HTTP 400 Bad Request   — field "Name" doesn't exist in NinjaOne, or value
#                            exceeds the max length for the field type
#   HTTP 401 Unauthorized  — token expired or "management" scope not granted
#   HTTP 403 Forbidden     — M2M app does not have write permission on org fields
#   HTTP 404 Not Found     — orgId doesn't exist (org deleted between list + PATCH)
#
# REFERENCE:
#   https://app.ninjarmm.com/apidocs-beta/core-resources/operations/updateNodeAttributeValues_1
#   https://app.ninjarmm.com/apidocs-beta/core-resources/articles/customFields/customFieldsValues/custom-fields-values
function Set-NinjaOrgCustomField {
    param(
        [string]$Token,
        [int]$OrgId,
        [string]$FieldName,    # the "Name" value of the NinjaOne custom field (not the Label)
        [string]$FieldValue    # the value to write — the Cove Installation Package GUID
    )

    $Headers = @{
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/json"
        Accept         = "application/json"
    }
    $Url = "$BaseUrl/v2/organization/$OrgId/custom-fields"

    # Build the JSON body as a single key-value pair.
    # Using @{ $FieldName = $FieldValue } makes the key name dynamic so this
    # function works for any custom field, not just CoveInstallationID.
    # -Compress removes whitespace; result looks like:
    #   {"CoveInstallationID":"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"}
    $Body = @{ $FieldName = $FieldValue } | ConvertTo-Json -Compress

    # PATCH the field.  204 No Content is the success response — pipe to
    # Out-Null because Invoke-RestMethod would otherwise return an empty object.
    Invoke-RestMethod -Uri $Url -Headers $Headers -Method PATCH -Body $Body | Out-Null
}

################################################################################
# FUNCTION: Normalize-OrgName
################################################################################
# PURPOSE:
#   Strip punctuation, spaces, and case from a name string so that two names
#   that refer to the same organization can be compared reliably, even when
#   they differ in formatting.
#
# WHAT IT DOES:
#   Removes every character that is NOT a letter (a-z) or digit (0-9),
#   then lowercases the result.
#
# EXAMPLES:
#   "Acme Corp, LLC"         → "acmecorpllc"
#   "ACME CORP LLC"          → "acmecorpllc"   ← matches above
#   "Animal Care & Clinic"   → "animalcareclinic"
#   "Animal Care + Clinic"   → "animalcareclinic"  ← matches above
#   "All Dogs & Cats (Glenwood Springs)" → "alldogsandcatsglenwooodsprings"
#     ↑ Note: "&" is stripped, NOT converted to "and" — adjust the regex if needed
#
# WHAT IT DOES NOT HANDLE:
#   • Abbreviations:  "St." vs "Saint", "Mt." vs "Mount"
#   • Completely different names between Cove and NinjaOne
#   • Unicode / accented characters (treated as non-alphanumeric, so stripped)
#
# ROWS THAT STILL DO NOT MATCH after normalization will appear as "NoMatch"
# in the output CSV.  To fix them:
#   1. Note the CoveName and LegalCompanyName values in the NoMatch rows.
#   2. Find the corresponding org in the NinjaOne dashboard.
#   3. Edit the CoveName or LegalCompanyName in the source CSV to more closely
#      match the NinjaOne org name (punctuation/case/spacing don't matter).
#   4. Re-run the script — rows that were already updated will be skipped.
function Normalize-OrgName ([string]$Name) {
    # Replace every character that is NOT a-z or 0-9 with nothing, then lowercase.
    # The regex character class [^a-z0-9] matches anything outside a-z and 0-9.
    # PowerShell's -replace operator is case-insensitive by default, so uppercase
    # A-Z are also matched and removed before .ToLower() is called.
    return ($Name -replace '[^a-z0-9]', '').ToLower()
}

################################################################################
# MAIN — Script body begins here
################################################################################

# ── Step 1: Load and validate the input CSV ─────────────────────────────────
#
# File check: stop immediately if the path doesn't exist — nothing else can run.
if (-not (Test-Path $CsvPath)) {
    throw "CSV not found: $CsvPath"
}
$Rows = Import-Csv -Path $CsvPath
Write-Host "Loaded $($Rows.Count) rows from: $CsvPath" -ForegroundColor Cyan

# Header validation: confirm the two required columns exist.
# Import-Csv creates PSCustomObjects whose property names come from the header row.
# If a required column is missing, accessing $Row.ExternalCode would silently
# return $null for every row — the script would run but produce 100% NoPackage
# results with no obvious error.  Catching it here saves time.
#
# We test against the first row's property names rather than hardcoding a fixed
# column order, so the CSV columns can appear in any order.
$RequiredColumns = @('ExternalCode', 'InstallationPackage')
if ($Rows.Count -gt 0) {
    $ActualColumns = $Rows[0].PSObject.Properties.Name
    $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $ActualColumns }
    if ($MissingColumns) {
        throw "CSV is missing required column(s): $($MissingColumns -join ', ')."
    }
    Write-Host "  Required columns present: $($RequiredColumns -join ', ')" -ForegroundColor Green

    # Informational: note which optional fallback columns are present.
    $OptionalColumns = @('CoveName', 'LegalCompanyName')
    $PresentOptional = $OptionalColumns | Where-Object { $_ -in $ActualColumns }
    $MissingOptional = $OptionalColumns | Where-Object { $_ -notin $ActualColumns }
    if ($PresentOptional) {
        Write-Host "  Optional fallback columns present: $($PresentOptional -join ', ')" -ForegroundColor Green
    }
    if ($MissingOptional) {
        Write-Host "  Optional fallback columns absent (ExternalCode match only): $($MissingOptional -join ', ')" `
            -ForegroundColor Yellow
    }
}
else {
    Write-Warning "CSV loaded but contains no data rows. Nothing to process."
}

# ── Optional: interactive row selection via grid view ─────────────────────────
# When -SelectFromGridView is specified, show the loaded rows in a GUI grid so
# the user can hand-pick which customers to process.  The grid displays the four
# most useful columns for decision-making (ExternalCode, CoveName, LegalCompanyName,
# InstallationPackage).  All other columns from the CSV are preserved on the
# selected rows and available to the rest of the script.
#
# If the user closes the window without selecting anything, or clicks Cancel,
# the script exits cleanly rather than processing zero rows silently.
if ($SelectFromGridView) {
    if ($Rows.Count -eq 0) {
        Write-Warning "No rows to display — grid view skipped."
    }
    else {
        # Build a display-friendly view: show only the key columns in a logical order.
        # The Title string appears in the window title bar.
        $GridProps = @('ExternalCode','CoveName','LegalCompanyName','InstallationPackage') |
            Where-Object { $_ -in $Rows[0].PSObject.Properties.Name }

        $Selected = $Rows |
            Select-Object -Property $GridProps |
            Out-GridView -Title "Select customers to update CoveInstallationID  ($($Rows.Count) total) — Ctrl/Shift+click to multi-select, then OK" `
                         -PassThru

        if (-not $Selected -or $Selected.Count -eq 0) {
            Write-Host "No rows selected. Exiting." -ForegroundColor Yellow
            exit 0
        }

        # Map the selected display rows back to the original full rows by
        # matching on ExternalCode (the primary key).
        $SelectedCodes = $Selected | ForEach-Object { $_.ExternalCode }
        $Rows = $Rows | Where-Object { $_.ExternalCode -in $SelectedCodes }
        Write-Host "  $($Rows.Count) row(s) selected for processing." -ForegroundColor Cyan
    }
}

# GUID validation regex: standard UUID/GUID format xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
# Used below to reject malformed InstallationPackage values before sending to the API.
# A Cove Installation Package GUID always follows this 8-4-4-4-12 hex pattern.
# Example valid value:  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
$GuidRegex = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

# ── Step 2 / 3: Authenticate + fetch org list ───────────────────────────────
#
# WHATIF (no credentials):
#   -WhatIf was passed and no ClientId/ClientSecret were supplied or prompted.
#   All NinjaOne API calls are skipped.  Each CSV row is previewed using its
#   ExternalCode directly as the org name — no org-ID lookup is performed.
#   NinjaOrgId will be blank in the output CSV; use this mode to verify the
#   CSV contents before obtaining API credentials.
#
# LIVE / WHATIF with credentials:
#   Authenticates and fetches the full org list so matching can be validated
#   against real NinjaOne data.  -WhatIf still prevents any PATCH writes.

if ($SkipApi) {
    Write-Host ""
    Write-Host "  *** WHATIF (no credentials) — NinjaOne API calls skipped ***" -ForegroundColor Magenta
    Write-Host "  Previewing CSV rows only.  NinjaOrgId will be blank in output." -ForegroundColor Magenta
    Write-Host ""
    $Token     = $null
    $NinjaOrgs = @()
}
else {
    # ── Live mode (or WhatIf with credentials): authenticate then fetch orgs ──
    Write-Host "Authenticating with NinjaOne ($BaseUrl)..." -ForegroundColor Cyan
    $Token = Get-NinjaToken
    Write-Host "  Authentication OK." -ForegroundColor Green

    Write-Host "Fetching NinjaOne organizations..." -ForegroundColor Cyan
    $NinjaOrgs = Get-AllNinjaOrganizations -Token $Token
    Write-Host "  Found $($NinjaOrgs.Count) NinjaOne organizations." -ForegroundColor Green
}

# ── Step 4: Build in-memory lookup index ─────────────────────────────────────
# Converts the org list into a hashtable keyed by normalized org name.
# This gives O(1) (instant) lookups during the per-row loop instead of
# scanning the full list each time.
#
# Example entries after building the index:
#   "exampleanimalclinic"    → { id = 42,  name = "Example Animal Clinic" }
#   "acmecorp"               → { id = 99,  name = "Acme Corp" }
#
# COLLISION NOTE:
# If two NinjaOne orgs normalize to the same key (e.g. "Acme Corp" and
# "AcmeCorp"), the last one encountered wins.  This is rare in practice and
# harmless — both represent the same customer and would receive the same GUID.
$NinjaByNorm = @{}
foreach ($Org in $NinjaOrgs) {
    $Key = Normalize-OrgName $Org.name
    $NinjaByNorm[$Key] = $Org   # overwrites if duplicate key — last-write-wins
}

# ── Step 5: Process each CSV row ─────────────────────────────────────────────
# Initialize the results list and per-outcome counters used in the summary.
$Results      = [System.Collections.Generic.List[PSCustomObject]]::new()
$CountUpdated = 0   # PATCH succeeded — field written
$CountSkipped = 0   # AlreadySet or NoPackage — no write needed/possible
$CountNoMatch = 0   # No NinjaOne org found matching this CSV row
$CountError   = 0   # API call threw an exception
$CountWhatIf  = 0   # -WhatIf preview — no write made

foreach ($Row in $Rows) {

    # ── Extract and clean key columns from this CSV row ─────────────────────
    # .Trim() removes stray leading/trailing whitespace baked into CSV values.
    # We use the null-coalescing pattern (?? '') so the script doesn't crash on
    # CSVs that omit the optional CoveName / LegalCompanyName columns entirely.
    $ExternalCode = ($Row.ExternalCode          ?? '').Trim()   # PRIMARY match key
    $PackageGuid  = ($Row.InstallationPackage   ?? '').Trim()   # value to write
    $CoveName     = ($Row.CoveName              ?? '').Trim()   # fallback match
    $LegalName    = ($Row.LegalCompanyName      ?? '').Trim()   # fallback match

    # ── Guard: skip rows with no Installation Package GUID ───────────────────
    # If InstallationPackage is blank, the Cove customer was created but the
    # package was never generated (or the column was not populated in the export).
    # There is nothing to write to NinjaOne — log it and move on.
    # Resolution: go to the Cove Add Devices wizard, complete the package setup
    # for this customer, copy the GUID, update the CSV, and re-run.
    if (-not $PackageGuid) {
        $Results.Add([PSCustomObject]@{
            ExternalCode = $ExternalCode
            CoveName     = $CoveName
            LegalName    = $LegalName
            PackageGuid  = ''
            NinjaOrgId   = ''
            NinjaOrgName = ''
            Result       = 'NoPackage'
            Detail       = 'InstallationPackage column is empty — re-generate in Cove then re-run'
        })
        $CountSkipped++
        continue   # skip to the next CSV row
    }

    # ── Guard: validate GUID format ───────────────────────────────────────────
    # The Cove Installation Package GUID must match the standard UUID format:
    #   xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  (8-4-4-4-12 lowercase hex groups)
    # Example: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    #
    # Sending a malformed value to the NinjaOne API would cause a 400 Bad Request
    # error.  Catching it here gives a clearer error message and avoids an API call.
    # Note: the regex is case-insensitive (accepts uppercase hex) because Cove
    # sometimes formats GUIDs with mixed case.
    if ($PackageGuid -notmatch $GuidRegex) {
        Write-Warning "  BAD GUID format: '$ExternalCode' → '$PackageGuid'"
        $Results.Add([PSCustomObject]@{
            ExternalCode = $ExternalCode
            CoveName     = $CoveName
            LegalName    = $LegalName
            PackageGuid  = $PackageGuid
            NinjaOrgId   = ''
            NinjaOrgName = ''
            Result       = 'BadGuid'
            Detail       = "InstallationPackage '$PackageGuid' is not a valid GUID — " +
                           'expected format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        })
        $CountError++
        continue   # skip to the next CSV row
    }

    # ── Name matching: 3-pass normalised comparison ───────────────────────────
    # We normalize both sides (strip punctuation/spaces/case) before comparing.
    # See the Normalize-OrgName function for what normalization does.
    #
    # Pass 1 — ExternalCode  (PRIMARY — almost always succeeds)
    #   The Cove ExternalCode field was intentionally populated with the NinjaOne
    #   org name (e.g. "ST001 - Example Animal Clinic").  After normalization
    #   this becomes "st001exampleanimalclinic" on both sides — an exact hit.
    #   If this misses, the ExternalCode in Cove was set to something different
    #   from the NinjaOne org name.
    #
    # Pass 2 — LegalCompanyName  (fallback)
    #   The full legal name from Cove.  May match if the NinjaOne org name happens
    #   to match the legal name (some orgs omit the state-code prefix).
    #
    # Pass 3 — CoveName  (last resort)
    #   The short internal code (e.g. "ACRES").  Unlikely to match a NinjaOne org
    #   name directly, but included as a safety net.
    #
    # If all three passes miss → "NoMatch" row logged in output CSV.
    $MatchedOrg = $null
    $MatchedBy  = ''

    $NormExternal = Normalize-OrgName $ExternalCode
    $NormLegal    = Normalize-OrgName $LegalName
    $NormCove     = Normalize-OrgName $CoveName

    if ($NormExternal -and $NinjaByNorm.ContainsKey($NormExternal)) {
        # Pass 1: ExternalCode matched a NinjaOne org name — expected primary path
        $MatchedOrg = $NinjaByNorm[$NormExternal]
        $MatchedBy  = 'ExternalCode'
    }
    elseif ($NormLegal -and $NinjaByNorm.ContainsKey($NormLegal)) {
        # Pass 2: LegalCompanyName matched — ExternalCode was absent or different
        $MatchedOrg = $NinjaByNorm[$NormLegal]
        $MatchedBy  = 'LegalCompanyName'
    }
    elseif ($NormCove -and $NinjaByNorm.ContainsKey($NormCove)) {
        # Pass 3: CoveName matched — both primary and secondary failed
        $MatchedOrg = $NinjaByNorm[$NormCove]
        $MatchedBy  = 'CoveName'
    }

    if (-not $MatchedOrg) {
        if ($SkipApi) {
            # No org list available — treat ExternalCode as the org name and
            # show a WhatIf preview row.  Org ID is unknown without credentials.
            $OrgId   = ''
            $OrgName = $ExternalCode
        }
        else {
            # All three passes failed against a real NinjaOne org list.
            # Most likely cause: the ExternalCode in Cove does not match the NinjaOne
            # org name.  Check the dashboard and compare with the ExternalCode column.
            Write-Warning "  NO MATCH: ExternalCode='$ExternalCode' / CoveName='$CoveName' / LegalName='$LegalName'"
            $Results.Add([PSCustomObject]@{
                ExternalCode = $ExternalCode
                CoveName     = $CoveName
                LegalName    = $LegalName
                PackageGuid  = $PackageGuid
                NinjaOrgId   = ''
                NinjaOrgName = ''
                Result       = 'NoMatch'
                Detail       = "ExternalCode '$ExternalCode' did not match any NinjaOne org name " +
                               '(and LegalCompanyName / CoveName fallbacks also failed) — ' +
                               'verify that ExternalCode in Cove matches the NinjaOne org name exactly'
            })
            $CountNoMatch++
            continue   # skip to the next CSV row
        }
    }
    else {
        # Unpack the matched org's id (integer) and display name (string).
        $OrgId   = $MatchedOrg.id
        $OrgName = $MatchedOrg.name
    }

    # ── Guard: skip if CoveInstallationID is already set (unless -Force) ──────
    # Before writing, read the org's current custom field values.
    # If CoveInstallationID already has a non-empty value, skip this org.
    # This prevents accidentally overwriting a value that was manually corrected
    # in the NinjaOne dashboard (e.g. after a package was regenerated individually).
    #
    # Use -Force to bypass this check and always write the CSV value.
    # Use -Verbose to see a message for each skipped "AlreadySet" org.
    if (-not $Force) {
        $CurrentFields = Get-NinjaOrgCustomFields -Token $Token -OrgId $OrgId

        # Access the field by its "Name" using PowerShell dot notation.
        # If $CurrentFields is $null (API error, or mock mode) OR the property
        # doesn't exist on the returned object, $ExistingValue is $null — treat
        # as unset and proceed to write.
        $ExistingValue = if ($CurrentFields) { $CurrentFields.$CustomFieldName } else { $null }

        if ($ExistingValue -and $ExistingValue.Trim()) {
            # Field is populated — skip and log the existing value for comparison.
            Write-Verbose "  SKIP (already set): '$OrgName'  existing=$ExistingValue"
            $Results.Add([PSCustomObject]@{
                ExternalCode = $ExternalCode
                CoveName     = $CoveName
                LegalName    = $LegalName
                PackageGuid  = $PackageGuid
                NinjaOrgId   = $OrgId
                NinjaOrgName = $OrgName
                Result       = 'AlreadySet'
                Detail       = "Existing value: $ExistingValue  |  Add -Force to overwrite"
            })
            $CountSkipped++
            continue   # skip to the next CSV row
        }
    }

    # ── Apply (or preview with -WhatIf) ──────────────────────────────────────
    # $WhatIfPreference is automatically $true when the caller passes -WhatIf.
    # [CmdletBinding(SupportsShouldProcess)] in the param block enables this.
    # -WhatIf runs the full matching and guard logic but skips all API writes,
    # making it safe to use for a complete dry-run on a real NinjaOne tenant.
    if ($WhatIfPreference) {
        # -WhatIf mode: show what WOULD happen without making any API call.
        # The output CSV row has Result = "WhatIf" so you can count what
        # would be changed before committing.
        Write-Host "  WHATIF: Would set '$OrgName' (id=$OrgId)  $CustomFieldName = $PackageGuid  [matched by $MatchedBy]" `
            -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{
            ExternalCode = $ExternalCode
            CoveName     = $CoveName
            LegalName    = $LegalName
            PackageGuid  = $PackageGuid
            NinjaOrgId   = $OrgId
            NinjaOrgName = $OrgName
            Result       = 'WhatIf'
            Detail       = "Matched by $MatchedBy — remove -WhatIf to apply"
        })
        $CountWhatIf++
    }
    else {
        # Live mode: send the PATCH request to NinjaOne.
        # On success: HTTP 204 No Content — the field is now set.
        # On failure: catch block logs the error and continues to the next row.
        try {
            Set-NinjaOrgCustomField -Token $Token -OrgId $OrgId `
                -FieldName $CustomFieldName -FieldValue $PackageGuid

            # Success — CoveInstallationID is now set on this NinjaOne org.
            # NinjaOne's deployment automation policy will read this field to
            # silently install and configure the Cove Backup Manager on all
            # devices in the org (Step 4–6 of the Cove/NinjaOne integration guide).
            Write-Host "  SET: '$OrgName' (id=$OrgId)  ←  $PackageGuid  [matched by $MatchedBy]" `
                -ForegroundColor Green
            $Results.Add([PSCustomObject]@{
                ExternalCode = $ExternalCode
                CoveName     = $CoveName
                LegalName    = $LegalName
                PackageGuid  = $PackageGuid
                NinjaOrgId   = $OrgId
                NinjaOrgName = $OrgName
                Result       = 'Updated'
                Detail       = "Matched by $MatchedBy"
            })
            $CountUpdated++
        }
        catch {
            # PATCH failed — log the full error message in the Detail column.
            # The script continues to the next row rather than aborting.
            # Common causes:
            #   "400 Bad Request"   — CoveInstallationID field doesn't exist in
            #                         NinjaOne, or the GUID string is malformed
            #   "401 Unauthorized"  — token expired mid-run (>1 hour elapsed)
            #   "403 Forbidden"     — M2M app is missing the "management" scope
            #   "404 Not Found"     — org was deleted in NinjaOne between the list
            #                         call and this PATCH (extremely rare)
            Write-Warning "  ERROR updating '$OrgName' (id=$OrgId): $_"
            $Results.Add([PSCustomObject]@{
                ExternalCode = $ExternalCode
                CoveName     = $CoveName
                LegalName    = $LegalName
                PackageGuid  = $PackageGuid
                NinjaOrgId   = $OrgId
                NinjaOrgName = $OrgName
                Result       = 'Error'
                Detail       = $_.ToString()
            })
            $CountError++
        }
    }
}

################################################################################
# SECTION — Summary report
################################################################################
# Print a human-readable totals block to the console.
# The full per-row output is in the CSV written below.
Write-Host ""
Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan
Write-Host " NinjaOne CoveInstallationID Sync Summary" -ForegroundColor Cyan
Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "  CSV rows processed   : $($Rows.Count)"

if ($WhatIfPreference) {
    Write-Host "  Would update         : $CountWhatIf" -ForegroundColor Yellow
    Write-Host "  (No changes made — remove -WhatIf to apply)" -ForegroundColor Yellow
}
else {
    Write-Host "  Updated              : $CountUpdated" `
        -ForegroundColor $(if ($CountUpdated -gt 0) { "Green" } else { "Gray" })
}

Write-Host "  Skipped (no change)  : $CountSkipped"
Write-Host "  No NinjaOne match    : $CountNoMatch" `
    -ForegroundColor $(if ($CountNoMatch -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Errors               : $CountError" `
    -ForegroundColor $(if ($CountError -gt 0) { "Red" } else { "Green" })
Write-Host ""

# ── Guidance when there are unmatched rows ────────────────────────────────────
if ($CountNoMatch -gt 0) {
    Write-Host "  TIP — Resolving NoMatch rows:" -ForegroundColor Yellow
    Write-Host "    1. Open the output CSV and filter Result = 'NoMatch'" -ForegroundColor Yellow
    Write-Host "    2. Look at the ExternalCode column for each NoMatch row" -ForegroundColor Yellow
    Write-Host "    3. In NinjaOne, find the org and note its exact org name" -ForegroundColor Yellow
    Write-Host "    4. In the Cove Management Console, update that customer's" -ForegroundColor Yellow
    Write-Host "       External Code field to match the NinjaOne org name exactly" -ForegroundColor Yellow
    Write-Host "       e.g. 'ST001 - Example Animal Clinic'" -ForegroundColor Yellow
    Write-Host "    5. Re-export the CSV and re-run — already-updated orgs are skipped" -ForegroundColor Yellow
    Write-Host "    NOTE: Punctuation, case, and spacing do NOT matter (normalization handles them)" -ForegroundColor Yellow
    Write-Host ""
}

################################################################################
# SECTION — Write output CSV
################################################################################
# Write every processed row to the output CSV, regardless of outcome.
# This gives you a complete audit trail showing exactly what was changed,
# what was skipped, and what failed — useful for review before running
# Steps 4–6 (generating and deploying the NinjaOne policy).
$Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Results written to: $OutputCsv" -ForegroundColor Cyan
