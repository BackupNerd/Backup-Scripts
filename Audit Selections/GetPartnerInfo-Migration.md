# Migrating from `GetPartnerInfo` to `GetPartnerTree` + `GetPartnerInfoById`

## Why migrate

N-able has deprecated `GetPartnerInfo(name=...)` because partner names are no longer unique. The deprecated method can silently select an unintended older partner or return `Not Found` when another matching partner is visible in the caller's hierarchy.

The replacement separates the operation into two steps:

1. **Find candidate IDs** with `GetPartnerTree` and `partnerFilter.NamePattern`.
2. **Retrieve authoritative details** with `GetPartnerInfoById` after the user or script chooses an ID.

This makes duplicate-name handling explicit and prevents silently operating on the wrong partner.

## Behavior differences

| Behavior | Deprecated `GetPartnerInfo` | Replacement |
|---|---|---|
| Matching | Exact, case-sensitive name | Case-insensitive substring search |
| Duplicate names | May silently select one | Returns candidates for selection |
| Full partner details | Returned directly | Retrieved by ID in a second call |
| Restricted levels | Script-side rejection | Filter `Root`, `Sub-root`, and `Distributor` candidates |
| Scope | Global name lookup behavior | Descendants of an explicit `partnerId` scope |

## Recommended workflow

1. Login and capture `Login.result.result.PartnerId` as `$Script:AuthPartnerId`.
2. Call `GetPartnerInfoById` for the authenticated partner.
3. If its level is not restricted, use that partner directly when no name was supplied.
4. If it is `Root`, `Sub-root`, or `Distributor`, prompt for a customer/partner name.
5. Call `GetPartnerTree` using the authenticated partner ID as the search scope.
6. Recursively walk `result.result` and keep nodes whose own `Info.Name` matches the entered text.
7. Filter restricted-level matches.
8. Automatically use a single match; use `Out-GridView` for multiple matches.
9. Call `GetPartnerInfoById` for the selected ID.

## Important API details

- `NamePattern` is a bare substring. Pass `Henry Schein`, not `*Henry Schein*` or `%Henry Schein%`.
- Matching is case-insensitive.
- `GetPartnerTree` responses are double-wrapped: `response.result.result`.
- The root/self node is not matched against its own `NamePattern`; only descendants are searched.
- `childrenLimit` does not guarantee an uncapped match count. Treat a round count such as 50 as “at least 50” and ask the user to refine the search.
- `fields` is an integer array such as `@(0,1,3,4)`.
- Never log or persist visa values or passwords.
- Guard recursive traversal against null children and cycles.

## Sample functions

### Get partner details by ID

```powershell
Function Get-CovePartnerInfoById {
    param(
        [Parameter(Mandatory=$True)]
        [int]$PartnerId
    )

    $data = @{
        jsonrpc = '2.0'
        id      = '2'
        visa    = $Script:visa
        method  = 'GetPartnerInfoById'
        params  = @{ partnerId = [int]$PartnerId }
    }

    $response = Invoke-RestMethod -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -Depth 5) `
        -Uri 'https://api.backup.management/jsonapi'

    if ($response.error) {
        throw $response.error.message
    }

    $info = $response.result.result
    return [PSCustomObject]@{
        Uid   = [string]$info.Uid
        Id    = [int]$info.Id
        Name  = [string]$info.Name
        Level = [string]$info.Level
    }
}
```

### Find a partner by name with tree search

This is the direct replacement for a deprecated `GetPartnerInfo(name=...)` call. The `Id` and `ParentId` properties remain integers for later API calls.

```powershell
Function Get-CovePartnerByName {
    param(
        [Parameter(Mandatory=$True)]
        [string]$PartnerName,

        [Parameter(Mandatory=$False)]
        [int]$PartnerId = $Script:AuthPartnerId
    )

    $data = @{
        jsonrpc = '2.0'
        id      = '2'
        visa    = $Script:visa
        method  = 'GetPartnerTree'
        params  = @{
            partnerId     = [int]$PartnerId
            fields        = @(0,1,3,4)
            childrenLimit = 5000
            partnerFilter = @{ NamePattern = [string]$PartnerName }
        }
    }

    $treeResponse = Invoke-RestMethod -Method POST `
        -ContentType 'application/json' `
        -Body (ConvertTo-Json $data -Depth 6) `
        -Uri 'https://api.backup.management/jsonapi'

    if ($treeResponse.error) {
        throw $treeResponse.error.message
    }

    $treeRoot = $treeResponse.result.result
    $restrictedLevels = @('Root','Sub-root','Distributor')
    $partnerMatches = [System.Collections.Generic.List[object]]::new()
    $infoById = @{}

    Function Add-PartnerTreeMatches {
        param($Node, [string]$Needle)

        if ($Node.Info -and $Node.Info.Id) {
            $infoById[[int]$Node.Info.Id] = $Node.Info
        }

        if ($Node.Info -and $Node.Info.Name -and ($Node.Info.Name -like "*$Needle*")) {
            if ($Node.Info.Level -notin $restrictedLevels) {
                $partnerMatches.Add([PSCustomObject]@{
                    Id       = [int]$Node.Info.Id
                    Name     = [string]$Node.Info.Name
                    Level    = [string]$Node.Info.Level
                    ParentId = [int]$Node.Info.ParentId
                    Path     = ''
                }) | Out-Null
            }
        }

        foreach ($child in $Node.Children) {
            Add-PartnerTreeMatches -Node $child -Needle $Needle
        }
    }

    if ($treeRoot) {
        Add-PartnerTreeMatches -Node $treeRoot -Needle $PartnerName
    }

    foreach ($match in $partnerMatches) {
        $pathParts = [System.Collections.Generic.List[string]]::new()
        $node = $infoById[[int]$match.Id]
        $visited = [System.Collections.Generic.HashSet[int]]::new()

        while ($node -and $visited.Add([int]$node.Id)) {
            $pathParts.Insert(0, [string]$node.Name)
            $node = $infoById[[int]$node.ParentId]
        }

        $match.Path = $pathParts -join ' > '
    }

    if ($partnerMatches.Count -eq 0) {
        throw "No partner found matching '$PartnerName'."
    }

    if ($partnerMatches.Count -gt 1) {
        # Select-Object controls the display order. IDs remain numeric in the source objects.
        $selected = $partnerMatches |
            Select-Object Id, Name, Level, ParentId, Path |
            Out-GridView -Title "Select Partner matching '$PartnerName'" -OutputMode Single

        if (-not $selected) {
            throw 'No partner was selected.'
        }
    } else {
        $selected = $partnerMatches[0]
    }

    return Get-CovePartnerInfoById -PartnerId ([int]$selected.Id)
}
```

### Combined resolver for a simple call site

Use this when a script should either use the authenticated partner or prompt/search when that partner is restricted.

```powershell
Function Get-CovePartnerInfo {
    param(
        [string]$PartnerName,
        [int]$SearchPartnerId = $Script:AuthPartnerId
    )

    $restrictedLevels = @('Root','Sub-root','Distributor')

    if (-not $PartnerName) {
        $self = Get-CovePartnerInfoById -PartnerId $SearchPartnerId
        if ($self.Level -notin $restrictedLevels) {
            return $self
        }

        do {
            $PartnerName = Read-Host 'Enter Customer/Partner display name (or partial name)'
        } while ([string]::IsNullOrWhiteSpace($PartnerName))
    }

    return Get-CovePartnerByName `
        -PartnerName $PartnerName `
        -PartnerId $SearchPartnerId
}
```

## Login integration

Capture the authenticated partner ID during login:

```powershell
if ($response.visa) {
    $Script:visa          = $response.visa
    $Script:AuthPartnerId = [int]$response.result.result.PartnerId
}
```

Then resolve the target partner:

```powershell
$target = Get-CovePartnerInfo -PartnerName $PartnerName
$Script:PartnerId = [int]$target.Id
$Script:PartnerName = $target.Name
```

Use `$Script:PartnerId` for subsequent `EnumerateAccountStatistics`, `EnumeratePartners`, or other partner-scoped calls.

## Migration checklist

- [ ] Capture `AuthPartnerId` from the Login response.
- [ ] Remove the `GetPartnerInfo` JSON-RPC call.
- [ ] Add `GetPartnerInfoById` for ID-based details.
- [ ] Add `GetPartnerTree` for name-to-ID searches.
- [ ] Unwrap `response.result.result`.
- [ ] Filter `Root`, `Sub-root`, and `Distributor` matches.
- [ ] Handle duplicate matches with `Out-GridView` or another explicit selector.
- [ ] Cast the selected ID to `[int]` before later API calls.
- [ ] Preserve the selected partner name and ID for report output.
- [ ] Test authentication and lookup under both restricted and non-restricted accounts.

## Reference implementation

The validated project reference is:

`C:\Scripts\0-Script Master\CDP.Github Get Device Installations\GetAllDeviceInstallations.v10.ps1`

Additional migration notes:

`C:\Scripts\0-Script Master\Cove-MCP-Server\docs\getpartnerinfo-migration.md`
