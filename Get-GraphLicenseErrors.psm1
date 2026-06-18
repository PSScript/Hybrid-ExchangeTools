#Requires -Version 5.1
<#
.SYNOPSIS
    License Assignment Error Discovery via Microsoft Graph REST API
.DESCRIPTION
    Queries user licenseAssignmentStates to find assignment failures.
    Requires: User.Read.All (AuditLog.Read.All optional for richer context)
.NOTES
    Token-based, no Microsoft.Graph module dependency
#>

function Get-GraphLicenseErrors {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [string]$UserPrincipalName,

        [Parameter()]
        [switch]$ErrorsOnly,

        [Parameter()]
        [int]$Top = 999
    )

    begin {
        $headers = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        try {
            if ($UserPrincipalName) {
                $uri = "$baseUri/users/$UserPrincipalName`?`$select=id,userPrincipalName,displayName,licenseAssignmentStates"
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                $users = @($response)
            }
            else {
                $uri = "$baseUri/users?`$select=id,userPrincipalName,displayName,licenseAssignmentStates&`$top=$Top"
                $users = [System.Collections.Generic.List[PSObject]]::new()

                do {
                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                    if ($response.value) { 
                        foreach ($u in $response.value) { $users.Add($u) }
                    }
                    $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
                } while ($uri)
            }

            foreach ($user in $users) {
                if (-not $user.licenseAssignmentStates) { continue }

                foreach ($state in $user.licenseAssignmentStates) {
                    $hasError = -not [string]::IsNullOrEmpty($state.error) -and $state.error -ne 'None'

                    if ($ErrorsOnly -and -not $hasError) { continue }

                    $results.Add([PSCustomObject]@{
                        UserPrincipalName    = $user.userPrincipalName
                        DisplayName          = $user.displayName
                        SkuId                = $state.skuId
                        AssignedByGroup      = $state.assignedByGroup
                        State                = $state.state
                        Error                = if ($hasError) { $state.error } else { $null }
                        LastUpdatedDateTime  = $state.lastUpdatedDateTime
                        DisabledPlans        = ($state.disabledPlans -join ';')
                    })
                }
            }
        }
        catch {
            Write-Error "Graph API call failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                Write-Error "Detail: $($errDetail.error.message)"
            }
        }
    }

    end {
        return $results
    }
}

function Get-GraphSkuFriendlyName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [string]$SkuId
    )

    $headers = @{
        Authorization  = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Headers $headers -Method Get
        $skuMap = @{}
        foreach ($sku in $response.value) {
            $skuMap[$sku.skuId] = $sku.skuPartNumber
        }

        if ($SkuId) {
            return $skuMap[$SkuId]
        }
        return $skuMap
    }
    catch {
        Write-Error "Failed to retrieve SKUs: $($_.Exception.Message)"
        return $null
    }
}

#region Provisioning Errors (Duplicate Proxy, UPN Conflicts, Sync Issues)

function Get-GraphProvisioningErrors {
    <#
    .SYNOPSIS
        Finds users with onPremisesProvisioningErrors (duplicate proxy, UPN conflicts, etc.)
    .DESCRIPTION
        Classic hybrid sync pain: ProxyAddressConflict, AttributeValueMustBeUnique, etc.
        These blocked license assignment in the old days and still cause grief.
    .EXAMPLE
        $syncErrors = Get-GraphProvisioningErrors -AccessToken $token
        $syncErrors | Where-Object Category -eq 'PropertyConflict' | Format-Table
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [string]$UserPrincipalName,

        [Parameter()]
        [switch]$ErrorsOnly,

        [Parameter()]
        [int]$Top = 999
    )

    begin {
        $headers = @{
            Authorization      = "Bearer $AccessToken"
            'Content-Type'     = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        try {
            $selectProps = 'id,userPrincipalName,displayName,mail,proxyAddresses,onPremisesProvisioningErrors,onPremisesSyncEnabled,onPremisesLastSyncDateTime'

            if ($UserPrincipalName) {
                $uri = "$baseUri/users/$UserPrincipalName`?`$select=$selectProps"
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                $users = @($response)
            }
            else {
                # Fetch all users, filter client-side for errors (Graph doesn't support onPremisesProvisioningErrors filter well)
                $uri = "$baseUri/users?`$select=$selectProps&`$top=$Top"
                $users = [System.Collections.Generic.List[PSObject]]::new()

                do {
                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                    if ($response.value) { 
                        foreach ($u in $response.value) { $users.Add($u) }
                    }
                    $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
                } while ($uri)
            }

            foreach ($user in $users) {
                if ($ErrorsOnly -and (-not $user.onPremisesProvisioningErrors -or $user.onPremisesProvisioningErrors.Count -eq 0)) {
                    continue
                }

                if ($user.onPremisesProvisioningErrors -and $user.onPremisesProvisioningErrors.Count -gt 0) {
                    foreach ($err in $user.onPremisesProvisioningErrors) {
                        $results.Add([PSCustomObject]@{
                            UserPrincipalName     = $user.userPrincipalName
                            DisplayName           = $user.displayName
                            Mail                  = $user.mail
                            Category              = $err.category
                            PropertyCausingError  = $err.propertyCausingError
                            Value                 = $err.value
                            OccurredDateTime      = $err.occurredDateTime
                            OnPremisesSyncEnabled = $user.onPremisesSyncEnabled
                            LastSyncDateTime      = $user.onPremisesLastSyncDateTime
                            ProxyAddresses        = ($user.proxyAddresses -join '; ')
                        })
                    }
                }
                elseif (-not $ErrorsOnly) {
                    # Include clean users when not filtering
                    $results.Add([PSCustomObject]@{
                        UserPrincipalName     = $user.userPrincipalName
                        DisplayName           = $user.displayName
                        Mail                  = $user.mail
                        Category              = $null
                        PropertyCausingError  = $null
                        Value                 = $null
                        OccurredDateTime      = $null
                        OnPremisesSyncEnabled = $user.onPremisesSyncEnabled
                        LastSyncDateTime      = $user.onPremisesLastSyncDateTime
                        ProxyAddresses        = ($user.proxyAddresses -join '; ')
                    })
                }
            }
        }
        catch {
            Write-Error "Graph API call failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                Write-Error "Detail: $($errDetail.error.message)"
            }
        }
    }

    end {
        return $results
    }
}

function Find-DuplicateProxyAddresses {
    <#
    .SYNOPSIS
        Scans all users for duplicate proxyAddresses that could cause sync/license issues
    .DESCRIPTION
        Finds SMTP/smtp collisions across users - the classic Exchange hybrid nightmare.
        Returns grouped duplicates with all affected users.
    .EXAMPLE
        $dupes = Find-DuplicateProxyAddresses -AccessToken $token
        $dupes | Where-Object Count -gt 1 | Format-Table ProxyAddress, Count, Users
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [int]$Top = 999
    )

    begin {
        $headers = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $proxyMap = @{}
    }

    process {
        try {
            $uri = "$baseUri/users?`$select=userPrincipalName,displayName,proxyAddresses&`$top=$Top"
            $users = [System.Collections.Generic.List[PSObject]]::new()

            do {
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                if ($response.value) { 
                    foreach ($u in $response.value) { $users.Add($u) }
                }
                $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
            } while ($uri)

            foreach ($user in $users) {
                if (-not $user.proxyAddresses) { continue }

                foreach ($proxy in $user.proxyAddresses) {
                    # Normalize: lowercase for comparison
                    $normalized = $proxy.ToLower()

                    if (-not $proxyMap.ContainsKey($normalized)) {
                        $proxyMap[$normalized] = [System.Collections.Generic.List[string]]::new()
                    }
                    $proxyMap[$normalized].Add($user.userPrincipalName)
                }
            }
        }
        catch {
            Write-Error "Graph API call failed: $($_.Exception.Message)"
        }
    }

    end {
        $results = foreach ($kvp in $proxyMap.GetEnumerator()) {
            [PSCustomObject]@{
                ProxyAddress = $kvp.Key
                Count        = $kvp.Value.Count
                IsDuplicate  = $kvp.Value.Count -gt 1
                Users        = ($kvp.Value -join '; ')
                Type         = if ($kvp.Key -cmatch '^SMTP:') { 'Primary' } elseif ($kvp.Key -match '^smtp:') { 'Alias' } else { 'Other' }
            }
        }

        return $results | Sort-Object -Property Count -Descending
    }
}

function Get-GraphUserConflicts {
    <#
    .SYNOPSIS
        Combined view: license errors + provisioning errors + duplicate proxy scan
    .DESCRIPTION
        One-stop shop for all the assignment/sync drama.
    .EXAMPLE
        $allIssues = Get-GraphUserConflicts -AccessToken $token
        $allIssues | Export-Csv -Path ".\conflicts.csv" -NoTypeInformation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken
    )

    Write-Host "Scanning license assignment errors..." -ForegroundColor Cyan
    $licenseErrors = Get-GraphLicenseErrors -AccessToken $AccessToken -ErrorsOnly

    Write-Host "Scanning provisioning errors (sync conflicts)..." -ForegroundColor Cyan
    $provErrors = Get-GraphProvisioningErrors -AccessToken $AccessToken -ErrorsOnly

    Write-Host "Scanning for duplicate proxy addresses..." -ForegroundColor Cyan
    $dupeProxies = Find-DuplicateProxyAddresses -AccessToken $AccessToken | Where-Object IsDuplicate

    $summary = [PSCustomObject]@{
        LicenseErrors      = $licenseErrors
        ProvisioningErrors = $provErrors
        DuplicateProxies   = $dupeProxies
        Stats              = [PSCustomObject]@{
            UsersWithLicenseErrors = ($licenseErrors | Select-Object -Unique UserPrincipalName).Count
            UsersWithSyncErrors    = ($provErrors | Select-Object -Unique UserPrincipalName).Count
            DuplicateProxyCount    = $dupeProxies.Count
        }
    }

    Write-Host "`n=== CONFLICT SUMMARY ===" -ForegroundColor Yellow
    Write-Host "License Errors:      $($summary.Stats.UsersWithLicenseErrors) users"
    Write-Host "Sync/Prov Errors:    $($summary.Stats.UsersWithSyncErrors) users"
    Write-Host "Duplicate Proxies:   $($summary.Stats.DuplicateProxyCount) addresses"

    return $summary
}

function Resolve-GraphProxyConflict {
    <#
    .SYNOPSIS
        For users with proxy conflicts, finds WHO else owns the conflicting address
    .DESCRIPTION
        Takes provisioning errors and searches for each conflicting proxy address
        across all users' proxyAddresses. Shows the conflict owner.
    .EXAMPLE
        $conflicts = Resolve-GraphProxyConflict -AccessToken $token
        $conflicts | Format-Table UserWithError, ConflictingProxy, OwnedBy, OwnerType
    .EXAMPLE
        # For specific user
        Resolve-GraphProxyConflict -AccessToken $token -UserPrincipalName "deeg.jutta@elkw.de"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [string]$UserPrincipalName,

        [Parameter()]
        [string[]]$ProxyAddresses
    )

    begin {
        $headers = @{
            Authorization      = "Bearer $AccessToken"
            'Content-Type'     = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        try {
            # Get users with provisioning errors, or specific user
            if ($ProxyAddresses) {
                # Direct proxy lookup mode
                $proxiesToCheck = $ProxyAddresses
                $sourceUser = if ($UserPrincipalName) { $UserPrincipalName } else { 'Manual lookup' }
            }
            else {
                # Get provisioning errors first
                Write-Host "Fetching users with provisioning errors..." -ForegroundColor Cyan
                if ($UserPrincipalName) {
                    $provErrors = Get-GraphProvisioningErrors -AccessToken $AccessToken -ErrorsOnly -UserPrincipalName $UserPrincipalName
                }
                else {
                    $provErrors = Get-GraphProvisioningErrors -AccessToken $AccessToken -ErrorsOnly
                }
                
                if (-not $provErrors) {
                    Write-Host "No provisioning errors found." -ForegroundColor Green
                }

                # Also get group license errors with ProxyAddressConflict
                Write-Host "Fetching group license errors..." -ForegroundColor Cyan
                $groupErrors = @()
                try {
                    $allGroupErrors = Get-GraphGroupLicenseErrors -AccessToken $AccessToken
                    $groupErrors = $allGroupErrors | Where-Object Error -eq 'ProxyAddressConflict'
                }
                catch {
                    Write-Warning "Could not fetch group license errors: $($_.Exception.Message)"
                }
            }

            # Build list of users to check and their proxies
            $usersToCheck = @{}

            # From provisioning errors - the Value field contains the conflicting address
            foreach ($err in $provErrors) {
                if ($err.PropertyCausingError -match 'proxy' -and $err.Value) {
                    if (-not $usersToCheck.ContainsKey($err.UserPrincipalName)) {
                        $usersToCheck[$err.UserPrincipalName] = [System.Collections.Generic.List[string]]::new()
                    }
                    $usersToCheck[$err.UserPrincipalName].Add($err.Value)
                }
            }

            # From group license errors - need to get user's proxyAddresses
            foreach ($err in $groupErrors) {
                if (-not $usersToCheck.ContainsKey($err.UserPrincipalName)) {
                    $usersToCheck[$err.UserPrincipalName] = [System.Collections.Generic.List[string]]::new()
                    
                    # Fetch user's proxy addresses to check each one
                    try {
                        $userUri = "$baseUri/users/$($err.UserPrincipalName)?`$select=proxyAddresses"
                        $userData = Invoke-RestMethod -Uri $userUri -Headers $headers -Method Get
                        if ($userData.proxyAddresses) {
                            foreach ($p in $userData.proxyAddresses) {
                                $usersToCheck[$err.UserPrincipalName].Add($p)
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not fetch proxies for $($err.UserPrincipalName)"
                    }
                }
            }

            Write-Host "Checking $($usersToCheck.Count) users for proxy conflicts..." -ForegroundColor Cyan

            # For each user's proxies, search who else has them
            foreach ($upn in $usersToCheck.Keys) {
                foreach ($proxy in $usersToCheck[$upn]) {
                    Write-Verbose "Checking: $proxy (from $upn)"

                    # Normalize for search - Graph search is case-insensitive but let's be safe
                    $searchProxy = $proxy.Trim()
                    
                    # Try filter on proxyAddresses (requires ConsistencyLevel: eventual)
                    $owners = @()
                    
                    # Method 1: Direct filter (works for exact match)
                    try {
                        $filterUri = "$baseUri/users?`$filter=proxyAddresses/any(p:p eq '$searchProxy')&`$select=userPrincipalName,displayName,proxyAddresses,onPremisesSyncEnabled"
                        $filterResponse = Invoke-RestMethod -Uri $filterUri -Headers $headers -Method Get
                        if ($filterResponse.value) {
                            $owners += $filterResponse.value
                        }
                    }
                    catch {
                        Write-Verbose "Direct filter failed, trying search..."
                    }

                    # Method 2: Try with mail/SMTP part extracted
                    if (-not $owners -and $searchProxy -match ':(.+)$') {
                        $emailPart = $Matches[1]
                        try {
                            # Search by mail field
                            $mailUri = "$baseUri/users?`$filter=mail eq '$emailPart'&`$select=userPrincipalName,displayName,proxyAddresses,onPremisesSyncEnabled"
                            $mailResponse = Invoke-RestMethod -Uri $mailUri -Headers $headers -Method Get
                            if ($mailResponse.value) {
                                $owners += $mailResponse.value
                            }

                            # Also check userPrincipalName
                            $upnUri = "$baseUri/users?`$filter=userPrincipalName eq '$emailPart'&`$select=userPrincipalName,displayName,proxyAddresses,onPremisesSyncEnabled"
                            $upnResponse = Invoke-RestMethod -Uri $upnUri -Headers $headers -Method Get
                            if ($upnResponse.value) {
                                $owners += $upnResponse.value
                            }
                        }
                        catch {
                            Write-Verbose "Mail/UPN search failed for $emailPart"
                        }
                    }

                    # Method 3: Search in groups (distribution lists, M365 groups)
                    if ($searchProxy -match ':(.+)$') {
                        $emailPart = $Matches[1]
                        try {
                            $groupUri = "$baseUri/groups?`$filter=mail eq '$emailPart' or proxyAddresses/any(p:p eq '$searchProxy')&`$select=displayName,mail,proxyAddresses,groupTypes"
                            $groupResponse = Invoke-RestMethod -Uri $groupUri -Headers $headers -Method Get
                            if ($groupResponse.value) {
                                foreach ($g in $groupResponse.value) {
                                    $owners += [PSCustomObject]@{
                                        userPrincipalName    = "[GROUP] $($g.displayName)"
                                        displayName          = $g.displayName
                                        proxyAddresses       = $g.proxyAddresses
                                        onPremisesSyncEnabled = $null
                                        objectType           = 'Group'
                                        mail                 = $g.mail
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Group search failed for $emailPart"
                        }
                    }

                    # Dedupe and exclude the source user
                    $otherOwners = $owners | Where-Object { 
                        $_.userPrincipalName -and $_.userPrincipalName -ne $upn 
                    } | Select-Object -Unique userPrincipalName, displayName, proxyAddresses, onPremisesSyncEnabled, objectType, mail

                    foreach ($owner in $otherOwners) {
                        $results.Add([PSCustomObject]@{
                            UserWithError         = $upn
                            ConflictingProxy      = $proxy
                            OwnedBy               = $owner.userPrincipalName
                            OwnerDisplayName      = $owner.displayName
                            OwnerType             = if ($owner.objectType -eq 'Group') { 'Group' } elseif ($owner.onPremisesSyncEnabled) { 'Synced User' } else { 'Cloud User' }
                            OwnerMail             = $owner.mail
                            OwnerProxyAddresses   = ($owner.proxyAddresses -join '; ')
                            Resolution            = if ($owner.objectType -eq 'Group') { 
                                'Remove from group or rename group alias' 
                            } elseif ($owner.onPremisesSyncEnabled) { 
                                'Fix in on-premises AD, wait for sync' 
                            } else { 
                                'Edit cloud user proxy addresses' 
                            }
                        })
                    }

                    # If no other owner found, might be self-conflict or deleted object
                    if (-not $otherOwners) {
                        $results.Add([PSCustomObject]@{
                            UserWithError         = $upn
                            ConflictingProxy      = $proxy
                            OwnedBy               = '[NOT FOUND - possibly deleted object or soft-match conflict]'
                            OwnerDisplayName      = $null
                            OwnerType             = 'Unknown'
                            OwnerMail             = $null
                            OwnerProxyAddresses   = $null
                            Resolution            = 'Check deleted users, contacts, or soft-match issues'
                        })
                    }
                }
            }
        }
        catch {
            Write-Error "Error resolving proxy conflicts: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                Write-Error "Detail: $($errDetail.error.message)"
            }
        }
    }

    end {
        return $results
    }
}

function Search-GraphProxyAddress {
    <#
    .SYNOPSIS
        Searches for a specific proxy address across users, groups, and contacts
    .DESCRIPTION
        Manual lookup tool - "who owns smtp:xyz@domain.com?"
    .EXAMPLE
        Search-GraphProxyAddress -AccessToken $token -ProxyAddress "smtp:info@contoso.com"
    .EXAMPLE
        Search-GraphProxyAddress -AccessToken $token -EmailAddress "info@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Proxy')]
        [string]$ProxyAddress,

        [Parameter(Mandatory, ParameterSetName = 'Email')]
        [string]$EmailAddress,

        [Parameter(Mandatory)]
        [string]$AccessToken
    )

    begin {
        $headers = @{
            Authorization      = "Bearer $AccessToken"
            'Content-Type'     = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        # Normalize input
        if ($EmailAddress) {
            $searchEmail = $EmailAddress.Trim()
            $searchProxies = @("SMTP:$searchEmail", "smtp:$searchEmail")
        }
        else {
            $searchProxies = @($ProxyAddress.Trim())
            $searchEmail = if ($ProxyAddress -match ':(.+)$') { $Matches[1] } else { $ProxyAddress }
        }

        Write-Host "Searching for: $searchEmail" -ForegroundColor Cyan

        # Search Users
        Write-Host "  Checking users..." -ForegroundColor Gray
        try {
            foreach ($proxy in $searchProxies) {
                $uri = "$baseUri/users?`$filter=proxyAddresses/any(p:p eq '$proxy')&`$select=userPrincipalName,displayName,mail,proxyAddresses,onPremisesSyncEnabled"
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                foreach ($u in $response.value) {
                    $results.Add([PSCustomObject]@{
                        ObjectType        = 'User'
                        Identifier        = $u.userPrincipalName
                        DisplayName       = $u.displayName
                        Mail              = $u.mail
                        MatchedOn         = $proxy
                        IsSynced          = [bool]$u.onPremisesSyncEnabled
                        AllProxyAddresses = ($u.proxyAddresses -join '; ')
                    })
                }
            }

            # Also check mail field directly
            $mailUri = "$baseUri/users?`$filter=mail eq '$searchEmail'&`$select=userPrincipalName,displayName,mail,proxyAddresses,onPremisesSyncEnabled"
            $mailResponse = Invoke-RestMethod -Uri $mailUri -Headers $headers -Method Get
            foreach ($u in $mailResponse.value) {
                if (-not ($results | Where-Object { $_.Identifier -eq $u.userPrincipalName })) {
                    $results.Add([PSCustomObject]@{
                        ObjectType        = 'User'
                        Identifier        = $u.userPrincipalName
                        DisplayName       = $u.displayName
                        Mail              = $u.mail
                        MatchedOn         = "mail=$searchEmail"
                        IsSynced          = [bool]$u.onPremisesSyncEnabled
                        AllProxyAddresses = ($u.proxyAddresses -join '; ')
                    })
                }
            }
        }
        catch {
            Write-Verbose "User search error: $($_.Exception.Message)"
        }

        # Search Groups
        Write-Host "  Checking groups..." -ForegroundColor Gray
        try {
            $groupUri = "$baseUri/groups?`$filter=mail eq '$searchEmail'&`$select=id,displayName,mail,proxyAddresses,groupTypes"
            $groupResponse = Invoke-RestMethod -Uri $groupUri -Headers $headers -Method Get
            foreach ($g in $groupResponse.value) {
                $groupType = if ($g.groupTypes -contains 'Unified') { 'M365 Group' } else { 'Distribution/Security' }
                $results.Add([PSCustomObject]@{
                    ObjectType        = "Group ($groupType)"
                    Identifier        = $g.id
                    DisplayName       = $g.displayName
                    Mail              = $g.mail
                    MatchedOn         = "mail=$searchEmail"
                    IsSynced          = $null
                    AllProxyAddresses = ($g.proxyAddresses -join '; ')
                })
            }
        }
        catch {
            Write-Verbose "Group search error: $($_.Exception.Message)"
        }

        # Search Contacts
        Write-Host "  Checking contacts..." -ForegroundColor Gray
        try {
            $contactUri = "$baseUri/contacts?`$filter=mail eq '$searchEmail'&`$select=id,displayName,mail,proxyAddresses"
            $contactResponse = Invoke-RestMethod -Uri $contactUri -Headers $headers -Method Get
            foreach ($c in $contactResponse.value) {
                $results.Add([PSCustomObject]@{
                    ObjectType        = 'Contact'
                    Identifier        = $c.id
                    DisplayName       = $c.displayName
                    Mail              = $c.mail
                    MatchedOn         = "mail=$searchEmail"
                    IsSynced          = $null
                    AllProxyAddresses = ($c.proxyAddresses -join '; ')
                })
            }
        }
        catch {
            Write-Verbose "Contact search error: $($_.Exception.Message)"
        }

        Write-Host "  Found $($results.Count) object(s)" -ForegroundColor $(if ($results.Count -gt 1) { 'Yellow' } else { 'Green' })
    }

    end {
        return $results
    }
}

#endregion

#region Group-Based Licensing Errors

function Get-GraphGroupLicenseErrors {
    <#
    .SYNOPSIS
        Gets groups with license assignment errors and their affected members
    .DESCRIPTION
        Queries groups that have hasMembersWithLicenseErrors = true, then retrieves
        the members with errors. This matches the M365 Admin Center "Handlung erforderlich" view.
    .EXAMPLE
        $groupErrors = Get-GraphGroupLicenseErrors -AccessToken $token
        $groupErrors | Format-Table GroupName, UserPrincipalName, Error
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [string]$GroupId,

        [Parameter()]
        [string]$GroupDisplayName,

        [Parameter()]
        [int]$Top = 999
    )

    begin {
        $headers = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        try {
            # Get groups with license errors
            if ($GroupId) {
                $groups = @([PSCustomObject]@{ id = $GroupId; displayName = $GroupDisplayName })
            }
            else {
                $uri = "$baseUri/groups?`$filter=hasMembersWithLicenseErrors eq true&`$select=id,displayName,assignedLicenses,licenseProcessingState&`$top=$Top"
                $groups = [System.Collections.Generic.List[PSObject]]::new()

                do {
                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                    if ($response.value) {
                        foreach ($g in $response.value) { $groups.Add($g) }
                    }
                    $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
                } while ($uri)
            }

            Write-Host "Found $($groups.Count) group(s) with license errors" -ForegroundColor Cyan

            # For each group, get members with license errors
            foreach ($group in $groups) {
                Write-Verbose "Processing group: $($group.displayName)"
                
                $membersUri = "$baseUri/groups/$($group.id)/membersWithLicenseErrors?`$select=id,userPrincipalName,displayName,licenseAssignmentStates"
                
                do {
                    $memberResponse = Invoke-RestMethod -Uri $membersUri -Headers $headers -Method Get
                    
                    if ($memberResponse.value) {
                        foreach ($member in $memberResponse.value) {
                            # Get the specific error for this group assignment
                            $groupErrors = $member.licenseAssignmentStates | Where-Object { 
                                $_.assignedByGroup -eq $group.id -and 
                                (-not [string]::IsNullOrEmpty($_.error)) -and 
                                $_.error -ne 'None'
                            }

                            foreach ($err in $groupErrors) {
                                $results.Add([PSCustomObject]@{
                                    GroupId              = $group.id
                                    GroupName            = $group.displayName
                                    UserId               = $member.id
                                    UserPrincipalName    = $member.userPrincipalName
                                    DisplayName          = $member.displayName
                                    SkuId                = $err.skuId
                                    Error                = $err.error
                                    State                = $err.state
                                    LastUpdatedDateTime  = $err.lastUpdatedDateTime
                                })
                            }
                        }
                    }
                    
                    $membersUri = if ($memberResponse.PSObject.Properties['@odata.nextLink']) { $memberResponse.'@odata.nextLink' } else { $null }
                } while ($membersUri)
            }
        }
        catch {
            Write-Error "Graph API call failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                Write-Error "Detail: $($errDetail.error.message)"
            }
        }
    }

    end {
        return $results
    }
}

function Get-GraphGroupsWithLicenseErrors {
    <#
    .SYNOPSIS
        Lists all groups that have license assignment errors (summary view)
    .DESCRIPTION
        Quick overview of which license groups have problems - matches left panel in Admin Center
    .EXAMPLE
        Get-GraphGroupsWithLicenseErrors -AccessToken $token | Format-Table DisplayName, ErrorCount
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [int]$Top = 999
    )

    begin {
        $headers = @{
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        try {
            $uri = "$baseUri/groups?`$filter=hasMembersWithLicenseErrors eq true&`$select=id,displayName,assignedLicenses,licenseProcessingState&`$top=$Top"

            do {
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                if ($response.value) {
                    foreach ($group in $response.value) {
                        # Get count of members with errors
                        $countUri = "$baseUri/groups/$($group.id)/membersWithLicenseErrors/`$count"
                        try {
                            $errorCount = Invoke-RestMethod -Uri $countUri -Headers $headers -Method Get
                        }
                        catch {
                            $errorCount = '?'
                        }

                        $results.Add([PSCustomObject]@{
                            GroupId               = $group.id
                            DisplayName           = $group.displayName
                            AssignedLicenseCount  = ($group.assignedLicenses | Measure-Object).Count
                            LicenseSkuIds         = ($group.assignedLicenses.skuId -join '; ')
                            ProcessingState       = $group.licenseProcessingState
                            MembersWithErrors     = $errorCount
                        })
                    }
                }
                $uri = if ($response.PSObject.Properties['@odata.nextLink']) { $response.'@odata.nextLink' } else { $null }
            } while ($uri)
        }
        catch {
            Write-Error "Graph API call failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                Write-Error "Detail: $($errDetail.error.message)"
            }
        }
    }

    end {
        return $results
    }
}

function Get-GraphLicenseErrorReport {
    <#
    .SYNOPSIS
        Complete license error report: user-level + group-level + provisioning errors
    .DESCRIPTION
        One command to get everything for the L2 handoff
    .EXAMPLE
        $report = Get-GraphLicenseErrorReport -AccessToken $token
        $report.GroupErrors | Export-Csv ".\group-license-errors.csv" -NoTypeInformation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter()]
        [switch]$IncludeDuplicateProxyScan
    )

    Write-Host "`n=== LICENSE ERROR REPORT ===" -ForegroundColor Yellow
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

    Write-Host "[1/4] Scanning groups with license errors..." -ForegroundColor Cyan
    $groupSummary = Get-GraphGroupsWithLicenseErrors -AccessToken $AccessToken
    
    Write-Host "[2/4] Getting affected users per group..." -ForegroundColor Cyan
    $groupErrors = Get-GraphGroupLicenseErrors -AccessToken $AccessToken

    Write-Host "[3/4] Scanning user-level license assignment states..." -ForegroundColor Cyan
    $userErrors = Get-GraphLicenseErrors -AccessToken $AccessToken -ErrorsOnly

    Write-Host "[4/4] Scanning provisioning errors (sync conflicts)..." -ForegroundColor Cyan
    $provErrors = Get-GraphProvisioningErrors -AccessToken $AccessToken -ErrorsOnly

    $dupeProxies = $null
    if ($IncludeDuplicateProxyScan) {
        Write-Host "[+] Running full duplicate proxy scan..." -ForegroundColor Cyan
        $dupeProxies = Find-DuplicateProxyAddresses -AccessToken $AccessToken | Where-Object IsDuplicate
    }

    # Get SKU friendly names for mapping
    $skuMap = Get-GraphSkuFriendlyName -AccessToken $AccessToken

    # Enrich with friendly names
    $groupErrors | ForEach-Object {
        $licenseName = if ($skuMap[$_.SkuId]) { $skuMap[$_.SkuId] } else { $_.SkuId }
        $_ | Add-Member -NotePropertyName 'LicenseName' -NotePropertyValue $licenseName -Force
    }
    $userErrors | ForEach-Object {
        $licenseName = if ($skuMap[$_.SkuId]) { $skuMap[$_.SkuId] } else { $_.SkuId }
        $_ | Add-Member -NotePropertyName 'LicenseName' -NotePropertyValue $licenseName -Force
    }

    $report = [PSCustomObject]@{
        GeneratedAt        = Get-Date
        GroupSummary       = $groupSummary
        GroupErrors        = $groupErrors
        UserErrors         = $userErrors
        ProvisioningErrors = $provErrors
        DuplicateProxies   = $dupeProxies
        SkuMap             = $skuMap
        Stats              = [PSCustomObject]@{
            GroupsWithErrors       = ($groupSummary | Measure-Object).Count
            UsersInGroupErrors     = ($groupErrors | Select-Object -Unique UserPrincipalName | Measure-Object).Count
            DirectAssignmentErrors = ($userErrors | Select-Object -Unique UserPrincipalName | Measure-Object).Count
            ProvisioningErrors     = ($provErrors | Select-Object -Unique UserPrincipalName | Measure-Object).Count
            DuplicateProxies       = if ($dupeProxies) { ($dupeProxies | Measure-Object).Count } else { 'Not scanned' }
        }
    }

    Write-Host "`n=== SUMMARY ===" -ForegroundColor Yellow
    Write-Host "Groups with errors:        $($report.Stats.GroupsWithErrors)"
    Write-Host "Users (group licensing):   $($report.Stats.UsersInGroupErrors)"
    Write-Host "Users (direct assignment): $($report.Stats.DirectAssignmentErrors)"
    Write-Host "Users (sync/provisioning): $($report.Stats.ProvisioningErrors)"
    Write-Host "Duplicate proxy addresses: $($report.Stats.DuplicateProxies)"
    Write-Host ""

    # Show error breakdown
    if ($groupErrors) {
        Write-Host "Error Types (Group-Based):" -ForegroundColor Cyan
        $groupErrors | Group-Object Error | Sort-Object Count -Descending | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count)"
        }
    }

    return $report
}

#endregion

#region Credential Obfuscation (XOR + Base64)

function Protect-CredentialJson {
    <#
    .SYNOPSIS
        Creates obfuscated JSON credential file for handoff
    .DESCRIPTION
        XOR + Base64 obfuscation. NOT encryption - just keeps plaintext out of sight.
        Share the JSON file + key separately (e.g., file via Teams, key via phone).
    .EXAMPLE
        Protect-CredentialJson -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz" -Key "L2SharedKey2024" -OutputPath ".\creds.json"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$Key,
        [Parameter()][string]$OutputPath = ".\graph-creds.json"
    )

    $xorEncode = {
        param([string]$Text, [string]$XorKey)
        $textBytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($XorKey)
        $result = [byte[]]::new($textBytes.Length)
        for ($i = 0; $i -lt $textBytes.Length; $i++) {
            $result[$i] = $textBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        return [Convert]::ToBase64String($result)
    }

    $payload = @{
        _meta      = "XOR+B64 obfuscated - not encryption"
        _created   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        tenantId   = & $xorEncode $TenantId $Key
        clientId   = & $xorEncode $ClientId $Key
        secret     = & $xorEncode $ClientSecret $Key
    }

    $payload | ConvertTo-Json -Depth 2 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "Credential file written: $OutputPath" -ForegroundColor Green
    Write-Host "Share file and key via SEPARATE channels!" -ForegroundColor Yellow

    return $OutputPath
}

function Unprotect-CredentialJson {
    <#
    .SYNOPSIS
        Reads obfuscated credential JSON and returns usable values
    .EXAMPLE
        $creds = Unprotect-CredentialJson -Path ".\creds.json" -Key "L2SharedKey2024"
        $token = Get-GraphToken @creds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Key
    )

    $xorDecode = {
        param([string]$EncodedText, [string]$XorKey)
        $encBytes = [Convert]::FromBase64String($EncodedText)
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($XorKey)
        $result = [byte[]]::new($encBytes.Length)
        for ($i = 0; $i -lt $encBytes.Length; $i++) {
            $result[$i] = $encBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        return [System.Text.Encoding]::UTF8.GetString($result)
    }

    $json = Get-Content -Path $Path -Raw | ConvertFrom-Json

    # Return hashtable for splatting compatibility
    return @{
        TenantId     = & $xorDecode $json.tenantId $Key
        ClientId     = & $xorDecode $json.clientId $Key
        ClientSecret = & $xorDecode $json.secret $Key
    }
}

function Get-GraphToken {
    <#
    .SYNOPSIS
        Gets OAuth2 token for Graph API using client credentials
    .EXAMPLE
        $creds = Unprotect-CredentialJson -Path ".\creds.json" -Key "MyKey"
        $token = Get-GraphToken -TenantId $creds.TenantId -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        $errors = Get-GraphLicenseErrors -AccessToken $token -ErrorsOnly
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter()][string]$Scope = "https://graph.microsoft.com/.default"
    )

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = $Scope
    }

    try {
        $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Token acquisition failed: $($_.Exception.Message)"
        return $null
    }
}

#endregion

#region Conflict Resolution Pipeline

function Get-ProxyConflictAnalysis {
    <#
    .SYNOPSIS
        Deep analysis of proxy conflicts with intelligent resolution logic
    .DESCRIPTION
        Analyzes conflicts and determines winner/loser based on:
        1. Eineindeutigkeit: SAM/identity matches email prefix = rightful owner
        2. Object Type: RemoteMailbox > Contact
        3. Activity Status: Active users > Inactive (deaktiviert/ausgeschieden/inaktiv)
        4. Seniority: Older user (by sync/creation date) wins ties
        
        Generates smart suffixes:
        - .inact for inactive/ausgeschieden users
        - .<sam> for identity-misaligned users
        - .dup for other duplicates
    .EXAMPLE
        $analysis = Get-ProxyConflictAnalysis -AccessToken $token -Conflicts $conflicts
        $analysis | Format-Table Winner, Loser, Resolution, NewProxyForLoser
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject[]]$Conflicts
    )

    begin {
        $headers = @{
            Authorization      = "Bearer $AccessToken"
            'Content-Type'     = 'application/json'
            'ConsistencyLevel' = 'eventual'
        }
        $baseUri = 'https://graph.microsoft.com/v1.0'
        $results = [System.Collections.Generic.List[PSObject]]::new()
        
        # Cache for user lookups
        $userCache = @{}
        
        # Inactive patterns (German + English)
        $inactivePattern = '(?i)(deactivated|deaktiviert|inactive|inaktiv|ausgeschieden|disabled|ehemalig|former|terminated|gesperrt)'
    }

    process {
        foreach ($conflict in $Conflicts) {
            Write-Verbose "Analyzing conflict: $($conflict.ConflictingProxy)"
            
            $selectProps = 'id,userPrincipalName,displayName,mail,proxyAddresses,onPremisesSyncEnabled,onPremisesDistinguishedName,onPremisesSamAccountName,onPremisesUserPrincipalName,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,accountEnabled,createdDateTime'
            
            # Fetch user with error
            $userWithError = $null
            try {
                if (-not $userCache.ContainsKey($conflict.UserWithError)) {
                    $uri = "$baseUri/users/$($conflict.UserWithError)?`$select=$selectProps"
                    $userCache[$conflict.UserWithError] = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                }
                $userWithError = $userCache[$conflict.UserWithError]
            }
            catch {
                Write-Warning "Could not fetch user with error: $($conflict.UserWithError)"
                continue
            }
            
            # Fetch conflict owner (if user)
            $conflictOwner = $null
            $ownerIsContact = $false
            
            if ($conflict.OwnedBy -and $conflict.OwnedBy -notmatch '^\[NOT FOUND') {
                if ($conflict.OwnerType -eq 'Group') {
                    # Group owns it - user must change
                    $results.Add([PSCustomObject]@{
                        ConflictId           = [guid]::NewGuid().ToString('N').Substring(0, 8)
                        ConflictingProxy     = $conflict.ConflictingProxy
                        ConflictType         = 'Group-Collision'
                        Priority             = 2
                        Winner               = "[Group: $($conflict.OwnedBy)]"
                        WinnerSAM            = $null
                        WinnerDN             = $null
                        WinnerReason         = 'Groups have precedence - remove from group or change user'
                        Loser                = $conflict.UserWithError
                        LoserSAM             = $userWithError.onPremisesSamAccountName
                        LoserDN              = $userWithError.onPremisesDistinguishedName
                        LoserCurrentProxies  = ($userWithError.proxyAddresses -join '; ')
                        Resolution           = 'Manual - Remove address from group or use different alias for user'
                        NewProxyForLoser     = $null
                        ADCommand            = $null
                        RollbackData         = $null
                        Status               = 'Pending'
                    })
                    continue
                }
                
                # Check if owner is a contact
                if ($conflict.OwnerType -match 'Contact') {
                    $ownerIsContact = $true
                }
                else {
                    try {
                        if (-not $userCache.ContainsKey($conflict.OwnedBy)) {
                            $uri = "$baseUri/users/$($conflict.OwnedBy)?`$select=$selectProps"
                            $userCache[$conflict.OwnedBy] = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                        }
                        $conflictOwner = $userCache[$conflict.OwnedBy]
                    }
                    catch {
                        Write-Warning "Could not fetch conflict owner: $($conflict.OwnedBy)"
                    }
                }
            }
            
            # Extract email prefix from conflicting proxy
            $conflictProxy = $conflict.ConflictingProxy
            $emailPart = if ($conflictProxy -match ':(.+)@') { $Matches[1].ToLower() } else { $null }
            
            # === DECISION LOGIC ===
            $winner = $null
            $loser = $null
            $winnerUser = $null
            $loserUser = $null
            $winReason = $null
            $resolution = $null
            $suffix = $null
            
            # --- Rule 1: Contact vs User - User (RemoteMailbox) wins ---
            if ($ownerIsContact -and $userWithError) {
                $winner = $conflict.UserWithError
                $loser = $conflict.OwnedBy
                $winnerUser = $userWithError
                $loserUser = $null
                $winReason = 'RemoteMailbox has precedence over Contact'
                $resolution = 'Delete or modify Contact, user keeps address'
                
                $results.Add([PSCustomObject]@{
                    ConflictId           = [guid]::NewGuid().ToString('N').Substring(0, 8)
                    ConflictingProxy     = $conflictProxy
                    ConflictType         = 'Contact-Collision'
                    Priority             = 1
                    Winner               = $winner
                    WinnerSAM            = $userWithError.onPremisesSamAccountName
                    WinnerDN             = $userWithError.onPremisesDistinguishedName
                    WinnerReason         = $winReason
                    Loser                = $loser
                    LoserSAM             = $null
                    LoserDN              = $null
                    LoserCurrentProxies  = $conflict.OwnerProxyAddresses
                    Resolution           = $resolution
                    NewProxyForLoser     = $null
                    ADCommand            = "# Contact must be deleted or modified in Exchange/AD`n# Remove-MailContact -Identity '$loser' or modify proxyAddresses"
                    RollbackData         = $null
                    Status               = 'Pending'
                })
                continue
            }
            
            if (-not $conflictOwner) {
                # Orphan collision - can't determine winner
                $results.Add([PSCustomObject]@{
                    ConflictId           = [guid]::NewGuid().ToString('N').Substring(0, 8)
                    ConflictingProxy     = $conflictProxy
                    ConflictType         = 'Orphan-Collision'
                    Priority             = 1
                    Winner               = '[Unknown]'
                    WinnerSAM            = $null
                    WinnerDN             = $null
                    WinnerReason         = 'Owner not found - check deleted users, soft-match, or contacts'
                    Loser                = $conflict.UserWithError
                    LoserSAM             = $userWithError.onPremisesSamAccountName
                    LoserDN              = $userWithError.onPremisesDistinguishedName
                    LoserCurrentProxies  = ($userWithError.proxyAddresses -join '; ')
                    Resolution           = 'Manual investigation required'
                    NewProxyForLoser     = $null
                    ADCommand            = $null
                    RollbackData         = $null
                    Status               = 'Pending'
                })
                continue
            }
            
            # Both are users - apply full logic
            $user1 = $userWithError
            $user2 = $conflictOwner
            
            # Get SAMs
            $sam1 = if ($user1.onPremisesSamAccountName) { $user1.onPremisesSamAccountName.ToLower() } else { '' }
            $sam2 = if ($user2.onPremisesSamAccountName) { $user2.onPremisesSamAccountName.ToLower() } else { '' }
            
            # Check Eineindeutigkeit: Does SAM match email prefix?
            $user1IsEineindeutig = ($sam1 -eq $emailPart)
            $user2IsEineindeutig = ($sam2 -eq $emailPart)
            
            # Check activity status from DN
            $dn1 = if ($user1.onPremisesDistinguishedName) { $user1.onPremisesDistinguishedName } else { '' }
            $dn2 = if ($user2.onPremisesDistinguishedName) { $user2.onPremisesDistinguishedName } else { '' }
            
            $user1IsInactive = ($dn1 -match $inactivePattern) -or (-not $user1.accountEnabled)
            $user2IsInactive = ($dn2 -match $inactivePattern) -or (-not $user2.accountEnabled)
            
            # Get creation/sync dates for seniority
            $date1 = if ($user1.onPremisesLastSyncDateTime) { 
                [DateTime]$user1.onPremisesLastSyncDateTime 
            } elseif ($user1.createdDateTime) { 
                [DateTime]$user1.createdDateTime 
            } else { 
                [DateTime]::MaxValue 
            }
            
            $date2 = if ($user2.onPremisesLastSyncDateTime) { 
                [DateTime]$user2.onPremisesLastSyncDateTime 
            } elseif ($user2.createdDateTime) { 
                [DateTime]$user2.createdDateTime 
            } else { 
                [DateTime]::MaxValue 
            }
            
            # === APPLY RULES IN ORDER ===
            
            # Rule 2: Active vs Inactive - Active wins
            if ($user1IsInactive -and -not $user2IsInactive) {
                $winner = $user2.userPrincipalName
                $loser = $user1.userPrincipalName
                $winnerUser = $user2
                $loserUser = $user1
                $winReason = "User1 is inactive (DN matches: $inactivePattern)"
                $suffix = '.inact'
            }
            elseif ($user2IsInactive -and -not $user1IsInactive) {
                $winner = $user1.userPrincipalName
                $loser = $user2.userPrincipalName
                $winnerUser = $user1
                $loserUser = $user2
                $winReason = "User2 is inactive (DN matches: $inactivePattern)"
                $suffix = '.inact'
            }
            # Rule 3: Eineindeutigkeit - SAM matches email wins
            elseif ($user1IsEineindeutig -and -not $user2IsEineindeutig) {
                $winner = $user1.userPrincipalName
                $loser = $user2.userPrincipalName
                $winnerUser = $user1
                $loserUser = $user2
                $winReason = "User1 SAM ($sam1) matches email prefix - Eineindeutig"
                $suffix = ".$sam2"  # Loser gets their own SAM as suffix
            }
            elseif ($user2IsEineindeutig -and -not $user1IsEineindeutig) {
                $winner = $user2.userPrincipalName
                $loser = $user1.userPrincipalName
                $winnerUser = $user2
                $loserUser = $user1
                $winReason = "User2 SAM ($sam2) matches email prefix - Eineindeutig"
                $suffix = ".$sam1"  # Loser gets their own SAM as suffix
            }
            # Rule 4: Seniority - Older user wins
            elseif ($date1 -lt $date2) {
                $winner = $user1.userPrincipalName
                $loser = $user2.userPrincipalName
                $winnerUser = $user1
                $loserUser = $user2
                $winReason = "User1 is older (synced/created: $($date1.ToString('yyyy-MM-dd')))"
                $suffix = if ($sam2) { ".$sam2" } else { '.dup' }
            }
            elseif ($date2 -lt $date1) {
                $winner = $user2.userPrincipalName
                $loser = $user1.userPrincipalName
                $winnerUser = $user2
                $loserUser = $user1
                $winReason = "User2 is older (synced/created: $($date2.ToString('yyyy-MM-dd')))"
                $suffix = if ($sam1) { ".$sam1" } else { '.dup' }
            }
            # Rule 5: Tie - User with error loses (they're the one failing)
            else {
                $winner = $user2.userPrincipalName
                $loser = $user1.userPrincipalName
                $winnerUser = $user2
                $loserUser = $user1
                $winReason = "Tie-breaker: Existing owner keeps address"
                $suffix = if ($sam1) { ".$sam1" } else { '.dup' }
            }
            
            # Generate new proxy for loser
            $newProxy = $null
            $adCommand = $null
            $rollbackData = $null
            
            if ($loserUser -and $loserUser.onPremisesSamAccountName) {
                # Build new proxy address with smart suffix
                if ($conflictProxy -match '^(smtp:|SMTP:)([^@]+)@(.+)$') {
                    $prefix = $Matches[1]
                    $localPart = $Matches[2]
                    $domain = $Matches[3]
                    $newProxy = "$prefix$localPart$suffix@$domain"
                }
                
                # Determine conflict type
                $conflictType = if ($conflictProxy -match '\.mail\.onmicrosoft\.com$') {
                    'Routing-Address-Collision'
                } elseif ($conflictProxy -cmatch '^SMTP:') {
                    'Primary-SMTP-Collision'
                } else {
                    'Alias-Collision'
                }
                
                $loserSAM = $loserUser.onPremisesSamAccountName
                $currentProxy = $conflictProxy
                
                $adCommand = @"
# Resolution for $loser
# Reason: $winReason
# Winner: $winner

`$user = Get-ADUser -Identity '$loserSAM' -Properties proxyAddresses, targetAddress, mailNickname
Write-Host "Current state for $loserSAM :" -ForegroundColor Cyan
Write-Host "  mailNickname: `$(`$user.mailNickname)"
Write-Host "  proxyAddresses:" 
`$user.proxyAddresses | ForEach-Object { Write-Host "    `$_" }

# Build new proxy list - replace conflicting address
`$newProxies = @()
foreach (`$p in `$user.proxyAddresses) {
    if (`$p -eq '$currentProxy') {
        `$newProxies += '$newProxy'
        Write-Host "Changing: `$p -> $newProxy" -ForegroundColor Yellow
    } else {
        `$newProxies += `$p
    }
}

# Apply proxy change
Set-ADUser -Identity '$loserSAM' -Replace @{proxyAddresses = `$newProxies}

# Update targetAddress if it was the routing address
if (`$user.targetAddress -and `$user.targetAddress -match '$($currentProxy -replace '^smtp:', '' -replace '^SMTP:', '')') {
    `$newTarget = 'SMTP:$($newProxy -replace '^smtp:', '' -replace '^SMTP:', '')'
    Set-ADUser -Identity '$loserSAM' -Replace @{targetAddress = `$newTarget}
    Write-Host "Updated targetAddress -> `$newTarget" -ForegroundColor Yellow
}

Write-Host "`nDone. Trigger Azure AD Connect sync or wait for scheduled sync." -ForegroundColor Green
"@

                $rollbackData = @{
                    SamAccountName  = $loserSAM
                    OriginalProxies = $loserUser.proxyAddresses
                    OriginalTarget  = $null  # Would need to fetch
                }
                
                $resolution = "Change loser's proxy from $currentProxy to $newProxy"
            }
            else {
                $resolution = "Manual: Loser not synced or SAM unknown"
                $conflictType = 'Unknown'
            }
            
            $results.Add([PSCustomObject]@{
                ConflictId           = [guid]::NewGuid().ToString('N').Substring(0, 8)
                ConflictingProxy     = $conflictProxy
                ConflictType         = $conflictType
                Priority             = if ($loserUser.accountEnabled -eq $false -or $suffix -eq '.inact') { 3 } else { 1 }
                
                Winner               = $winner
                WinnerSAM            = $winnerUser.onPremisesSamAccountName
                WinnerDN             = $winnerUser.onPremisesDistinguishedName
                WinnerIsEineindeutig = ($winnerUser -eq $user1 -and $user1IsEineindeutig) -or ($winnerUser -eq $user2 -and $user2IsEineindeutig)
                WinnerSyncDate       = if ($winnerUser -eq $user1) { $date1 } else { $date2 }
                WinnerReason         = $winReason
                
                Loser                = $loser
                LoserSAM             = if ($loserUser) { $loserUser.onPremisesSamAccountName } else { $null }
                LoserDN              = if ($loserUser) { $loserUser.onPremisesDistinguishedName } else { $null }
                LoserIsInactive      = ($loserUser -eq $user1 -and $user1IsInactive) -or ($loserUser -eq $user2 -and $user2IsInactive)
                LoserSyncDate        = if ($loserUser -eq $user1) { $date1 } else { $date2 }
                LoserCurrentProxies  = if ($loserUser) { ($loserUser.proxyAddresses -join '; ') } else { $null }
                
                Suffix               = $suffix
                NewProxyForLoser     = $newProxy
                Resolution           = $resolution
                ADCommand            = $adCommand
                RollbackData         = $rollbackData
                Status               = 'Pending'
            })
        }
    }

    end {
        Write-Host "`n=== ANALYSIS SUMMARY ===" -ForegroundColor Yellow
        $byType = $results | Group-Object ConflictType
        foreach ($g in $byType) {
            Write-Host "  $($g.Name): $($g.Count)" -ForegroundColor Cyan
        }
        
        $byReason = $results | Group-Object WinnerReason
        Write-Host "`nDecision reasons:" -ForegroundColor Yellow
        foreach ($g in $byReason) {
            Write-Host "  $($g.Name): $($g.Count)" -ForegroundColor White
        }
        
        return $results | Sort-Object Priority, ConflictType
    }
}

function New-ProxyConflictResolutionPlan {
    <#
    .SYNOPSIS
        Creates an executable resolution plan from conflict analysis
    .DESCRIPTION
        Takes analyzed conflicts (from Get-ProxyConflictAnalysis) and generates a plan with:
        - Pre-generated AD commands from analysis
        - Execution order based on priority
        - Smart suffixes (.inact, .$sam, .dup)
        - Rollback data preservation
        - WhatIf preview mode
        
        Resolution logic (from analysis):
        1. Active user beats inactive (ausgeschieden/deaktiviert)
        2. Eineindeutig user (SAM=email) wins
        3. Older user (by sync date) wins ties
        4. Existing owner wins final ties
    .EXAMPLE
        $plan = New-ProxyConflictResolutionPlan -Analysis $analysis -WhatIf
        $plan = New-ProxyConflictResolutionPlan -Analysis $analysis -OutputPath "C:\temp\plan.json"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject[]]$Analysis,

        [Parameter()]
        [string]$OutputPath
    )

    begin {
        $plan = [System.Collections.Generic.List[PSObject]]::new()
        $order = 0
    }

    process {
        foreach ($item in $Analysis) {
            $order++
            
            # Use pre-generated values from analysis
            $action = $item.Resolution
            $adCommand = $item.ADCommand
            $rollbackData = $item.RollbackData
            $isManual = [string]::IsNullOrEmpty($adCommand) -or $item.ConflictType -match 'Manual|Group|Orphan|Contact'
            
            # Build human-readable summary
            $summary = @"
Winner: $($item.Winner)
  Reason: $($item.WinnerReason)
  SAM: $($item.WinnerSAM)
$(if ($item.WinnerIsEineindeutig) { "  [Eineindeutig - SAM matches email]" })

Loser: $($item.Loser)
  SAM: $($item.LoserSAM)
  Current proxy: $($item.ConflictingProxy)
  New proxy: $($item.NewProxyForLoser)
  Suffix applied: $($item.Suffix)
$(if ($item.LoserIsInactive) { "  [INACTIVE - ausgeschieden/deaktiviert]" })
"@
            
            $planItem = [PSCustomObject]@{
                Order            = $order
                ConflictId       = $item.ConflictId
                ConflictType     = $item.ConflictType
                Priority         = $item.Priority
                
                # Winner info
                Winner           = $item.Winner
                WinnerSAM        = $item.WinnerSAM
                WinnerReason     = $item.WinnerReason
                
                # Loser info (the one who changes)
                Loser            = $item.Loser
                LoserSAM         = $item.LoserSAM
                LoserIsInactive  = $item.LoserIsInactive
                
                # For backwards compatibility
                UserWithError    = $item.Loser
                SamAccountName   = $item.LoserSAM
                ConflictOwner    = $item.Winner
                
                # Proxy change
                ConflictingProxy = $item.ConflictingProxy
                NewProxy         = $item.NewProxyForLoser
                Suffix           = $item.Suffix
                
                # Execution
                Summary          = $summary
                Action           = $action
                ADCommand        = $adCommand
                RollbackData     = $rollbackData
                Status           = if ($isManual) { 'Manual' } else { 'Ready' }
                ExecutedAt       = $null
                ExecutedBy       = $null
                Error            = $null
            }
            
            if ($PSCmdlet.ShouldProcess("$($item.Loser) -> $($item.NewProxyForLoser)", $action)) {
                Write-Host "`n[$order] $($item.ConflictType) - Priority $($item.Priority)" -ForegroundColor Cyan
                Write-Host "    Proxy: $($item.ConflictingProxy)" -ForegroundColor White
                Write-Host "    Winner: $($item.Winner)" -ForegroundColor Green
                Write-Host "      Reason: $($item.WinnerReason)" -ForegroundColor DarkGreen
                Write-Host "    Loser: $($item.Loser)" -ForegroundColor Yellow
                if ($item.NewProxyForLoser) {
                    Write-Host "      New proxy: $($item.NewProxyForLoser) (suffix: $($item.Suffix))" -ForegroundColor Yellow
                }
                if ($item.LoserIsInactive) {
                    Write-Host "      [INACTIVE USER - ausgeschieden/deaktiviert]" -ForegroundColor DarkYellow
                }
                
                if (-not $isManual) {
                    Write-Host "    Status: Ready for execution" -ForegroundColor Green
                }
                else {
                    Write-Host "    Status: MANUAL intervention required" -ForegroundColor Red
                    Write-Host "      $action" -ForegroundColor DarkRed
                }
            }
            
            $plan.Add($planItem)
        }
    }

    end {
        if ($OutputPath) {
            # Convert rollback data for JSON serialization
            $exportPlan = $plan | ForEach-Object {
                $copy = $_ | Select-Object *
                if ($copy.RollbackData -and $copy.RollbackData.OriginalProxies) {
                    $copy.RollbackData = @{
                        SamAccountName  = $copy.RollbackData.SamAccountName
                        OriginalProxies = @($copy.RollbackData.OriginalProxies)
                    }
                }
                $copy
            }
            $exportPlan | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputPath -Encoding UTF8
            Write-Host "`nPlan saved to: $OutputPath" -ForegroundColor Green
        }
        
        Write-Host "`n" + ('=' * 50) -ForegroundColor Yellow
        Write-Host "RESOLUTION PLAN SUMMARY" -ForegroundColor Yellow
        Write-Host ('=' * 50) -ForegroundColor Yellow
        
        $ready = @($plan | Where-Object Status -eq 'Ready')
        $manual = @($plan | Where-Object Status -eq 'Manual')
        
        Write-Host "`n  Ready for execution: $($ready.Count)" -ForegroundColor Green
        Write-Host "  Requires manual action: $($manual.Count)" -ForegroundColor Yellow
        
        # Group by decision reason
        $byReason = $plan | Where-Object { $_.WinnerReason } | Group-Object WinnerReason
        if ($byReason) {
            Write-Host "`n  Decisions by reason:" -ForegroundColor Cyan
            foreach ($g in $byReason) {
                Write-Host "    $($g.Name): $($g.Count)" -ForegroundColor White
            }
        }
        
        # Group by suffix
        $bySuffix = $plan | Where-Object { $_.Suffix } | Group-Object Suffix
        if ($bySuffix) {
            Write-Host "`n  Suffixes applied:" -ForegroundColor Cyan
            foreach ($g in $bySuffix) {
                Write-Host "    $($g.Name): $($g.Count)" -ForegroundColor White
            }
        }
        
        # List inactive users being renamed
        $inactive = @($plan | Where-Object LoserIsInactive)
        if ($inactive.Count -gt 0) {
            Write-Host "`n  Inactive users being renamed ($($inactive.Count)):" -ForegroundColor DarkYellow
            foreach ($i in $inactive) {
                Write-Host "    $($i.Loser) -> suffix $($i.Suffix)" -ForegroundColor DarkYellow
            }
        }
        
        return $plan
    }
}

function Invoke-ProxyConflictResolution {
    <#
    .SYNOPSIS
        Executes resolution plan items one by one with audit logging
    .DESCRIPTION
        Runs AD commands from the resolution plan with:
        - Step-by-step confirmation
        - Audit CSV with timestamps
        - Rollback script generation
        - Error handling and continuation
    .EXAMPLE
        # WhatIf mode - preview only
        Invoke-ProxyConflictResolution -Plan $plan -WhatIf

        # Execute specific item
        Invoke-ProxyConflictResolution -Plan $plan -ConflictId "abc123"

        # Execute all with confirmation
        Invoke-ProxyConflictResolution -Plan $plan -AuditPath "C:\temp\resolution-audit.csv"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [PSObject[]]$Plan,

        [Parameter()]
        [string]$ConflictId,

        [Parameter()]
        [string]$AuditPath = ".\ProxyConflictResolution-Audit-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv",

        [Parameter()]
        [string]$RollbackScriptPath = ".\ProxyConflictResolution-Rollback-$(Get-Date -Format 'yyyyMMdd-HHmmss').ps1",

        [Parameter()]
        [switch]$Force
    )

    begin {
        $auditLog = [System.Collections.Generic.List[PSObject]]::new()
        $rollbackCommands = [System.Collections.Generic.List[string]]::new()
        $rollbackCommands.Add("# Rollback script generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $rollbackCommands.Add("# Run this script to undo all changes`n")
        
        $itemsToProcess = if ($ConflictId) {
            $Plan | Where-Object ConflictId -eq $ConflictId
        } else {
            # Include both 'Ready' (new) and 'Pending' (legacy) statuses
            $Plan | Where-Object { $_.Status -eq 'Ready' -or $_.Status -eq 'Pending' }
        }
        
        if (-not $itemsToProcess) {
            Write-Warning "No items to process (looking for Status = 'Ready' or 'Pending')"
            return
        }
        
        Write-Host "`n" + ('=' * 60) -ForegroundColor Yellow
        Write-Host "PROXY CONFLICT RESOLUTION EXECUTION" -ForegroundColor Yellow
        Write-Host ('=' * 60) -ForegroundColor Yellow
        Write-Host "Items to process: $($itemsToProcess.Count)" -ForegroundColor Cyan
        Write-Host "Audit log: $AuditPath" -ForegroundColor Cyan
        Write-Host "Rollback script: $RollbackScriptPath`n" -ForegroundColor Cyan
    }

    process {
        foreach ($item in $itemsToProcess) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            
            # Use Loser (new) or UserWithError (legacy)
            $targetUser = if ($item.Loser) { $item.Loser } else { $item.UserWithError }
            $targetSAM = if ($item.LoserSAM) { $item.LoserSAM } else { $item.SamAccountName }
            $winnerUser = if ($item.Winner) { $item.Winner } else { $item.ConflictOwner }
            
            Write-Host "`n[$($item.Order)] $($item.ConflictType) - Priority $($item.Priority)" -ForegroundColor Cyan
            Write-Host "    Loser (changing): $targetUser" -ForegroundColor Yellow
            Write-Host "    Winner (keeping): $winnerUser" -ForegroundColor Green
            if ($item.WinnerReason) {
                Write-Host "    Decision: $($item.WinnerReason)" -ForegroundColor DarkGreen
            }
            Write-Host "    Proxy: $($item.ConflictingProxy) -> $($item.NewProxy)" -ForegroundColor White
            if ($item.LoserIsInactive) {
                Write-Host "    [INACTIVE USER]" -ForegroundColor DarkYellow
            }
            Write-Host "    Action: $($item.Action)" -ForegroundColor White
            
            if ($item.Status -eq 'Manual' -or $item.Action -match '^MANUAL:' -or $item.Action -match '^Manual:') {
                Write-Host "    SKIPPED: Requires manual intervention" -ForegroundColor Yellow
                $auditLog.Add([PSCustomObject]@{
                    Timestamp        = $timestamp
                    ConflictId       = $item.ConflictId
                    User             = $targetUser
                    SamAccount       = $targetSAM
                    Winner           = $winnerUser
                    WinnerReason     = $item.WinnerReason
                    ConflictType     = $item.ConflictType
                    ConflictingProxy = $item.ConflictingProxy
                    NewProxy         = $item.NewProxy
                    Suffix           = $item.Suffix
                    Action           = $item.Action
                    Status           = 'Skipped-Manual'
                    Error            = $null
                    ExecutedBy       = $env:USERNAME
                    Computer         = $env:COMPUTERNAME
                    OriginalProxies  = $null
                })
                continue
            }
            
            if (-not $item.ADCommand) {
                Write-Host "    SKIPPED: No AD command defined" -ForegroundColor Yellow
                continue
            }
            
            # Show the command
            Write-Host "`n    --- AD Command ---" -ForegroundColor DarkGray
            $item.ADCommand -split "`n" | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
            Write-Host "    --- End Command ---`n" -ForegroundColor DarkGray
            
            $shouldExecute = $Force -or $PSCmdlet.ShouldProcess(
                "$targetUser - $($item.Action)",
                "Execute AD command"
            )
            
            if ($shouldExecute) {
                try {
                    # Execute the AD command
                    $scriptBlock = [ScriptBlock]::Create($item.ADCommand)
                    
                    Write-Host "    Executing..." -ForegroundColor Yellow
                    & $scriptBlock
                    
                    $item.Status = 'Completed'
                    $item.ExecutedAt = $timestamp
                    $item.ExecutedBy = $env:USERNAME
                    
                    Write-Host "    SUCCESS" -ForegroundColor Green
                    
                    # Add to audit log
                    $auditLog.Add([PSCustomObject]@{
                        Timestamp        = $timestamp
                        ConflictId       = $item.ConflictId
                        User             = $targetUser
                        SamAccount       = $targetSAM
                        Winner           = $winnerUser
                        WinnerReason     = $item.WinnerReason
                        ConflictType     = $item.ConflictType
                        ConflictingProxy = $item.ConflictingProxy
                        NewProxy         = $item.NewProxy
                        Suffix           = $item.Suffix
                        Action           = $item.Action
                        Status           = 'Completed'
                        Error            = $null
                        ExecutedBy       = $env:USERNAME
                        Computer         = $env:COMPUTERNAME
                        OriginalProxies  = if ($item.RollbackData.OriginalProxies) { ($item.RollbackData.OriginalProxies -join '; ') } else { $null }
                    })
                    
                    # Add rollback command
                    if ($item.RollbackData -and $item.RollbackData.SamAccountName) {
                        $rollbackCommands.Add(@"

# Rollback for $targetUser - ConflictId: $($item.ConflictId)
# Winner was: $winnerUser
# Reason: $($item.WinnerReason)
# Executed: $timestamp
`$originalProxies = @(
    $($item.RollbackData.OriginalProxies | ForEach-Object { "    '$_'" } | Out-String)
)
Set-ADUser -Identity '$($item.RollbackData.SamAccountName)' -Replace @{proxyAddresses = `$originalProxies}
Write-Host "Rolled back: $targetUser" -ForegroundColor Green
"@)
                    }
                }
                catch {
                    $item.Status = 'Failed'
                    $item.Error = $_.Exception.Message
                    
                    Write-Host "    FAILED: $($_.Exception.Message)" -ForegroundColor Red
                    
                    $auditLog.Add([PSCustomObject]@{
                        Timestamp        = $timestamp
                        ConflictId       = $item.ConflictId
                        User             = $targetUser
                        SamAccount       = $targetSAM
                        Winner           = $winnerUser
                        WinnerReason     = $item.WinnerReason
                        ConflictType     = $item.ConflictType
                        ConflictingProxy = $item.ConflictingProxy
                        NewProxy         = $item.NewProxy
                        Suffix           = $item.Suffix
                        Action           = $item.Action
                        Status           = 'Failed'
                        Error            = $_.Exception.Message
                        ExecutedBy       = $env:USERNAME
                        Computer         = $env:COMPUTERNAME
                        OriginalProxies  = $null
                    })
                }
            }
            else {
                Write-Host "    SKIPPED by user" -ForegroundColor Yellow
            }
        }
    }

    end {
        # Save audit log
        if ($auditLog.Count -gt 0) {
            $auditLog | Export-Csv -Path $AuditPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nAudit log saved: $AuditPath" -ForegroundColor Green
        }
        
        # Save rollback script
        if ($rollbackCommands.Count -gt 2) {
            $rollbackCommands -join "`n" | Set-Content -Path $RollbackScriptPath -Encoding UTF8
            Write-Host "Rollback script saved: $RollbackScriptPath" -ForegroundColor Green
        }
        
        # Summary
        $completed = @($Plan | Where-Object Status -eq 'Completed').Count
        $failed = @($Plan | Where-Object Status -eq 'Failed').Count
        $pending = @($Plan | Where-Object { $_.Status -eq 'Ready' -or $_.Status -eq 'Pending' }).Count
        $manual = @($Plan | Where-Object Status -eq 'Manual').Count
        
        Write-Host "`n" + ('=' * 40) -ForegroundColor Yellow
        Write-Host "EXECUTION SUMMARY" -ForegroundColor Yellow
        Write-Host ('=' * 40) -ForegroundColor Yellow
        Write-Host "  Completed: $completed" -ForegroundColor Green
        Write-Host "  Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'White' })
        Write-Host "  Pending: $pending" -ForegroundColor $(if ($pending -gt 0) { 'Yellow' } else { 'White' })
        Write-Host "  Manual: $manual" -ForegroundColor $(if ($manual -gt 0) { 'DarkYellow' } else { 'White' })
        
        if ($completed -gt 0) {
            Write-Host "`nNext steps:" -ForegroundColor Cyan
            Write-Host "  1. Trigger Azure AD Connect sync or wait for scheduled sync" -ForegroundColor White
            Write-Host "  2. Verify conflicts resolved in M365 Admin Center" -ForegroundColor White
            Write-Host "  3. If issues, run rollback: $RollbackScriptPath" -ForegroundColor White
        }
        
        return $Plan
    }
}

function Undo-ProxyConflictResolution {
    <#
    .SYNOPSIS
        Rolls back changes using the audit CSV
    .DESCRIPTION
        Reads the audit log and reverts proxyAddresses to original values.
        Can rollback all or specific ConflictIds.
    .EXAMPLE
        Undo-ProxyConflictResolution -AuditPath "C:\temp\resolution-audit.csv" -WhatIf
    .EXAMPLE
        Undo-ProxyConflictResolution -AuditPath "C:\temp\resolution-audit.csv" -ConflictId "abc123"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [string]$AuditPath,

        [Parameter()]
        [string]$ConflictId,

        [Parameter()]
        [switch]$Force
    )

    if (-not (Test-Path $AuditPath)) {
        Write-Error "Audit file not found: $AuditPath"
        return
    }

    $audit = Import-Csv -Path $AuditPath
    $itemsToRollback = if ($ConflictId) {
        $audit | Where-Object { $_.ConflictId -eq $ConflictId -and $_.Status -eq 'Completed' }
    } else {
        $audit | Where-Object Status -eq 'Completed'
    }

    if (-not $itemsToRollback) {
        Write-Warning "No completed items to rollback"
        return
    }

    Write-Host "`n=== ROLLBACK ===" -ForegroundColor Yellow
    Write-Host "Items to rollback: $($itemsToRollback.Count)" -ForegroundColor Cyan

    foreach ($item in $itemsToRollback) {
        if (-not $item.OriginalProxies -or -not $item.SamAccount) {
            Write-Warning "Cannot rollback $($item.User) - missing original data"
            continue
        }

        $originalProxies = $item.OriginalProxies -split '; '
        
        Write-Host "`nRolling back: $($item.User)" -ForegroundColor Cyan
        Write-Host "  SAM: $($item.SamAccount)" -ForegroundColor White
        Write-Host "  Original proxies: $($originalProxies.Count) addresses" -ForegroundColor White

        $shouldRollback = $Force -or $PSCmdlet.ShouldProcess(
            $item.User,
            "Restore original proxyAddresses"
        )

        if ($shouldRollback) {
            try {
                Set-ADUser -Identity $item.SamAccount -Replace @{proxyAddresses = $originalProxies}
                Write-Host "  SUCCESS" -ForegroundColor Green
            }
            catch {
                Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

#endregion Conflict Resolution Pipeline

#region Email Notification

function Protect-MultiKeyCredentialJson {
    <#
    .SYNOPSIS
        Creates obfuscated credential JSON with multiple keys (Graph + SMTP)
    .DESCRIPTION
        Extends Protect-CredentialJson to support multiple app registrations:
        - Main app for Graph read operations
        - SMTP app for Mail.Send operations
        Each secret has its own obfuscation key for separate channel distribution.
    .EXAMPLE
        Protect-MultiKeyCredentialJson -TenantId "xxx" `
            -ClientId "graph-app-id" -ClientSecret "graph-secret" -Key "ELKW_L2-Nov2025" `
            -SMTPClientId "smtp-app-id" -SMTPClientSecret "smtp-secret" -SMTPKey "ELKW_SMTP-Nov2025" `
            -OutputPath "C:\temp\multi-creds.json"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$ClientSecret,

        [Parameter(Mandatory)]
        [string]$Key,

        [Parameter()]
        [string]$SMTPClientId,

        [Parameter()]
        [string]$SMTPClientSecret,

        [Parameter()]
        [string]$SMTPKey,

        [Parameter()]
        [string]$OutputPath
    )

    # XOR obfuscation function
    function Invoke-XorObfuscate {
        param([string]$Text, [string]$KeyText)
        $textBytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($KeyText)
        $result = for ($i = 0; $i -lt $textBytes.Length; $i++) {
            $textBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        return [Convert]::ToBase64String([byte[]]$result)
    }

    $json = @{
        TenantId           = $TenantId
        ClientId           = $ClientId
        ClientSecretEnc    = Invoke-XorObfuscate -Text $ClientSecret -KeyText $Key
        KeyHint            = "Key provided via separate channel"
    }

    if ($SMTPClientId -and $SMTPClientSecret -and $SMTPKey) {
        $json['SMTPClientId'] = $SMTPClientId
        $json['SMTPClientSecretEnc'] = Invoke-XorObfuscate -Text $SMTPClientSecret -KeyText $SMTPKey
        $json['SMTPKeyHint'] = "SMTP key provided via separate channel"
    }

    $jsonText = $json | ConvertTo-Json -Depth 5

    if ($OutputPath) {
        $jsonText | Set-Content -Path $OutputPath -Encoding UTF8
        Write-Host "Multi-key credential file saved: $OutputPath" -ForegroundColor Green
        Write-Host "`nDistribution:" -ForegroundColor Yellow
        Write-Host "  - JSON file via Teams/Email" -ForegroundColor White
        Write-Host "  - Graph Key via phone/Signal: $Key" -ForegroundColor White
        if ($SMTPKey) {
            Write-Host "  - SMTP Key via separate channel: $SMTPKey" -ForegroundColor White
        }
    }

    return $jsonText
}

function Unprotect-MultiKeyCredentialJson {
    <#
    .SYNOPSIS
        Decodes multi-key credential JSON
    .DESCRIPTION
        Decodes the obfuscated credentials using provided keys.
        Returns hashtable with Graph and SMTP credentials.
    .EXAMPLE
        $creds = Unprotect-MultiKeyCredentialJson -Path "C:\temp\multi-creds.json" -Key "ELKW_L2-Nov2025" -SMTPKey "ELKW_SMTP-Nov2025"
        $graphToken = Get-GraphToken -TenantId $creds.TenantId -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
        $smtpToken = Get-GraphToken -TenantId $creds.TenantId -ClientId $creds.SMTPClientId -ClientSecret $creds.SMTPClientSecret
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Key,

        [Parameter()]
        [string]$SMTPKey
    )

    function Invoke-XorDeobfuscate {
        param([string]$EncodedText, [string]$KeyText)
        $encBytes = [Convert]::FromBase64String($EncodedText)
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($KeyText)
        $result = for ($i = 0; $i -lt $encBytes.Length; $i++) {
            $encBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        return [System.Text.Encoding]::UTF8.GetString([byte[]]$result)
    }

    $json = Get-Content -Path $Path -Raw | ConvertFrom-Json

    $result = @{
        TenantId     = $json.TenantId
        ClientId     = $json.ClientId
        ClientSecret = Invoke-XorDeobfuscate -EncodedText $json.ClientSecretEnc -KeyText $Key
    }

    if ($json.SMTPClientId -and $SMTPKey) {
        $result['SMTPClientId'] = $json.SMTPClientId
        $result['SMTPClientSecret'] = Invoke-XorDeobfuscate -EncodedText $json.SMTPClientSecretEnc -KeyText $SMTPKey
    }

    return $result
}

function Send-ConflictResolutionReport {
    <#
    .SYNOPSIS
        Sends email report of conflict resolution results
    .DESCRIPTION
        Uses Graph API with Mail.Send permission to send summary email.
        Includes resolution statistics, audit log attachment, and rollback info.
    .EXAMPLE
        Send-ConflictResolutionReport -SMTPAccessToken $smtpToken `
            -From "noreply@contoso.com" `
            -To "admin@contoso.com" `
            -AuditPath "C:\temp\audit.csv" `
            -Plan $plan
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SMTPAccessToken,

        [Parameter(Mandatory)]
        [string]$From,

        [Parameter(Mandatory)]
        [string[]]$To,

        [Parameter()]
        [string[]]$Cc,

        [Parameter(Mandatory)]
        [string]$AuditPath,

        [Parameter(Mandatory)]
        [PSObject[]]$Plan,

        [Parameter()]
        [string]$Subject = "Proxy Conflict Resolution Report - $(Get-Date -Format 'yyyy-MM-dd')"
    )

    # Build statistics
    $completed = @($Plan | Where-Object Status -eq 'Completed').Count
    $failed = @($Plan | Where-Object Status -eq 'Failed').Count
    $pending = @($Plan | Where-Object { $_.Status -eq 'Ready' -or $_.Status -eq 'Pending' }).Count
    $manual = @($Plan | Where-Object Status -eq 'Manual').Count
    
    # Group by decision reason
    $reasonStatsHtml = ''
    $Plan | Where-Object { $_.WinnerReason } | Group-Object WinnerReason | ForEach-Object {
        $reasonStatsHtml += "<li><strong>$($_.Name):</strong> $($_.Count)</li>"
    }
    
    # Group by suffix
    $suffixStatsHtml = ''
    $Plan | Where-Object { $_.Suffix } | Group-Object Suffix | ForEach-Object {
        $suffixStatsHtml += "<li><code>$($_.Name)</code>: $($_.Count)</li>"
    }
    
    # Build table rows
    $tableRows = ''
    foreach ($item in $Plan) {
        $statusClass = switch ($item.Status) {
            'Completed' { 'status-completed' }
            'Failed' { 'status-failed' }
            'Manual' { 'status-manual' }
            default { 'status-pending' }
        }
        $loserDisplay = if ($item.Loser) { $item.Loser } else { $item.UserWithError }
        $winnerDisplay = if ($item.Winner) { $item.Winner } else { $item.ConflictOwner }
        $inactiveBadge = if ($item.LoserIsInactive) { "<span class='inactive-badge'>INACTIVE</span>" } else { '' }
        $suffixDisplay = if ($item.Suffix) { "<span class='suffix'>$($item.Suffix)</span>" } else { '' }
        $newProxyDisplay = if ($item.NewProxy) { "-> $($item.NewProxy)" } else { '' }
        $loserSam = if ($item.LoserSAM) { $item.LoserSAM } else { $item.SamAccountName }
        
        $tableRows += @"
<tr>
    <td class='winner'><strong>$winnerDisplay</strong><br><small>$($item.ConflictType)</small></td>
    <td class='loser'>$loserDisplay $inactiveBadge<br><small>SAM: $loserSam</small></td>
    <td><code>$($item.ConflictingProxy)</code><br>$newProxyDisplay $suffixDisplay</td>
    <td class='$statusClass'>$($item.Status)</td>
</tr>
"@
    }
    
    # Build failed rows if any
    $failedHtml = ''
    if ($failed -gt 0) {
        $failedRows = ''
        foreach ($item in ($Plan | Where-Object Status -eq 'Failed')) {
            $userDisplay = if ($item.Loser) { $item.Loser } else { $item.UserWithError }
            $failedRows += "<tr><td>$userDisplay</td><td>$($item.Error)</td></tr>"
        }
        $failedHtml = @"
<h2>Warning: Failed Resolutions</h2>
<table>
<tr><th>User</th><th>Error</th></tr>
$failedRows
</table>
"@
    }
    
    # Build reason box
    $reasonBoxHtml = ''
    if ($reasonStatsHtml) {
        $reasonBoxHtml = @"
<div class="reason-box">
<h3>Decision Logic Applied</h3>
<ul>$reasonStatsHtml</ul>
</div>
"@
    }
    
    # Build suffix box
    $suffixBoxHtml = ''
    if ($suffixStatsHtml) {
        $suffixBoxHtml = @"
<div class="reason-box">
<h3>Suffixes Applied</h3>
<ul>$suffixStatsHtml</ul>
</div>
"@
    }

    # Build HTML body
    $body = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f9f9f9; }
    .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
    h2 { color: #333; margin-top: 30px; }
    h3 { color: #555; margin-top: 20px; font-size: 14px; }
    .stats { display: flex; gap: 15px; margin: 20px 0; flex-wrap: wrap; }
    .stat-box { padding: 15px 25px; border-radius: 8px; text-align: center; min-width: 100px; }
    .stat-completed { background: #dff6dd; color: #107c10; }
    .stat-failed { background: #fde7e9; color: #d13438; }
    .stat-pending { background: #fff4ce; color: #8a6914; }
    .stat-manual { background: #e6e6e6; color: #333; }
    .stat-number { font-size: 28px; font-weight: bold; }
    .stat-label { font-size: 12px; text-transform: uppercase; }
    table { border-collapse: collapse; width: 100%; margin-top: 15px; font-size: 13px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background: #f5f5f5; font-weight: 600; }
    tr:nth-child(even) { background: #fafafa; }
    .status-completed { color: #107c10; font-weight: bold; }
    .status-failed { color: #d13438; font-weight: bold; }
    .status-pending { color: #8a6914; }
    .status-manual { color: #666; font-style: italic; }
    .winner { color: #107c10; }
    .loser { color: #8a6914; }
    .suffix { background: #e8f4fd; padding: 2px 6px; border-radius: 4px; font-family: monospace; font-size: 11px; }
    .inactive-badge { background: #fff4ce; color: #8a6914; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 5px; }
    .reason-box { background: #f0f0f0; padding: 15px; border-radius: 6px; margin: 15px 0; }
    .reason-box ul { margin: 10px 0; padding-left: 20px; }
    .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; color: #666; font-size: 12px; }
    code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
</style>
</head>
<body>
<div class="container">
<h1>Proxy Conflict Resolution Report</h1>

<p>Resolution executed on <strong>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</strong> by <strong>$env:USERNAME</strong> on <strong>$env:COMPUTERNAME</strong></p>

<div class="stats">
    <div class="stat-box stat-completed">
        <div class="stat-number">$completed</div>
        <div class="stat-label">Completed</div>
    </div>
    <div class="stat-box stat-failed">
        <div class="stat-number">$failed</div>
        <div class="stat-label">Failed</div>
    </div>
    <div class="stat-box stat-pending">
        <div class="stat-number">$pending</div>
        <div class="stat-label">Pending</div>
    </div>
    <div class="stat-box stat-manual">
        <div class="stat-number">$manual</div>
        <div class="stat-label">Manual</div>
    </div>
</div>

$reasonBoxHtml

$suffixBoxHtml

<h2>Resolution Details</h2>
<table>
<tr>
    <th style="width:25%">Winner (keeps address)</th>
    <th style="width:25%">Loser (changes)</th>
    <th>Proxy Change</th>
    <th style="width:15%">Status</th>
</tr>
$tableRows
</table>

$failedHtml

<div class="footer">
<p><strong>Audit Log:</strong> See attached CSV for full details including rollback data.</p>
<p><strong>Next Steps:</strong></p>
<ul>
    <li>Wait for Azure AD Connect sync (default: 30 minutes)</li>
    <li>Verify license assignment errors are resolved in M365 Admin Center</li>
    <li>For rollback, use the attached audit CSV with <code>Undo-ProxyConflictResolution</code></li>
</ul>
<p><strong>Decision Logic:</strong></p>
<ol>
    <li><strong>Active vs Inactive:</strong> Active users win over ausgeschieden/deaktiviert</li>
    <li><strong>Eineindeutigkeit:</strong> User whose SAM matches email prefix wins</li>
    <li><strong>Seniority:</strong> Older user (by sync date) wins ties</li>
    <li><strong>Existing owner:</strong> Current owner wins final ties</li>
</ol>
</div>
</div>
</body>
</html>
"@

    # Read and encode audit file
    $auditContent = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($AuditPath))
    $auditFileName = [System.IO.Path]::GetFileName($AuditPath)

    # Build message
    $message = @{
        message = @{
            subject = $Subject
            body    = @{
                contentType = 'HTML'
                content     = $body
            }
            toRecipients = @(
                foreach ($recipient in $To) {
                    @{ emailAddress = @{ address = $recipient } }
                }
            )
            attachments = @(
                @{
                    '@odata.type'  = '#microsoft.graph.fileAttachment'
                    name           = $auditFileName
                    contentType    = 'text/csv'
                    contentBytes   = $auditContent
                }
            )
        }
        saveToSentItems = $true
    }

    if ($Cc) {
        $message.message['ccRecipients'] = @(
            foreach ($recipient in $Cc) {
                @{ emailAddress = @{ address = $recipient } }
            }
        )
    }

    # Send via Graph API
    $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    $headers = @{
        Authorization  = "Bearer $SMTPAccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body ($message | ConvertTo-Json -Depth 10)
        Write-Host "Email sent successfully to: $($To -join ', ')" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
        if ($_.ErrorDetails.Message) {
            $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            Write-Error "Detail: $($errDetail.error.message)"
        }
    }
}

#endregion Email Notification

#region Module Export / Dot-Source Safety

$ExportFunctions = @(
    # License/Provisioning Errors
    'Get-GraphLicenseErrors',
    'Get-GraphSkuFriendlyName',
    'Get-GraphProvisioningErrors',
    'Find-DuplicateProxyAddresses',
    'Get-GraphUserConflicts',
    'Get-GraphGroupLicenseErrors',
    'Get-GraphGroupsWithLicenseErrors',
    'Get-GraphLicenseErrorReport',
    
    # Conflict Resolution
    'Resolve-GraphProxyConflict',
    'Search-GraphProxyAddress',
    'Get-ProxyConflictAnalysis',
    'New-ProxyConflictResolutionPlan',
    'Invoke-ProxyConflictResolution',
    'Undo-ProxyConflictResolution',
    
    # Credentials
    'Protect-CredentialJson',
    'Unprotect-CredentialJson',
    'Protect-MultiKeyCredentialJson',
    'Unprotect-MultiKeyCredentialJson',
    'Get-GraphToken',
    
    # Email Notification
    'Send-ConflictResolutionReport'
)

# Detect execution context and export appropriately
if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    # Running as Import-Module
    Export-ModuleMember -Function $ExportFunctions
}
elseif ($MyInvocation.InvocationName -eq '.' -or $MyInvocation.Line -match '^\.\s') {
    # Dot-sourced: functions already in caller's scope, nothing to do
    Write-Verbose "Dot-sourced: Functions available in current scope"
}
else {
    # Direct execution (F5 / Run Script / .\script.ps1)
    # Make functions global so they persist
    foreach ($fn in $ExportFunctions) {
        if (Get-Command -Name $fn -ErrorAction SilentlyContinue) {
            Set-Item -Path "Function:Global:$fn" -Value (Get-Command $fn).ScriptBlock
        }
    }
    Write-Host "Functions exported to global scope: $($ExportFunctions -join ', ')" -ForegroundColor Green
}

#endregion
