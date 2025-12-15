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
                    if ($response.value) { $users.AddRange($response.value) }
                    $uri = $response.'@odata.nextLink'
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
            Authorization  = "Bearer $AccessToken"
            'Content-Type' = 'application/json'
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
                if ($ErrorsOnly) {
                    # Filter for users WITH provisioning errors only
                    $uri = "$baseUri/users?`$select=$selectProps&`$filter=onPremisesProvisioningErrors/any(x:x/category ne null)&`$top=$Top"
                }
                else {
                    $uri = "$baseUri/users?`$select=$selectProps&`$top=$Top"
                }
                $users = [System.Collections.Generic.List[PSObject]]::new()

                do {
                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
                    if ($response.value) { $users.AddRange($response.value) }
                    $uri = $response.'@odata.nextLink'
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
                if ($response.value) { $users.AddRange($response.value) }
                $uri = $response.'@odata.nextLink'
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
        $token = Get-GraphToken -TenantId $creds.TenantId -ClientId $creds.ClientId -ClientSecret $creds.ClientSecret
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

    return [PSCustomObject]@{
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

Export-ModuleMember -Function @(
    'Get-GraphLicenseErrors',
    'Get-GraphSkuFriendlyName',
    'Get-GraphProvisioningErrors',
    'Find-DuplicateProxyAddresses',
    'Get-GraphUserConflicts',
    'Protect-CredentialJson',
    'Unprotect-CredentialJson',
    'Get-GraphToken'
)
