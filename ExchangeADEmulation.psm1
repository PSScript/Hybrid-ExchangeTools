# === CORE CMDLET SKELETONS ===
# ExchangeADEmulation.psm1

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

# ============================================================================
# MODULE INITIALIZATION
# ============================================================================

# Emulation mode guard
$script:ExchangeEmulationMode = $true

# Load constants
$script:Constants = @{
    ProvisionMailbox = 1
    ProvisionArchive = 2
    Migrated = 4
    RoomMailbox = 32
    EquipmentMailbox = 64
    SharedMask = 96
    
    LegalComposites = @(1, 3, 4, 6, 33, 36, 38, 65, 68, 70, 100, 102)
    
    DisplayTypes = @{
        NonACLable = -2147483642
        ACLable = 1073741824
        Room = 7
        Equipment = 8
    }
    
    RecipientTypeDetails = @{
        User = [long]2147483648
        Shared = [long]34359738368
        Room = [long]8589934592
        Equipment = [long]17179869184
    }
    
    RoutingAddressPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.mail\.onmicrosoft\.com$'
}

# Seal constants
Set-Variable -Name Constants -Option ReadOnly -Scope Script -Force

# ============================================================================
# SELF-TEST ON IMPORT
# ============================================================================

function Test-ModuleIntegrity {
    [CmdletBinding()]
    param()
    
    Write-Verbose "Running module integrity self-test..."
    
    # Test 1: 97 must NEVER exist in legal composites
    if (97 -in $script:Constants.LegalComposites) {
        throw "CRITICAL: RRT=97 found in legal composites. This violates spec."
    }
    
    # Test 2: All legal composites must be positive
    foreach ($rrt in $script:Constants.LegalComposites) {
        if ($rrt -le 0) {
            throw "CRITICAL: Invalid RRT=$rrt in legal composites."
        }
    }
    
    # Test 3: Display types must be non-zero
    foreach ($dt in $script:Constants.DisplayTypes.Values) {
        if ($dt -eq 0) {
            throw "CRITICAL: DisplayType=0 detected. This is illegal."
        }
    }
    
    Write-Verbose "Module integrity verified: All guards passed."
}

Test-ModuleIntegrity

# ============================================================================
# PRIVATE FUNCTIONS
# ============================================================================

function Assert-EmulationMode {
    [CmdletBinding()]
    param()
    
    if (-not $script:ExchangeEmulationMode) {
        $ex = [System.InvalidOperationException]::new(
            "Exchange emulation mode is disabled. Set `$script:ExchangeEmulationMode = `$true"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'EmulationModeDisabled',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $null
            )
        )
    }
}

function Get-RecipientTypeComposite {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Shared', 'Room', 'Equipment')]
        [string]$Type,
        
        [Parameter(Mandatory)]
        [ValidateSet('Provision', 'Migrated')]
        [string]$Mode,
        
        [switch]$Archive
    )
    
    $base = switch ($Mode) {
        'Provision' {
            switch ($Type) {
                'User' { $script:Constants.ProvisionMailbox }
                'Shared' { 100 }  # ALWAYS 100 for Shared (NEVER 97)
                'Room' { $script:Constants.ProvisionMailbox + $script:Constants.RoomMailbox }
                'Equipment' { $script:Constants.ProvisionMailbox + $script:Constants.EquipmentMailbox }
            }
        }
        'Migrated' {
            switch ($Type) {
                'User' { $script:Constants.Migrated }
                'Shared' { 100 }  # ALWAYS 100 for Shared
                'Room' { $script:Constants.Migrated + $script:Constants.RoomMailbox }
                'Equipment' { $script:Constants.Migrated + $script:Constants.EquipmentMailbox }
            }
        }
    }
    
    if ($Archive) {
        $base += $script:Constants.ProvisionArchive
    }
    
    # CRITICAL GUARD: Verify result is legal
    if ($base -notin $script:Constants.LegalComposites) {
        throw [System.InvalidOperationException]::new(
            "FATAL: Computed illegal RRT=$base for Type=$Type Mode=$Mode Archive=$Archive"
        )
    }
    
    return $base
}

function Get-RecipientDisplayType {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Shared', 'Room', 'Equipment')]
        [string]$Type,
        
        [switch]$ACLableSyncedObjectEnabled
    )
    
    if ($Type -eq 'Room') {
        return $script:Constants.DisplayTypes.Room
    }
    
    if ($Type -eq 'Equipment') {
        return $script:Constants.DisplayTypes.Equipment
    }
    
    # User or Shared
    if ($ACLableSyncedObjectEnabled) {
        return $script:Constants.DisplayTypes.ACLable
    }
    else {
        return $script:Constants.DisplayTypes.NonACLable
    }
}

function Invoke-EmailNormalization {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [string]$PrimarySmtpAddress,
        [string[]]$ExistingProxyAddresses = @(),
        [string[]]$AddEmailAddresses = @(),
        [string[]]$RemoveEmailAddresses = @(),
        [string]$OldPrimarySmtpAddress,
        [string]$RoutingAddress
    )
    
    # Use HashSet for case-insensitive deduplication
    $proxySet = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
    
    # Load existing addresses
    foreach ($addr in $ExistingProxyAddresses) {
        if ($addr) {
            [void]$proxySet.Add($addr)
        }
    }
    
    # Handle primary SMTP change
    if ($PrimarySmtpAddress) {
        if ($OldPrimarySmtpAddress -and $PrimarySmtpAddress -ne $OldPrimarySmtpAddress) {
            # Demote old primary to alias
            [void]$proxySet.Remove("SMTP:$OldPrimarySmtpAddress")
            [void]$proxySet.Add("smtp:$OldPrimarySmtpAddress")
        }
        
        # Remove any existing variants of new primary
        $toRemove = $proxySet | Where-Object { 
            $_ -imatch "^SMTP:$([regex]::Escape($PrimarySmtpAddress))$" 
        }
        foreach ($r in $toRemove) {
            [void]$proxySet.Remove($r)
        }
        
        # Set new primary
        [void]$proxySet.Add("SMTP:$PrimarySmtpAddress")
    }
    
    # Add new addresses
    foreach ($addr in $AddEmailAddresses) {
        if ($addr) {
            $normalized = if ($addr -notmatch '^smtp:') { "smtp:$addr" } else { $addr }
            [void]$proxySet.Add($normalized)
        }
    }
    
    # Remove addresses
    foreach ($addr in $RemoveEmailAddresses) {
        if ($addr) {
            # Check if removing current primary without replacement
            if ($addr -ieq $PrimarySmtpAddress) {
                throw [System.ArgumentException]::new(
                    "Cannot remove current primary SMTP address without providing a new one"
                )
            }
            
            $toRemove = $proxySet | Where-Object { 
                $_ -imatch "^smtp:$([regex]::Escape($addr))$" 
            }
            foreach ($r in $toRemove) {
                [void]$proxySet.Remove($r)
            }
        }
    }
    
    # Add routing address as alias
    if ($RoutingAddress) {
        $routingAlias = if ($RoutingAddress -match '^smtp:') { 
            $RoutingAddress 
        } else { 
            "smtp:$RoutingAddress" 
        }
        [void]$proxySet.Add($routingAlias)
    }
    
    return $proxySet.ToArray()
}

function Invoke-IdempotencyCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,
        
        [Parameter(Mandatory)]
        [hashtable]$Desired
    )
    
    $properties = ($Desired.Keys + @('proxyAddresses', 'mail')) | Select-Object -Unique
    
    try {
        $current = Get-ADUser -Identity $Identity -Properties $properties -ErrorAction Stop
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
    
    $delta = @{}
    
    foreach ($key in $Desired.Keys) {
        $oldValue = $current.$key
        $newValue = $Desired[$key]
        
        if ($key -eq 'proxyAddresses') {
            # Normalize and compare as sets
            $oldNorm = ($oldValue | ForEach-Object { $_.ToLower() }) | Sort-Object -Unique
            $newNorm = ($newValue | ForEach-Object { $_.ToLower() }) | Sort-Object -Unique
            
            $diff = Compare-Object -ReferenceObject $oldNorm -DifferenceObject $newNorm
            if ($diff) {
                $delta[$key] = $newValue
            }
        }
        else {
            # Case-sensitive comparison for scalar attributes
            if ($oldValue -cne $newValue) {
                $delta[$key] = $newValue
            }
        }
    }
    
    return @{
        HasChanges = ($delta.Count -gt 0)
        Delta = $delta
        Before = $current | Select-Object -Property $properties
    }
}

function Test-AttributeAllowlist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Attributes
    )
    
    $allowlistPath = "$PSScriptRoot/Data/schema_allowlist.json"
    
    if (-not (Test-Path $allowlistPath)) {
        throw [System.IO.FileNotFoundException]::new(
            "Allowlist file not found: $allowlistPath"
        )
    }
    
    $allowlist = Get-Content $allowlistPath | ConvertFrom-Json
    $illegal = $Attributes.Keys | Where-Object { $_ -notin $allowlist.write }
    
    if ($illegal) {
        $ex = [System.ArgumentException]::new(
            "Illegal attributes: $($illegal -join ', '). Allowlist: $($allowlist.write -join ', ')"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'IllegalAttribute',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $Attributes
            )
        )
    }
}

function Test-EmailFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )
    
    if ($Address -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        $ex = [System.FormatException]::new(
            "Invalid email format: $Address"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'InvalidFormat',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Address
            )
        )
    }
}

function Test-RoutingAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Address,
        
        [string]$PrimarySmtpAddress
    )
    
    if ($Address -notmatch $script:Constants.RoutingAddressPattern) {
        $ex = [System.ArgumentException]::new(
            "RemoteRoutingAddress must be *.mail.onmicrosoft.com format: $Address"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'InvalidRoutingDomain',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Address
            )
        )
    }
    
    if ($PrimarySmtpAddress -and ($Address -ieq $PrimarySmtpAddress)) {
        $ex = [System.ArgumentException]::new(
            "RemoteRoutingAddress cannot equal PrimarySmtpAddress"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'RoutingEqualsPrimary',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Address
            )
        )
    }
}

function Test-RecipientType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$Value
    )
    
    if ($Value -notin $script:Constants.LegalComposites) {
        $ex = [System.InvalidOperationException]::new(
            "Illegal msExchRemoteRecipientType=$Value. Legal values: $($script:Constants.LegalComposites -join ', ')"
        )
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                $ex,
                'IllegalComposite',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $Value
            )
        )
    }
}

function Write-ExchangeAudit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,
        
        [Parameter(Mandatory)]
        [hashtable]$Before,
        
        [Parameter(Mandatory)]
        [hashtable]$After,
        
        [Parameter(Mandatory)]
        [string]$Cmdlet,
        
        [Parameter(Mandatory)]
        [hashtable]$Parameters,
        
        [Parameter(Mandatory)]
        [ValidateSet('Success', 'Failed', 'NoChange')]
        [string]$Status
    )
    
    $logDir = "$PSScriptRoot/Logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    $logFile = "$logDir/audit-$(Get-Date -Format 'yyyy-MM-dd').jsonl"
    
    $auditEntry = @{
        timestamp = (Get-Date).ToUniversalTime().ToString('o')
        identity = $Identity
        cmdlet = $Cmdlet
        parameters = $Parameters
        before = $Before
        after = $After
        status = $Status
        user = $env:USERNAME
        computer = $env:COMPUTERNAME
    } | ConvertTo-Json -Compress
    
    Add-Content -Path $logFile -Value $auditEntry -Encoding UTF8
}

# ============================================================================
# PUBLIC CMDLETS
# ============================================================================

function New-RemoteMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'User')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Name,
        
        [Parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$UserPrincipalName,
        
        [string]$Alias,
        
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$PrimarySmtpAddress,
        
        [Parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.mail\.onmicrosoft\.com$')]
        [string]$RemoteRoutingAddress,
        
        [Parameter(ParameterSetName = 'Shared')]
        [switch]$Shared,
        
        [Parameter(ParameterSetName = 'Room')]
        [switch]$Room,
        
        [Parameter(ParameterSetName = 'Equipment')]
        [switch]$Equipment,
        
        [switch]$Archive,
        
        [switch]$ACLableSyncedObjectEnabled,
        
        [string]$FirstName,
        [string]$Initials,
        [string]$LastName,
        [string]$OrganizationalUnit,
        [string]$SamAccountName
    )
    
    begin {
        Assert-EmulationMode
    }
    
    process {
        # Determine mailbox type
        $type = if ($Shared) { 'Shared' }
                elseif ($Room) { 'Room' }
                elseif ($Equipment) { 'Equipment' }
                else { 'User' }
        
        # Check if user already exists
        $existing = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
        if ($existing) {
            $ex = [System.InvalidOperationException]::new(
                "User already exists: $UserPrincipalName"
            )
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    $ex,
                    'ResourceExists',
                    [System.Management.Automation.ErrorCategory]::ResourceExists,
                    $UserPrincipalName
                )
            )
        }
        
        # Set defaults
        if (-not $Alias) {
            $Alias = $UserPrincipalName.Split('@')[0]
        }
        if (-not $PrimarySmtpAddress) {
            $PrimarySmtpAddress = $UserPrincipalName
        }
        if (-not $SamAccountName) {
            $SamAccountName = $Alias
        }
        
        # Calculate recipient type
        $rrt = Get-RecipientTypeComposite -Type $type -Mode Provision -Archive:$Archive
        $displayType = Get-RecipientDisplayType -Type $type -ACLableSyncedObjectEnabled:$ACLableSyncedObjectEnabled
        $typeDetails = $script:Constants.RecipientTypeDetails[$type]
        
        # Normalize email addresses
        $proxyAddresses = Invoke-EmailNormalization `
            -PrimarySmtpAddress $PrimarySmtpAddress `
            -RoutingAddress $RemoteRoutingAddress
        
        # Build attribute bag
        $attributes = @{
            msExchRemoteRecipientType = $rrt
            msExchRecipientDisplayType = $displayType
            msExchRecipientTypeDetails = $typeDetails
            targetAddress = "smtp:$RemoteRoutingAddress"
            proxyAddresses = $proxyAddresses
            mailNickname = $Alias
            mail = $PrimarySmtpAddress
        }
        
        # Validate allowlist
        Test-AttributeAllowlist -Attributes $attributes
        
        # Create AD user
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create AD user and enable remote mailbox")) {
            $adParams = @{
                Name = $Name
                UserPrincipalName = $UserPrincipalName
                SamAccountName = $SamAccountName
                Enabled = $true
            }
            
            if ($FirstName) { $adParams.GivenName = $FirstName }
            if ($LastName) { $adParams.Surname = $LastName }
            if ($Initials) { $adParams.Initials = $Initials }
            if ($OrganizationalUnit) { $adParams.Path = $OrganizationalUnit }
            
            $user = New-ADUser @adParams -PassThru
            
            # Stamp Exchange attributes
            Set-ADUser -Identity $user.DistinguishedName -Replace $attributes
            
            # Audit
            Write-ExchangeAudit `
                -Identity $UserPrincipalName `
                -Before @{} `
                -After $attributes `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'Success'
            
            # Return result
            [PSCustomObject]@{
                Identity = $UserPrincipalName
                Operation = 'New'
                Status = 'Changed'
                RemoteRecipientType = $rrt
                RecipientDisplayType = $displayType
                RecipientTypeDetails = $typeDetails
                PrimarySmtpAddress = $PrimarySmtpAddress
                RemoteRoutingAddress = $RemoteRoutingAddress
                Delta = $attributes
            }
        }
    }
}

function Enable-RemoteMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'User')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name')]
        [string]$Identity,
        
        [string]$Alias,
        
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$PrimarySmtpAddress,
        
        [Parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.mail\.onmicrosoft\.com$')]
        [string]$RemoteRoutingAddress,
        
        [Parameter(ParameterSetName = 'Shared')]
        [switch]$Shared,
        
        [Parameter(ParameterSetName = 'Room')]
        [switch]$Room,
        
        [Parameter(ParameterSetName = 'Equipment')]
        [switch]$Equipment,
        
        [switch]$Archive,
        
        [switch]$ACLableSyncedObjectEnabled
    )
    
    begin {
        Assert-EmulationMode
    }
    
    process {
        # Get existing user
        try {
            $user = Get-ADUser -Identity $Identity -Properties proxyAddresses, mail, mailNickname -ErrorAction Stop
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        
        # Determine mailbox type
        $type = if ($Shared) { 'Shared' }
                elseif ($Room) { 'Room' }
                elseif ($Equipment) { 'Equipment' }
                else { 'User' }
        
        # Set defaults
        if (-not $Alias) {
            $Alias = $user.SamAccountName
        }
        if (-not $PrimarySmtpAddress) {
            $PrimarySmtpAddress = $user.UserPrincipalName
        }
        
        # Calculate recipient type
        $rrt = Get-RecipientTypeComposite -Type $type -Mode Provision -Archive:$Archive
        $displayType = Get-RecipientDisplayType -Type $type -ACLableSyncedObjectEnabled:$ACLableSyncedObjectEnabled
        $typeDetails = $script:Constants.RecipientTypeDetails[$type]
        
        # Normalize email addresses
        $proxyAddresses = Invoke-EmailNormalization `
            -PrimarySmtpAddress $PrimarySmtpAddress `
            -ExistingProxyAddresses $user.proxyAddresses `
            -RoutingAddress $RemoteRoutingAddress
        
        # Build desired state
        $desired = @{
            msExchRemoteRecipientType = $rrt
            msExchRecipientDisplayType = $displayType
            msExchRecipientTypeDetails = $typeDetails
            targetAddress = "smtp:$RemoteRoutingAddress"
            proxyAddresses = $proxyAddresses
            mailNickname = $Alias
            mail = $PrimarySmtpAddress
        }
        
        # Validate allowlist
        Test-AttributeAllowlist -Attributes $desired
        
        # Idempotency check
        $check = Invoke-IdempotencyCheck -Identity $Identity -Desired $desired
        
        if (-not $check.HasChanges) {
            Write-ExchangeAudit `
                -Identity $Identity `
                -Before $check.Before `
                -After @{} `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'NoChange'
            
            return [PSCustomObject]@{
                Identity = $Identity
                Operation = 'Enable'
                Status = 'NoChange'
                RemoteRecipientType = $rrt
                RecipientDisplayType = $displayType
                Delta = @{}
            }
        }
        
        # Apply changes
        if ($PSCmdlet.ShouldProcess($Identity, "Enable remote mailbox")) {
            Set-ADUser -Identity $Identity -Replace $check.Delta
            
            Write-ExchangeAudit `
                -Identity $Identity `
                -Before $check.Before `
                -After $check.Delta `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'Success'
            
            [PSCustomObject]@{
                Identity = $Identity
                Operation = 'Enable'
                Status = 'Changed'
                RemoteRecipientType = $rrt
                RecipientDisplayType = $displayType
                RecipientTypeDetails = $typeDetails
                PrimarySmtpAddress = $PrimarySmtpAddress
                RemoteRoutingAddress = $RemoteRoutingAddress
                Delta = $check.Delta
            }
        }
    }
}

function Set-RemoteMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Identity,
        
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$PrimarySmtpAddress,
        
        [string[]]$AddEmailAddresses,
        
        [string[]]$RemoveEmailAddresses,
        
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.mail\.onmicrosoft\.com$')]
        [string]$RemoteRoutingAddress,
        
        [ValidateSet('User', 'Shared', 'Room', 'Equipment')]
        [string]$Type,
        
        [bool]$ArchiveEnabled,
        
        [bool]$ACLableSyncedObjectEnabled,
        
        [string]$CustomAttribute1,
        [string]$CustomAttribute2,
        [string]$CustomAttribute3,
        [string]$CustomAttribute4,
        [string]$CustomAttribute5,
        [string]$CustomAttribute6,
        [string]$CustomAttribute7,
        [string]$CustomAttribute8,
        [string]$CustomAttribute9,
        [string]$CustomAttribute10,
        [string]$CustomAttribute11,
        [string]$CustomAttribute12,
        [string]$CustomAttribute13,
        [string]$CustomAttribute14,
        [string]$CustomAttribute15
    )
    
    begin {
        Assert-EmulationMode
    }
    
    process {
        # Get current state
        try {
            $user = Get-ADUser -Identity $Identity -Properties `
                msExchRemoteRecipientType, msExchRecipientDisplayType, msExchRecipientTypeDetails, `
                proxyAddresses, mail, targetAddress, mailNickname, `
                extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, `
                extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, `
                extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 `
                -ErrorAction Stop
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        
        $desired = @{}
        
        # Handle type conversion
        if ($Type) {
            $rrt = Get-RecipientTypeComposite -Type $Type -Mode Migrated
            
            # Handle archive bit
            if ($PSBoundParameters.ContainsKey('ArchiveEnabled')) {
                if ($ArchiveEnabled -and -not ($rrt -band $script:Constants.ProvisionArchive)) {
                    $rrt += $script:Constants.ProvisionArchive
                }
                elseif (-not $ArchiveEnabled -and ($rrt -band $script:Constants.ProvisionArchive)) {
                    $rrt -= $script:Constants.ProvisionArchive
                }
            }
            
            $desired.msExchRemoteRecipientType = $rrt
            $desired.msExchRecipientTypeDetails = $script:Constants.RecipientTypeDetails[$Type]
            
            $displayType = Get-RecipientDisplayType -Type $Type -ACLableSyncedObjectEnabled:($ACLableSyncedObjectEnabled -eq $true)
            $desired.msExchRecipientDisplayType = $displayType
        }
        
        # Handle archive toggle without type change
        if (-not $Type -and $PSBoundParameters.ContainsKey('ArchiveEnabled')) {
            $currentRRT = $user.msExchRemoteRecipientType
            if ($ArchiveEnabled -and -not ($currentRRT -band $script:Constants.ProvisionArchive)) {
                $desired.msExchRemoteRecipientType = $currentRRT + $script:Constants.ProvisionArchive
            }
            elseif (-not $ArchiveEnabled -and ($currentRRT -band $script:Constants.ProvisionArchive)) {
                $desired.msExchRemoteRecipientType = $currentRRT - $script:Constants.ProvisionArchive
            }
        }
        
        # Handle ACLable toggle
        if ($PSBoundParameters.ContainsKey('ACLableSyncedObjectEnabled')) {
            if ($ACLableSyncedObjectEnabled) {
                $desired.msExchRecipientDisplayType = $script:Constants.DisplayTypes.ACLable
            }
            else {
                $desired.msExchRecipientDisplayType = $script:Constants.DisplayTypes.NonACLable
            }
        }
        
        # Handle email addresses
        if ($PrimarySmtpAddress -or $AddEmailAddresses -or $RemoveEmailAddresses) {
            $proxyAddresses = Invoke-EmailNormalization `
                -PrimarySmtpAddress $PrimarySmtpAddress `
                -ExistingProxyAddresses $user.proxyAddresses `
                -AddEmailAddresses $AddEmailAddresses `
                -RemoveEmailAddresses $RemoveEmailAddresses `
                -OldPrimarySmtpAddress $user.mail
            
            $desired.proxyAddresses = $proxyAddresses
            
            if ($PrimarySmtpAddress) {
                $desired.mail = $PrimarySmtpAddress
            }
        }
        
        # Handle routing address
        if ($RemoteRoutingAddress) {
            Test-RoutingAddress -Address $RemoteRoutingAddress -PrimarySmtpAddress $user.mail
            $desired.targetAddress = "smtp:$RemoteRoutingAddress"
        }
        
        # Handle custom attributes
        for ($i = 1; $i -le 15; $i++) {
            $paramName = "CustomAttribute$i"
            if ($PSBoundParameters.ContainsKey($paramName)) {
                $desired["extensionAttribute$i"] = $PSBoundParameters[$paramName]
            }
        }
        
        if ($desired.Count -eq 0) {
            Write-Warning "No changes specified"
            return
        }
        
        # Validate allowlist
        Test-AttributeAllowlist -Attributes $desired
        
        # Idempotency check
        $check = Invoke-IdempotencyCheck -Identity $Identity -Desired $desired
        
        if (-not $check.HasChanges) {
            Write-ExchangeAudit `
                -Identity $Identity `
                -Before $check.Before `
                -After @{} `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'NoChange'
            
            return [PSCustomObject]@{
                Identity = $Identity
                Operation = 'Set'
                Status = 'NoChange'
                Delta = @{}
            }
        }
        
        # Apply changes
        if ($PSCmdlet.ShouldProcess($Identity, "Modify remote mailbox")) {
            Set-ADUser -Identity $Identity -Replace $check.Delta
            
            Write-ExchangeAudit `
                -Identity $Identity `
                -Before $check.Before `
                -After $check.Delta `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'Success'
            
            [PSCustomObject]@{
                Identity = $Identity
                Operation = 'Set'
                Status = 'Changed'
                Delta = $check.Delta
            }
        }
    }
}

function Disable-RemoteMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string]$Identity,
        
        [switch]$KeepEmailAddresses
    )
    
    begin {
        Assert-EmulationMode
    }
    
    process {
        try {
            $user = Get-ADUser -Identity $Identity -Properties `
                msExchRemoteRecipientType, msExchRecipientDisplayType, msExchRecipientTypeDetails, `
                proxyAddresses, mail, targetAddress, mailNickname `
                -ErrorAction Stop
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        
        $before = @{
            msExchRemoteRecipientType = $user.msExchRemoteRecipientType
            msExchRecipientDisplayType = $user.msExchRecipientDisplayType
            msExchRecipientTypeDetails = $user.msExchRecipientTypeDetails
        }
        
        $clearAttrs = @(
            'msExchRemoteRecipientType',
            'msExchRecipientDisplayType',
            'msExchRecipientTypeDetails',
            'targetAddress',
            'mailNickname'
        )
        
        if (-not $KeepEmailAddresses) {
            $clearAttrs += 'proxyAddresses'
            $clearAttrs += 'mail'
        }
        
        if ($PSCmdlet.ShouldProcess($Identity, "Disable remote mailbox")) {
            Set-ADUser -Identity $Identity -Clear $clearAttrs
            
            Write-ExchangeAudit `
                -Identity $Identity `
                -Before $before `
                -After @{} `
                -Cmdlet $MyInvocation.MyCommand.Name `
                -Parameters $PSBoundParameters `
                -Status 'Success'
            
            [PSCustomObject]@{
                Identity = $Identity
                Operation = 'Disable'
                Status = 'Changed'
                Delta = @{ Cleared = $clearAttrs }
            }
        }
    }
}

function Get-RemoteMailbox {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Position = 0, ValueFromPipeline)]
        [string]$Identity,
        
        [string]$Filter,
        
        [int]$ResultSize = 1000
    )
    
    begin {
        Assert-EmulationMode
    }
    
    process {
        $properties = @(
            'msExchRemoteRecipientType',
            'msExchRecipientDisplayType',
            'msExchRecipientTypeDetails',
            'targetAddress',
            'proxyAddresses',
            'mailNickname',
            'mail'
        )
        
        if ($Identity) {
            $user = Get-ADUser -Identity $Identity -Properties $properties -ErrorAction Stop
            
            [PSCustomObject]@{
                Identity = $user.UserPrincipalName
                Name = $user.Name
                RemoteRecipientType = $user.msExchRemoteRecipientType
                RecipientDisplayType = $user.msExchRecipientDisplayType
                RecipientTypeDetails = $user.msExchRecipientTypeDetails
                PrimarySmtpAddress = $user.mail
                RemoteRoutingAddress = $user.targetAddress
                EmailAddresses = $user.proxyAddresses
                Alias = $user.mailNickname
            }
        }
        else {
            $ldapFilter = if ($Filter) {
                $Filter
            } else {
                "(msExchRemoteRecipientType=*)"
            }
            
            Get-ADUser -LDAPFilter $ldapFilter -Properties $properties -ResultSetSize $ResultSize | 
                ForEach-Object {
                    [PSCustomObject]@{
                        Identity = $_.UserPrincipalName
                        Name = $_.Name
                        RemoteRecipientType = $_.msExchRemoteRecipientType
                        RecipientDisplayType = $_.msExchRecipientDisplayType
                        RecipientTypeDetails = $_.msExchRecipientTypeDetails
                        PrimarySmtpAddress = $_.mail
                        RemoteRoutingAddress = $_.targetAddress
                        EmailAddresses = $_.proxyAddresses
                        Alias = $_.mailNickname
                    }
                }
        }
    }
}

# ============================================================================
# MODULE EXPORT
# ============================================================================

Export-ModuleMember -Function @(
    'New-RemoteMailbox',
    'Enable-RemoteMailbox',
    'Set-RemoteMailbox',
    'Disable-RemoteMailbox',
    'Get-RemoteMailbox'
)